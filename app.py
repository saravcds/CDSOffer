"""
CDS Asia - Offer Letter Generator (HR app)

Run:  streamlit run app.py
"""
import io
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import importlib

import cloud_store
import offer_engine

# Streamlit Cloud keeps imported helper files in memory after an update; reload them so a
# new version of cloud_store.py / offer_engine.py is always used without a manual reboot.
importlib.reload(cloud_store)
importlib.reload(offer_engine)

from offer_engine import (DOC_ID_PREFIX, ROW_ORDER, Employee, Register, calc_breakup, clean_address,
                          clean_designation, clean_name, docx_bytes_to_pdf, docx_to_pdf_many,
                          employee_from_row, fill_letter, inr, make_doc_id, parse_ctc,
                          read_employees_xlsx, register_entry, safe_filename, salutation_name)

HERE = Path(__file__).parent
DEFAULT_TEMPLATE = HERE / "templates" / "CDS Appt Letter.docx"
REGISTER_PATH = HERE / "data" / "offer_register.xlsx"
OUTPUT_DIR = HERE / "output"

st.set_page_config(page_title="CDS Offer Letters", page_icon="📄", layout="wide")
if not cloud_store.check_password():
    st.stop()

# The register is the permanent log of every issued ID. On Streamlit Cloud it lives in the
# private GitHub repo; pull the latest copy on every run and never fall back to a blank file.
try:
    cloud_store.pull(REGISTER_PATH)
except cloud_store.RegisterUnavailable as ex:
    st.error("**Offer register not loaded — no document IDs can be issued.**\n\n" + str(ex))
    st.stop()
except Exception as ex:
    st.error(f"Could not load the offer register: {ex}. Not safe to issue IDs.")
    st.stop()
reg = Register(REGISTER_PATH)

# HR signature: kept in the private data repo (never in the public code repo)
SIGNATURE_PATH = HERE / "data" / "signature.png"
SIGNATURE_REPO_PATH = "signature.png"
if not st.session_state.get("sig_loaded"):  # retry each run until it loads
    try:
        st.session_state["sig_loaded"] = cloud_store.pull_file(SIGNATURE_REPO_PATH, SIGNATURE_PATH)
    except Exception as ex:
        st.session_state["sig_loaded"] = SIGNATURE_PATH.exists()
        st.warning(f"Could not load the saved HR signature: {ex}")


def save_register(entries):
    reg.add(entries)
    try:
        cloud_store.push(REGISTER_PATH, f"Offers: {entries[0]['Doc ID']} .. {entries[-1]['Doc ID']}")
    except Exception as ex:
        st.error(f"Letters made, but the register could not be saved to GitHub: {ex}. "
                 "Download the register below and send it to the admin.")

# ---------------------------------------------------------------- sidebar
LOGO = HERE / "CDSlogo.png"

# widen the sidebar a little so the 250 px logo fits without being scaled down
st.markdown("<style>section[data-testid='stSidebar']{width:330px !important;min-width:330px !important;}</style>",
            unsafe_allow_html=True)

with st.sidebar:
    if LOGO.exists():
        st.image(str(LOGO), width=250)  # fixed width, height follows aspect ratio
    st.header("Settings")
    letter_date = st.date_input("Letter date", value=date.today(), format="DD/MM/YYYY")
    default_doj = st.date_input("Default joining date", value=letter_date, format="DD/MM/YYYY",
                                help="Used for 'effective', 'w.e.f.' and Date of Joining unless set per employee")
    signatory = st.text_input("HR signatory", "Khushboo Kaur")
    signatory_title = st.text_input("Signatory title", "Head – HR")

    st.subheader("HR signature")
    add_sig = st.checkbox("Add signature to letters", value=True)
    if SIGNATURE_PATH.exists():
        st.image(str(SIGNATURE_PATH), width=180, caption="Current signature")
    else:
        st.warning("No signature saved yet. Upload one below.")
    up_sig = st.file_uploader("Upload new signature (PNG with transparent or white background)",
                              type=["png", "jpg", "jpeg"], key="sig_upload")
    if up_sig is not None:
        st.image(up_sig.getvalue(), width=180, caption="New signature (preview)")
        if st.button("Save as signature for all letters"):
            from PIL import Image
            im = Image.open(io.BytesIO(up_sig.getvalue()))
            im.save(SIGNATURE_PATH, "PNG")
            try:
                cloud_store.push_file(SIGNATURE_PATH, SIGNATURE_REPO_PATH, "Update HR signature")
                st.success("Signature saved. It will be used on all new letters.")
            except Exception as ex:
                st.error(f"Signature saved for this session only; could not save to GitHub: {ex}")
    signature = SIGNATURE_PATH.read_bytes() if (add_sig and SIGNATURE_PATH.exists()) else None
    make_pdf = st.checkbox("Also create PDF", value=True,
                           help="Needs MS Word (docx2pdf) or LibreOffice on this computer")
    up_tpl = st.file_uploader("Use a different template (.docx)", type=["docx"])
    template = DEFAULT_TEMPLATE
    if up_tpl:
        template = io.BytesIO(up_tpl.getvalue())
    year = letter_date.year
    next_seq = reg.next_seq(year)
    st.info(f"Next document ID:\n\n**{make_doc_id(year, next_seq)}**")
    _rows = reg.rows()
    st.caption(f"Register: {'GitHub · ' + cloud_store.describe() if cloud_store.enabled() else 'local file'}"
               f" · {len(_rows)} IDs logged" + (f" · last {_rows[-1]['Doc ID']}" if _rows else ""))


def tpl():
    if isinstance(template, io.BytesIO):
        template.seek(0)
        return io.BytesIO(template.getvalue())
    return template


def breakup_df(ctc: float) -> pd.DataFrame:
    b = calc_breakup(ctc)
    rows = []
    for k, label in ROW_ORDER:
        if k == "pf":
            rows.append({"Component": "Statutory Benefits", "Yearly": "", "Monthly": ""})
        rows.append({"Component": label, "Yearly": inr(b["yearly"][k]), "Monthly": inr(b["monthly"][k])})
    return pd.DataFrame(rows)


def save_outputs(emp: Employee, doc_id: str, doj: date, pdf: bool):
    docx_bytes = fill_letter(tpl(), emp, doc_id, letter_date, doj, signatory, signatory_title, signature)
    fname = safe_filename(doc_id, emp.name)
    folder = OUTPUT_DIR / f"{letter_date:%Y-%m}"
    (folder / "Word").mkdir(parents=True, exist_ok=True)
    (folder / "Word" / f"{fname}.docx").write_bytes(docx_bytes)
    pdf_bytes = None
    if pdf:
        try:
            pdf_bytes = docx_bytes_to_pdf(docx_bytes, fname)
            (folder / "PDF").mkdir(parents=True, exist_ok=True)
            (folder / "PDF" / f"{fname}.pdf").write_bytes(pdf_bytes)
        except Exception as ex:  # noqa
            st.warning(f"PDF not created: {ex}")
    return fname, docx_bytes, pdf_bytes


st.title("CDS Asia · Offer Letter Generator")
tab1, tab2, tab3, tab4 = st.tabs(["New offer", "Bulk offers (Excel)", "Register / re-issue", "Salary calculator"])

# ---------------------------------------------------------------- single offer
with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        name = st.text_input("Employee name *")
        sal = st.text_input("Dear ... (first and last name)", value=salutation_name(clean_name(name)) if name else "",
                            key=f"sal_{name}")
        designation = st.text_input("Designation *")
        address = st.text_area("Address (one line per row)", height=140)
        ctc_txt = st.text_input("Annual CTC *", placeholder="e.g. 1400000 or 14 LPA")
        doj = st.date_input("Joining date", value=default_doj, format="DD/MM/YYYY", key="doj_single")
    ctc = parse_ctc(ctc_txt) if ctc_txt else None
    with c2:
        st.markdown(f"**Document ID:** `{make_doc_id(year, next_seq)}`")
        if ctc:
            st.markdown(f"**CTC:** ₹ {inr(ctc)}")
            st.dataframe(breakup_df(ctc), hide_index=True, width="stretch")
            if calc_breakup(ctc)["yearly"]["additional"] < 0:
                st.error("Additional Allowance is negative. CTC is too low for this structure.")
        elif ctc_txt:
            st.error("Could not read CTC. Use a number (1400000) or '14 LPA' / '1 Cr'.")
        if address:
            st.caption("Address as it will print:")
            st.text("\n".join(clean_address(address)))

    if st.button("Generate offer letter", type="primary", disabled=not (name and designation and ctc)):
        emp = employee_from_row(name, ctc, designation, address, salutation=sal or None, joining=doj)
        seq = reg.next_seq(year)  # re-read in case someone else just issued one
        doc_id = make_doc_id(year, seq)
        fname, dx, pdf = save_outputs(emp, doc_id, doj, make_pdf)
        save_register([register_entry(emp, doc_id, year, seq, letter_date, doj, fname)])
        st.success(f"Created {doc_id} for {emp.name}")
        for f in emp.flags:
            st.warning(f)
        d1, d2 = st.columns(2)
        d1.download_button("Download Word", dx, f"{fname}.docx")
        if pdf:
            d2.download_button("Download PDF", pdf, f"{fname}.pdf", mime="application/pdf")

# ---------------------------------------------------------------- bulk
with tab2:
    st.write("Upload an Excel file with columns **Employee Name, Annual CTC, Designation, Address** "
             "(optional: *Salutation*, *Date of Joining*). Sheet name 'CTC details' or the last sheet.")
    up = st.file_uploader("Employee Excel", type=["xlsx"], key="bulk")
    if up:
        try:
            emps = read_employees_xlsx(io.BytesIO(up.getvalue()))
        except Exception as ex:
            st.error(str(ex))
            emps = []
        if emps:
            start = st.number_input("Start number", min_value=1, value=reg.next_seq(year), step=1,
                                    help="Taken from the register. Change only if you know why.")
            used = {r["Doc ID"] for r in reg.rows()}
            df = pd.DataFrame([{
                "Generate": e.ctc is not None,
                "Doc ID": make_doc_id(year, int(start) + i),
                "Employee Name": e.name,
                "Dear ...": e.salutation,
                "Designation": e.designation,
                "Address (use | for new line)": " | ".join(e.address_lines),
                "Annual CTC": e.ctc,
                "Joining Date": e.joining_date or default_doj,
                "Flags": "; ".join(e.flags),
            } for i, e in enumerate(emps)])
            st.caption(f"{len(df)} employees · {sum(bool(f) for f in df['Flags'])} with flags. "
                       "You can edit cells before generating.")
            edited = st.data_editor(df, hide_index=True, width="stretch", num_rows="fixed",
                                    disabled=["Doc ID", "Flags"],
                                    column_config={"Annual CTC": st.column_config.NumberColumn(format="%d"),
                                                   "Joining Date": st.column_config.DateColumn(format="DD/MM/YYYY")})
            sel = edited[edited["Generate"]]
            clash = [d for d in sel["Doc ID"] if d in used]
            if clash:
                st.error(f"These IDs are already in the register: {', '.join(clash[:5])}... Change the start number.")
            if st.button(f"Generate {len(sel)} letters", type="primary", disabled=bool(clash) or sel.empty):
                bar = st.progress(0.0)
                zbuf = io.BytesIO()
                entries, docx_paths = [], []
                folder = OUTPUT_DIR / f"{letter_date:%Y-%m}"
                (folder / "Word").mkdir(parents=True, exist_ok=True)
                # number sequentially over selected rows only
                with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                    for n, (_, r) in enumerate(sel.iterrows()):
                        seq = int(start) + n
                        doc_id = make_doc_id(year, seq)
                        doj_r = r["Joining Date"] if isinstance(r["Joining Date"], date) else default_doj
                        emp = employee_from_row(r["Employee Name"], r["Annual CTC"], r["Designation"],
                                                str(r["Address (use | for new line)"] or "").replace("|", "\n"),
                                                salutation=r["Dear ..."] or None, joining=doj_r)
                        dx = fill_letter(tpl(), emp, doc_id, letter_date, doj_r, signatory, signatory_title, signature)
                        fname = safe_filename(doc_id, emp.name)
                        p = folder / "Word" / f"{fname}.docx"
                        p.write_bytes(dx)
                        docx_paths.append(p)
                        z.writestr(f"Word/{fname}.docx", dx)
                        entries.append(register_entry(emp, doc_id, year, seq, letter_date, doj_r, fname,
                                                      "Issued - check flags" if emp.flags else "Issued"))
                        bar.progress((n + 1) / len(sel) * (0.5 if make_pdf else 1))
                    if make_pdf:
                        try:
                            with st.spinner("Creating PDFs..."):
                                pdfs = docx_to_pdf_many(docx_paths, folder / "PDF")
                            for pf in pdfs:
                                z.write(pf, f"PDF/{pf.name}")
                        except Exception as ex:
                            st.warning(f"PDFs not created: {ex}")
                    bar.progress(1.0)
                save_register(entries)
                st.success(f"Created {len(entries)} letters: {entries[0]['Doc ID']} to {entries[-1]['Doc ID']}. "
                           f"Saved in {folder}")
                st.download_button("Download all (zip)", zbuf.getvalue(),
                                   f"Offers_{letter_date:%Y%m%d}.zip", mime="application/zip")

# ---------------------------------------------------------------- register
with tab3:
    rows = reg.rows()
    if not rows:
        st.info("No offers issued yet.")
    else:
        rdf = pd.DataFrame(rows)
        q = st.text_input("Search name or ID")
        view = rdf[rdf.apply(lambda r: q.lower() in f"{r['Doc ID']} {r['Employee Name']}".lower(), axis=1)] if q else rdf
        st.dataframe(view, hide_index=True, width="stretch")
        st.download_button("Download register", REGISTER_PATH.read_bytes(), REGISTER_PATH.name)

        st.subheader("Re-issue a letter (same document ID)")
        st.caption("Use this to correct a letter, e.g. a wrong address or CTC. The ID stays the same.")
        pick = st.selectbox("Document ID", [r["Doc ID"] for r in rows])
        cur = reg.find(pick)
        with st.form("reissue"):
            r_name = st.text_input("Employee name", cur["Employee Name"])
            r_sal = st.text_input("Dear ...", salutation_name(cur["Employee Name"]))
            r_des = st.text_input("Designation", cur["Designation"])
            r_addr = st.text_area("Address (one line per row)")
            r_ctc = st.text_input("Annual CTC", str(int(cur["Annual CTC"] or 0)))
            r_doj = st.date_input("Joining date", value=cur["Joining Date"] or default_doj, format="DD/MM/YYYY")
            go = st.form_submit_button("Re-issue")
        if go:
            emp = employee_from_row(r_name, parse_ctc(r_ctc), r_des, r_addr, salutation=r_sal, joining=r_doj)
            if emp.ctc is None:
                st.error("CTC not readable")
            else:
                fname, dx, pdf = save_outputs(emp, pick, r_doj, make_pdf)
                e = register_entry(emp, pick, int(cur["Year"]), int(cur["Seq"]), letter_date, r_doj, fname, "Re-issued")
                save_register([e])
                st.success(f"Re-issued {pick}")
                c1, c2 = st.columns(2)
                c1.download_button("Download Word", dx, f"{fname}.docx")
                if pdf:
                    c2.download_button("Download PDF", pdf, f"{fname}.pdf", mime="application/pdf")

# ---------------------------------------------------------------- calculator
with tab4:
    t = st.text_input("Annual CTC", "14 LPA", key="calc")
    v = parse_ctc(t)
    if v:
        st.dataframe(breakup_df(v), hide_index=True)
        st.caption("Basic = 50% of CTC · HRA = 40% of Basic · PF 21,600 fixed · Gratuity = Basic × 15/26/12 · "
                   "PT 2,400 · Additional Allowance = balancing figure")
