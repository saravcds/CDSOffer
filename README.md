# CDS Asia – Offer Letter Generator

## Start the app (Windows)
1. Install Python 3.10+ (python.org, tick "Add to PATH").
2. Double-click `run_app.bat`. The first run installs packages, then the app opens in the browser.

PDFs are made with MS Word (via docx2pdf). If Word is not installed, install LibreOffice, or untick "Also create PDF".

## Tabs
- **New offer**: one employee. Shows the salary breakup before generating.
- **Bulk offers (Excel)**: upload a sheet with `Employee Name, Annual CTC, Designation, Address`
  (optional `Salutation`, `Date of Joining`). Edit cells if needed, then generate and download a zip.
- **Register / re-issue**: every issued letter. Re-issue corrects a letter and keeps the same ID.
- **Salary calculator**: quick breakup for any CTC.

## Rules used
| Component | Formula |
|---|---|
| Basic | CTC × 50% |
| HRA | Basic × 40% |
| PF Employer | 21,600 / year (fixed) |
| Gratuity | Basic × 15 / 26 / 12 |
| Professional Tax | 2,400 / year |
| Additional Allowance | CTC − (Basic + HRA + PF + Gratuity + PT) |
| Monthly | Yearly ÷ 12, rounded per row |

To change a rule, edit the constants at the top of `offer_engine.py`.

## Document IDs
Format `CDSA/HR/APPT/<year>/<number>`, starting at 101 each year.
The next number always comes from `data/offer_register.xlsx`. Do not delete or edit that file by hand.

## Files
- `app.py` – the HR app
- `offer_engine.py` – calculation + letter filling (shared by app and batch)
- `batch_generate.py` – command-line batch run (used for the first cycle)
- `templates/CDS Appt Letter.docx` – the letter template
- `data/offer_register.xlsx` – register of all issued IDs
- `output/<yyyy-mm>/Word|PDF` – letters made by the app

## Deploy on Streamlit Community Cloud
1. Push this folder to GitHub (`saravcds/CDSOffer`). `.gitignore` keeps `data/` and `output/` out, so no salary data goes into the code repo.
2. Create a second **private** repo `saravcds/CDSOffer-data` and upload `data/offer_register.xlsx` to it (this carries the cycle-1 IDs 101–152, so the next ID is 153).
3. Create a GitHub fine-grained token with access to **only** `CDSOffer-data`, permission *Contents: Read and write*.
4. On share.streamlit.io: New app → repo `saravcds/CDSOffer`, branch `main`, file `app.py`.
5. In the app's Settings → Secrets, paste the contents of `.streamlit/secrets.toml.example` with real values.
6. Settings → Sharing: restrict viewers to HR email addresses.

`packages.txt` installs LibreOffice on the server so PDFs work there.
