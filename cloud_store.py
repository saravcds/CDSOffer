"""
Keeps the offer register safe when the app runs on Streamlit Community Cloud.

The cloud server's disk is wiped on every restart, so the register is stored in a
private GitHub repo and pulled/pushed through the GitHub API.
If no [github] secrets are set, the app simply uses the local file (fine on a PC).

Secrets (Streamlit Cloud > App > Settings > Secrets):
    app_password = "..."            # login for HR
    [github]
    token  = "github_pat_..."       # fine-grained token, Contents: read & write, only the data repo
    repo   = "saravcds/CDSOffer-data"
    path   = "offer_register.xlsx"
    branch = "main"
"""
import base64
from pathlib import Path

import requests
import streamlit as st

API = "https://api.github.com"


def _cfg():
    try:
        g = st.secrets.get("github")
    except Exception:
        return None
    if not g or not g.get("token") or not g.get("repo"):
        return None
    return {"token": g["token"], "repo": g["repo"], "path": g.get("path", "offer_register.xlsx"),
            "branch": g.get("branch", "main")}


def enabled() -> bool:
    return _cfg() is not None


def _headers(c):
    return {"Authorization": f"Bearer {c['token']}", "Accept": "application/vnd.github+json"}


def pull(local: Path) -> str:
    """Download the register to `local`. Returns a status message."""
    c = _cfg()
    if not c:
        return "local"
    r = requests.get(f"{API}/repos/{c['repo']}/contents/{c['path']}", headers=_headers(c),
                     params={"ref": c["branch"]}, timeout=20)
    if r.status_code == 404:
        return "new"  # first run: register will be created and pushed
    r.raise_for_status()
    local.parent.mkdir(parents=True, exist_ok=True)
    local.write_bytes(base64.b64decode(r.json()["content"]))
    return "pulled"


def push(local: Path, message: str):
    c = _cfg()
    if not c:
        return
    url = f"{API}/repos/{c['repo']}/contents/{c['path']}"
    r = requests.get(url, headers=_headers(c), params={"ref": c["branch"]}, timeout=20)
    sha = r.json().get("sha") if r.status_code == 200 else None
    body = {"message": message, "branch": c["branch"],
            "content": base64.b64encode(local.read_bytes()).decode()}
    if sha:
        body["sha"] = sha
    r = requests.put(url, headers=_headers(c), json=body, timeout=30)
    if r.status_code == 409:
        raise RuntimeError("Register changed elsewhere. Reload the page and try again.")
    r.raise_for_status()


def check_password() -> bool:
    """Simple login when app_password is set in secrets."""
    try:
        pw = st.secrets.get("app_password")
    except Exception:
        pw = None
    if not pw:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.title("CDS Asia · Offer Letter Generator")
    entered = st.text_input("Password", type="password")
    if entered:
        if entered == pw:
            st.session_state["auth_ok"] = True
            st.rerun()
        st.error("Wrong password")
    return False
