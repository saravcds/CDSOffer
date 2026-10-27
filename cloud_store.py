"""
Keeps the offer register (the log of every issued document ID) safe on Streamlit Community Cloud.

The cloud server's disk is wiped on every restart, so the register is stored in the
private GitHub repo saravcds/CDSOffer-data and pulled/pushed through the GitHub API.

On Streamlit Cloud the app REFUSES to issue IDs unless this works, so numbering can
never silently restart at 101. On a PC with no secrets, the local file is used.

Secrets (Streamlit Cloud > App > Settings > Secrets):
    app_password = "..."
    [github]
    token  = "github_pat_..."          # fine-grained, CDSOffer-data only, Contents read/write
    repo   = "saravcds/CDSOffer-data"
    path   = "offer_register.xlsx"
    branch = "main"
(A top-level GITHUB_TOKEN / github_token is also accepted.)
"""
import base64
import os
from pathlib import Path

import requests
import streamlit as st

API = "https://api.github.com"
DEFAULTS = {"repo": "saravcds/CDSOffer-data", "path": "offer_register.xlsx", "branch": "main"}


def on_streamlit_cloud() -> bool:
    return Path("/mount/src").exists() or os.environ.get("HOSTNAME", "") == "streamlit"


def _secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def _cfg():
    g = _secret("github") or {}
    token = (g.get("token") if hasattr(g, "get") else None) or _secret("GITHUB_TOKEN") or _secret("github_token")
    if not token:
        return None
    get = (lambda k: (g.get(k) if hasattr(g, "get") else None) or DEFAULTS[k])
    return {"token": str(token).strip(), "repo": str(get("repo")).strip(),
            "path": str(get("path")).strip(), "branch": str(get("branch")).strip()}


def enabled() -> bool:
    return _cfg() is not None


def describe() -> str:
    c = _cfg()
    return f"{c['repo']}/{c['path']} ({c['branch']})" if c else "local file"


def _headers(c):
    return {"Authorization": f"Bearer {c['token']}", "Accept": "application/vnd.github+json"}


class RegisterUnavailable(RuntimeError):
    pass


def pull(local: Path) -> str:
    """Download the register to `local`. Raises RegisterUnavailable with a plain reason if it can't."""
    c = _cfg()
    if not c:
        if on_streamlit_cloud():
            raise RegisterUnavailable(
                "GitHub settings are missing from the app's Secrets, so the register of issued IDs "
                "cannot be loaded. Add the [github] section (token, repo, path, branch) in "
                "Settings → Secrets, then reboot the app.")
        return "local"
    url = f"{API}/repos/{c['repo']}/contents/{c['path']}"
    try:
        r = requests.get(url, headers=_headers(c), params={"ref": c["branch"]}, timeout=20)
    except requests.RequestException as ex:
        raise RegisterUnavailable(f"Could not reach GitHub: {ex}")
    if r.status_code == 401:
        raise RegisterUnavailable("GitHub rejected the token (expired or mistyped). Create a new token and update Secrets.")
    if r.status_code == 404:
        repo_ok = requests.get(f"{API}/repos/{c['repo']}", headers=_headers(c), timeout=20).status_code == 200
        if not repo_ok:
            raise RegisterUnavailable(f"The token cannot see the repo '{c['repo']}'. Give the token access to "
                                      "that repo (Contents: Read and write) or fix the repo name in Secrets.")
        raise RegisterUnavailable(f"'{c['path']}' was not found on branch '{c['branch']}' of {c['repo']}. "
                                  "Check the path/branch in Secrets.")
    if r.status_code == 403:
        raise RegisterUnavailable("The token is not allowed to read the repo contents. Give it Contents: Read and write.")
    r.raise_for_status()
    data = r.json()
    if data.get("content"):
        raw = base64.b64decode(data["content"])
    else:  # files over 1 MB come without inline content
        raw = requests.get(data["download_url"], headers=_headers(c), timeout=30).content
    local.parent.mkdir(parents=True, exist_ok=True)
    local.write_bytes(raw)
    return "pulled"


def push(local: Path, message: str):
    c = _cfg()
    if not c:
        return
    _push_cfg(c, local, message)


def _push_cfg(c, local: Path, message: str):
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
    if r.status_code == 403:
        raise RuntimeError("Token can read but not write. Give it Contents: Read and write.")
    r.raise_for_status()


def check_password() -> bool:
    pw = _secret("app_password")
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


# ---------------------------------------------------------------- other files (e.g. HR signature)
def pull_file(repo_path: str, local: Path) -> bool:
    """Download an optional file from the data repo. Returns False if it isn't there."""
    c = _cfg()
    if not c:
        return local.exists()
    r = requests.get(f"{API}/repos/{c['repo']}/contents/{repo_path}", headers=_headers(c),
                     params={"ref": c["branch"]}, timeout=20)
    if r.status_code == 404:
        return False
    r.raise_for_status()
    data = r.json()
    raw = base64.b64decode(data["content"]) if data.get("content") else \
        requests.get(data["download_url"], headers=_headers(c), timeout=30).content
    local.parent.mkdir(parents=True, exist_ok=True)
    local.write_bytes(raw)
    return True


def push_file(local: Path, repo_path: str, message: str):
    c = _cfg()
    if not c:
        return
    saved = c["path"]
    try:
        c["path"] = repo_path
        _push_cfg(c, local, message)
    finally:
        c["path"] = saved
