"""Test Equity - Cash Application report updater (Streamlit)."""
import base64
import datetime as dt

import requests
import streamlit as st

import ca_report as C

st.set_page_config(page_title="Cash Application Report", page_icon="📊", layout="centered")
st.title("Cash Application – daily report update")

# ------------------------------------------------------------------ password gate
PW = st.secrets.get("APP_PASSWORD", "")
if PW:
    if not st.session_state.get("ok"):
        typed = st.text_input("Password", type="password")
        if typed:
            if typed == PW:
                st.session_state["ok"] = True
                st.rerun()
            else:
                st.error("Wrong password.")
        st.stop()

# ------------------------------------------------------------------ GitHub helpers
GH = st.secrets.get("GITHUB", None)


def gh_parts():
    headers = {"Authorization": f"Bearer {GH['token']}", "X-GitHub-Api-Version": "2022-11-28"}
    url = f"https://api.github.com/repos/{GH['repo']}/contents/{GH['path']}"
    params = {"ref": GH.get("branch", "main")}
    return headers, url, params


def fetch_latest():
    """Always downloads the newest copy, so two people can never work on an old version."""
    headers, url, params = gh_parts()
    meta = requests.get(url, headers=headers, params=params, timeout=60)
    if meta.status_code != 200:
        try:
            msg = meta.json().get("message")
        except Exception:  # noqa
            msg = meta.text[:100]
        return None, None, (f"GitHub replied {meta.status_code}: {msg}. "
                            "Check the token, repo name and file name in Secrets.")
    raw = requests.get(url, headers={**headers, "Accept": "application/vnd.github.raw+json"},
                       params=params, timeout=300)
    if raw.status_code != 200:
        return None, None, f"Could not download the report (GitHub replied {raw.status_code})."
    return raw.content, meta.json()["sha"], None


# ------------------------------------------------------------------ check the connection once
manual_master = None
if GH:
    if "gh_checked" not in st.session_state:
        with st.spinner("Checking the connection to GitHub..."):
            headers, url, params = gh_parts()
            chk = requests.get(url, headers=headers, params=params, timeout=60)
        if chk.status_code != 200:
            try:
                msg = chk.json().get("message")
            except Exception:  # noqa
                msg = chk.text[:100]
            st.error(f"GitHub replied {chk.status_code}: {msg}. "
                     "Check the token, repo name and file name in Secrets.")
            st.stop()
        st.session_state["gh_checked"] = True
    st.caption("Connected to GitHub. The newest report is downloaded when you click Update report.")
else:
    up = st.file_uploader("Current report (.xlsx)", type="xlsx")
    manual_master = up.getvalue() if up else None
    if manual_master is None:
        st.info("Upload the current report to start. (Secrets are not set, so GitHub is not used.)")
        st.stop()

# ------------------------------------------------------------------ EOD files
st.subheader("1. Select the agents' EOD files")
st.write("Select **both** files together (hold Ctrl and click each file).")
eods = st.file_uploader("EOD files", type="xlsx", accept_multiple_files=True)
if not eods:
    st.stop()

files = [(f.name, f.getvalue()) for f in eods]
summary, pwarn, perr = C.preview(files)

st.subheader("2. Check what was found")
for e in perr:
    st.error(e)
for w in pwarn:
    st.warning(w)
if summary:
    st.table([{"EOD sheet": a, "Agent": b, "Date": c, "Rows": n} for a, b, c, n in summary])
    st.caption("The date and agent come from inside each file. Month and Week are filled in automatically.")

go = st.button("Update report", type="primary", disabled=bool(perr) or not summary)

if go:
    st.session_state.pop("result", None)
    sha = None
    if GH:
        with st.spinner("Downloading the newest report from GitHub..."):
            master, sha, err = fetch_latest()
        if err:
            st.error(err)
            st.stop()
    else:
        master = manual_master

    with st.spinner("Updating the report (about 10-20 seconds)..."):
        res = C.process(master, files)

    if not res["ok"]:
        for e in res["errors"]:
            st.error(e)
        for w in res["warnings"]:
            st.warning(w)
        st.stop()

    for m in res["messages"]:
        st.success(m)
    for w in res["warnings"]:
        st.warning(w)
    st.session_state["result"] = res["new_bytes"]

    if GH:
        headers, url, params = gh_parts()
        with st.spinner("Saving back to GitHub. Wait for the green message..."):
            body = {
                "message": "EOD update " + dt.date.today().isoformat(),
                "content": base64.b64encode(res["new_bytes"]).decode(),
                "sha": sha,
                "branch": GH.get("branch", "main"),
            }
            put = requests.put(url, headers=headers, json=body, timeout=300)
        if put.status_code in (200, 201):
            st.success("Saved back to GitHub. You can now download the report below.")
        else:
            try:
                msg = put.json().get("message")
            except Exception:  # noqa
                msg = put.text[:100]
            st.error(f"Could NOT save to GitHub (reply {put.status_code}: {msg}). "
                     "Download the report below and keep it, and tell the owner before loading anything else.")

if st.session_state.get("result"):
    st.subheader("3. Download")
    st.download_button(
        "Download updated report",
        data=st.session_state["result"],
        file_name="Test Equity - Cash Application Specialist - " + dt.date.today().strftime("%d-%b-%Y") + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.info("Open it in Excel. Excel recalculates the Daily, Weekly and month sheets when it opens, "
            "so the first open may take a few seconds.")
