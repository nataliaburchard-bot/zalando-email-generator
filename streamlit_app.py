import os
import re
import io
import time
import json
import requests
import streamlit as st
import mammoth
from datetime import datetime

# 💡 Step 1: Paste your CloudConvert key here
os.environ["CLOUDCONVERT_API_KEY"] = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiOTc1M2RkNjU3M2FlNTQ5NWU4ZTdhM2U1OWE1MWY2ZjhjNzFiNjQxODA2ZWI4NzA3MWFkNmU1NDcxNzlmZTY5ZGZkZjM3YzQ0MTVhNGExMDUiLCJpYXQiOjE3NjI3ODAxMDQuOTQ1ODQzLCJuYmYiOjE3NjI3ODAxMDQuOTQ1ODQ0LCJleHAiOjQ5MTg0NTM3MDQuOTM5MDc4LCJzdWIiOiI3MzQyODQ5NyIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIiwidXNlci5yZWFkIl19.cba1xzDx_RwGp2BGZVlcoBShNDLGbFTftpuV5zJf1hLk7jV-j2Wfin7zr-s6fiWF1wMOkpac4cMy_QNSEyI5kBLzy_RLsTmdluBNgqF6UvF5qOdE_fneHKIVH1mkUrgVENPsf9mmoXJbg7oOW9D9nmQA8wFZXaxboERZGkpeeackF66_cvOXOWt-8Yy05tMt5TBy7YtJ4zaw1BKMEwid0wotWLvxkmfiTjPun7GlPF-Jdsn0KVwdl4lEX7GNv4zPsJHBvEXdnkymiB3UmL3wACOyrsYJSO1EvO_qG93Y1VWxTM3Jb0Ynrfi_lOw3oXtywVyv0SZWwKmJgzntovky61p8LkLT-C8MGkDSSHv5zx2AIGC27Zh1qrbXDxuJeLVShf0v36j454W2sLHkQkhyp7yuNWV5risjWwQx96DHQQaW2eMJEV-mUAsxdDEp9i1KmVpW-3rB_MH6TdXmvKBKBmMjsPm88fvsApdnappYs3rxZOu3vBLBLxSbVE9BuWPbTs0NBMzoM1xkq7YrIpSkjMgyLarQMYfwyF_sq9Bi52p7P7C52LluREcrSEB4oKOT_Y6dzT7KCj4LjE8O4P8DVhuX0R6L9To-cLUlHWf5IptSY2Y2K3HdFmBq4auJJkh3A6w_bgfvJNtPTD5WTc80ur1QpvOJr-gFwoNw4QomH_g"  # ← replace with your real key

# ---------- CONFIG ----------
CLOUDCONVERT_API_KEY = os.getenv("CLOUDCONVERT_API_KEY")
CLOUDCONVERT_API = "https://api.cloudconvert.com/v2/jobs"

# ---------- STREAMLIT UI ----------
st.set_page_config(page_title="💌 Zalando Email Generator")
st.title("💌 Gemini Email Generator")

user_name = st.text_input("Your name for sign-off (e.g., Natalia Burchard)")
uploaded_file = st.file_uploader("Upload .doc Jira Ticket", type=["doc"])

# ---------- CloudConvert helpers ----------
def cc_create_job():
    payload = {
        "tasks": {
            "import-my-file": {"operation": "import/upload"},
            "convert-my-file": {
                "operation": "convert",
                "input": "import-my-file",
                "input_format": "doc",
                "output_format": "docx"
            },
            "export-my-file": {"operation": "export/url", "input": "convert-my-file"}
        }
    }
    headers = {
        "Authorization": f"Bearer {CLOUDCONVERT_API_KEY}",
        "Content-Type": "application/json"
    }
    r = requests.post(CLOUDCONVERT_API, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    return r.json()["data"]

def cc_upload_to_signed_url(job_data, file_bytes, filename):
    import_task = next(t for t in job_data["tasks"] if t["name"] == "import-my-file")
    upload_url = import_task["result"]["form"]["url"]
    form_params = import_task["result"]["form"]["parameters"]
    files = {"file": (filename, file_bytes)}
    r = requests.post(upload_url, data=form_params, files=files)
    if r.status_code not in (200, 201, 204):
        raise RuntimeError(f"Upload failed: {r.status_code} {r.text}")

def cc_poll_until_finished(job_id, timeout_s=120, poll_every_s=2):
