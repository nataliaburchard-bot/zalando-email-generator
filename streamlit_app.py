import os
import re
import io
import time
import json
import requests
import streamlit as st
import mammoth

# ---------- CONFIG ----------
CLOUDCONVERT_API = "https://api.cloudconvert.com/v2"

st.set_page_config(page_title="💌 Zalando Email Generator")
st.title("💌 Zalando Email Generator")

# --- UI ---
user_name = st.text_input("Your name for sign-off (e.g., Natalia)")
uploaded_file = st.file_uploader("Upload .doc Jira Ticket", type=["doc"])

# ---------- API KEY HANDLING ----------
def get_api_key() -> str | None:
    """
    Prefer Streamlit secrets; fall back to env var.
    Also strip stray quotes/spaces that break auth.
    """
    key = None
    try:
        # secrets if available
        key = st.secrets.get("CLOUDCONVERT_API_KEY", None)
    except Exception:
        pass
    if not key:
        key = os.environ.get("CLOUDCONVERT_API_KEY")
    if key:
        key = str(key).strip().strip('"').strip("'")
    return key

API_KEY = get_api_key()

def auth_headers():
    # Use both Authorization and X-API-KEY just in case
    return {
        "Authorization": f"Bearer {API_KEY}",
        "X-API-KEY": API_KEY,
        "Content-Type": "application/json"
    }

def mask(k: str) -> str:
    if not k:
        return "—"
    if len(k) <= 8:
        return "*" * len(k)
    return f"{k[:4]}***{k[-4:]}"

def verify_api_key() -> None:
    """
    Call /users/me to confirm the key is valid *before* creating jobs.
    Raises with a clear message if unauthorized.
    """
    r = requests.get(f"{CLOUDCONVERT_API}/users/me", headers=auth_headers())
    if r.status_code == 401:
        raise RuntimeError(
            "CloudConvert rejected the API key (401 Unauthorized). "
            "Double-check the key is correct, active, and not copied with spaces."
        )
    r.raise_for_status()

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
    r = requests.post(f"{CLOUDCONVERT_API}/jobs", headers=auth_headers(), data=json.dumps(payload))
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
    t0 = time.time()
    while True:
        r = requests.get(f"{CLOUDCONVERT_API}/jobs/{job_id}", headers=auth_headers())
        r.raise_for_status()
        data = r.json()["data"]
        if data["status"] == "finished":
            return data
        if data["status"] == "error":
            bad = next((t for t in data["tasks"] if t["status"] == "error"), None)
            msg = bad["message"] if bad and "message" in bad else "Unknown conversion error"
            raise RuntimeError(msg)
        if time.time() - t0 > timeout_s:
            raise TimeoutError("CloudConvert job timed out")
        time.sleep(poll_every_s)

def cc_download_converted_docx(job_data) -> bytes:
    export_task = next(t for t in job_data["tasks"] if t["name"] == "export-my-file")
    file_url = export_task["result"]["files"][0]["url"]
    r = requests.get(file_url)
    r.raise_for_status()
    return r.content

# ---------- Parsing helpers ----------
def extract_text_from_docx_bytes(docx_bytes: bytes):
    with io.BytesIO(docx_bytes) as memf:
        result = mammoth.extract_raw_text(memf)
        text = result.value
    return [line.strip() for line in text.split("\n") if line.strip()]

def extract_supplier(paragraphs):
    import re
    for p in paragraphs:
        m = re.search(r"Supplier:\s*(.*)", p, flags=re.I)
        if m:
            return m.group(1).strip()
    return "[Supplier]"

def extract_invoice_number(paragraphs):
    import re
    for p in paragraphs:
        m = re.search(r"Supplier Invoice Number:\s*(.*)", p, flags=re.I)
        if m:
            return m.group(1).strip()
    return "[Invoice Number]"

def extract_table(paragraphs):
    rows = [p for p in paragraphs if ("," in p and any(c.isdigit() for c in p))]
    return "\n".join(rows) if rows else "[Table information]"

def generate_price_variance_email(supplier, invoice, table, name):
    return f"""Dear {supplier},

I hope this email finds you well.
Your invoice {invoice} is currently in clarification due to a price discrepancy.

{table}

Please get back to us within the next 3 working days, otherwise we will have to deduct the difference automatically via debit note.
If you have further questions, please do not hesitate to reach out.

Thank you and kind regards,
{name}"""

def generate_article_not_ordered_email(supplier, sn_info, article_info, number_info, name):
    return f"""Dear {supplier},

I hope this email finds you well.
We have (an) item/s in quarantine storage as it looks like we have not ordered it and therefore we can not receive it/them. The articles have been delivered with {sn_info}.

When items cannot be directly received due to specific issues they are sidelined and stored in our quarantine storage area (= a separate area in our warehouse). This additional clarification process is causing capacity losses and unforeseen costs.

{article_info}

We have received {number_info} units of this style.

We have two options:

1. Return (please provide the address, return label, and return authorization number).
2. You can agree that we shall process the goods internally at Zalando's own discretion and thereby relinquish any and all rights that may have been reserved in the items (items will be processed further internally and then sold in bulk).

Please confirm how you would like to proceed within the next 3 working days.
Should we not hear from you until then, we will assume your tacit consent that we may proceed by processing the articles internally.

If you have further questions, please do not hesitate to reach out.

Thank you and kind regards,
{name}"""

# ---------- MAIN ----------
# Show what key the app sees (masked) so you can confirm it's actually loaded.
st.caption(f"CloudConvert API key detected: **{mask(API_KEY)}**")

if uploaded_file and user_name:
    try:
        if not API_KEY:
            st.error("No CloudConvert API key found. Add it to st.secrets['CLOUDCONVERT_API_KEY'] or set the CLOUDCONVERT_API_KEY env var.")
            st.stop()

        with st.status("Processing file…", expanded=False) as status:
            status.update(label="Verifying API key")
            verify_api_key()

            status.update(label="Creating conversion job")
            job = cc_create_job()

            status.update(label="Uploading .doc to CloudConvert")
            file_bytes = uploaded_file.read()
            cc_upload_to_signed_url(job, file_bytes, uploaded_file.name)

            status.update(label="Converting to .docx")
            finished_job = cc_poll_until_finished(job["id"])

            status.update(label="Downloading converted .docx")
            docx_bytes = cc_download_converted_docx(finished_job)

        paragraphs = extract_text_from_docx_bytes(docx_bytes)
        supplier = extract_supplier(paragraphs)
        invoice = extract_invoice_number(paragraphs)
        table = extract_table(paragraphs)

        if "price" in uploaded_file.name.lower():
            email_body = generate_price_variance_email(supplier, invoice, table, user_name)
        else:
            sn_info = "[SN Info from .doc]"
            article_info = table
            number_info = "[Number/size breakdown]"
            email_body = generate_article_not_ordered_email(supplier, sn_info, article_info, number_info, user_name)

        st.success("✅ File processed successfully!")
        st.markdown("**📧 Email Preview**")
        st.text_area("", email_body, height=500)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
