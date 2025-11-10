import os
import re
import io
import time
import json
import requests
import streamlit as st
import mammoth
from datetime import datetime

# ---------- CONFIG ----------
CLOUDCONVERT_API_KEY = st.secrets["CLOUDCONVERT_API_KEY"]  # <-- set this in your env once
CLOUDCONVERT_API = "https://api.cloudconvert.com/v2/jobs"

st.set_page_config(page_title="💌 Zalando Email Generator")
st.title("💌 Gemini Email Generator")

user_name = st.text_input("Your name for sign-off (e.g., Natalia Burchard)")
uploaded_file = st.file_uploader("Upload .doc Jira Ticket", type=["doc"])

# ---------- CloudConvert helpers ----------
def cc_create_job():
    """
    Create a job with 3 tasks:
    1) import/upload  -> we'll upload the file to the URL they return
    2) convert        -> doc -> docx
    3) export/url     -> we download the converted file from a signed URL
    """
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
    headers = {"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}",
               "Content-Type": "application/json"}
    r = requests.post(CLOUDCONVERT_API, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    return r.json()["data"]

def cc_upload_to_signed_url(job_data, file_bytes, filename):
    # Find the import task info with the signed URL
    import_task = next(t for t in job_data["tasks"] if t["name"] == "import-my-file")
    upload_url = import_task["result"]["form"]["url"]
    form_params = import_task["result"]["form"]["parameters"]

    files = {"file": (filename, file_bytes)}
    r = requests.post(upload_url, data=form_params, files=files)
    # CloudConvert returns 204 No Content on success
    if r.status_code not in (200, 201, 204):
        raise RuntimeError(f"Upload failed: {r.status_code} {r.text}")

def cc_poll_until_finished(job_id, timeout_s=120, poll_every_s=2):
    headers = {"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}"}
    t0 = time.time()
    while True:
        r = requests.get(f"{CLOUDCONVERT_API}/{job_id}", headers=headers)
        r.raise_for_status()
        data = r.json()["data"]
        status = data["status"]
        if status == "finished":
            return data
        if status == "error":
            # Find task with error to display details
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
    for p in paragraphs:
        m = re.search(r"Supplier:\s*(.*)", p, flags=re.I)
        if m:
            return m.group(1).strip()
    return "[Supplier]"

def extract_invoice_number(paragraphs):
    for p in paragraphs:
        m = re.search(r"Supplier Invoice Number:\s*(.*)", p, flags=re.I)
        if m:
            return m.group(1).strip()
    return "[Invoice Number]"

def extract_table(paragraphs):
    table_rows = []
    for p in paragraphs:
        # crude but effective for Jira dumps: lines with commas & digits look like row dumps
        if "," in p and any(c.isdigit() for c in p):
            table_rows.append(p)
    return "\n".join(table_rows) if table_rows else "[Table information]"

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

Please note, as per § 23 of the German Kreislaufwirtschaftsgesetz/Circular Economy Act and similar legislation in other countries, we as a distributor are obliged to ensure that the quality of the articles we distribute is maintained and that they do not become waste (“duty of care”).

In order to follow our sustainable and eco-friendly approach and by the German Circular Economy Act, we are unable to proceed with destruction and can offer a return or internal processing. Please note, although the law is a German one we do apply it to all our warehouses across Europe.

Please confirm how you would like to proceed within the next 3 working days.
Should we not hear from you until then, we will assume your tacit consent that we may proceed by processing the articles internally.

Therefore, if you do not wish to accept this and would like to take back the articles for further processing on your side, please reach out to us within the deadline set.
If you have further questions, please do not hesitate to reach out.

Thank you and kind regards,
{name}"""

# ---------- MAIN ----------
if uploaded_file and user_name:
    try:
        if not CLOUDCONVERT_API_KEY:
            st.error("CLOUDCONVERT_API_KEY is not set in the environment.")
            st.stop()

        with st.status("Processing file…", expanded=False) as status:
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

        # simple routing: use filename to decide template
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
