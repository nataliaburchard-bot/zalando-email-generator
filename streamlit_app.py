import streamlit as st
import mammoth
import tempfile
import requests
import os
import re

# CloudConvert API Key
api_key = os.getenv("CLOUDCONVERT_API_KEY")

# Title
st.markdown("## 💌 Gemini Email Generator")

# User input for sign-off name
user_name = st.text_input("Your name for sign-off (e.g., Natalia Burchard)", "")

# Upload block
uploaded_file = st.file_uploader("Upload .doc or .docx Jira Ticket", type=["doc", "docx"])

def convert_doc_to_docx_cloudconvert(doc_file):
    # 1. Request a signed import URL
    signed_upload = requests.post(
        "https://api.cloudconvert.com/v2/import/upload",
        headers={"Authorization": f"Bearer {api_key}"}
    ).json()

    upload_url = signed_upload["data"]["url"]
    upload_params = signed_upload["data"]["parameters"]

    # 2. Upload file using signed URL
    files = {'file': (doc_file.name, doc_file, 'application/msword')}
    requests.post(upload_url, data=upload_params, files=files)

    import_task_name = signed_upload["data"]["id"]

    # 3. Define export task
    job_payload = {
        "tasks": {
            "convert": {
                "operation": "convert",
                "input": import_task_name,
                "input_format": "doc",
                "output_format": "docx"
            },
            "export": {
                "operation": "export/url",
                "input": "convert"
            }
        }
    }

    job_response = requests.post(
        "https://api.cloudconvert.com/v2/jobs",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json=job_payload
    ).json()

    job_id = job_response["data"]["id"]

    # 4. Poll for job completion
    while True:
        job_status = requests.get(
            f"https://api.cloudconvert.com/v2/jobs/{job_id}",
            headers={"Authorization": f"Bearer {api_key}"}
        ).json()
        if job_status["data"]["status"] == "finished":
            break

    export_url = job_status["data"]["tasks"][-1]["result"]["files"][0]["url"]
    docx_content = requests.get(export_url).content
    return docx_content

def extract_text_from_docx_bytes(docx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
        temp_docx.write(docx_bytes)
        temp_docx.flush()
        with open(temp_docx.name, "rb") as docx_file:
            result = mammoth.extract_raw_text(docx_file)
            return result.value

def detect_template(text):
    if "price" in text.lower():
        return "price_variance"
    elif "not ordered" in text.lower() or "quarantine" in text.lower():
        return "article_not_ordered"
    return "unknown"

def extract_supplier(text):
    match = re.search(r"Supplier(?: Name)?:\s*(.*)", text)
    return match.group(1).strip() if match else "[supplier]"

def extract_invoice(text):
    match = re.search(r"Supplier Invoice Number(?:\:| -)?\s*(\S+)", text)
    return match.group(1).strip() if match else "[invoice]"

def extract_table(text):
    table_lines = []
    for line in text.splitlines():
        if re.search(r"\d{6,}", line):  # line with article number
            table_lines.append(line.strip())
    return "\n".join(table_lines) if table_lines else "[table info]"

def extract_breakdown(text):
    match = re.search(r"(\d+)\s*(pcs|pieces|units)", text, re.IGNORECASE)
    return match.group(0).strip() if match else "[number/size breakdown]"

def generate_email(template_type, supplier, invoice, table, breakdown, user_name):
    if template_type == "price_variance":
        return f"""Dear {supplier},

I hope this email finds you well.
Your invoice {invoice} is currently in clarification due to a price discrepancy.

{table}

Please get back to us within the next 3 working days, otherwise we will have to deduct the difference automatically via debit note.
If you have further questions, please do not hesitate to reach out.

Thank you and kind regards,
{user_name}
"""
    elif template_type == "article_not_ordered":
        return f"""Dear {supplier},

I hope this email finds you well.
We have (an) item/s in quarantine storage as it looks like we have not ordered it and therefore we can not receive it/them. The articles have been delivered with [SN].
When items cannot be directly received due to specific issues they are sidelined and stored in our quarantine storage area (= a separate area in our warehouse). This additional clarification process is causing capacity losses and unforeseen costs.

{table}
We have received {breakdown} of this style.

We have two options:

1. Return (please provide the address, return label, and return authorization number).
2. You can agree that we shall process the goods internally at Zalando's own discretion and thereby relinquish any and all rights that may have been reserved in the items (items will be processed further internally and then sold in bulk)

Please note, as per § 23 of the German Kreislaufwirtschaftsgesetz/Circular Economy Act (Kreislaufwirtschaftsgesetz) and similar legislation in other countries, we as a distributor are obliged to ensure that the quality of the articles we distribute is maintained and that they do not become waste (“duty of care”).

In order to follow our sustainable and eco-friendly approach and by the German Circular Economy Act, we are unable to proceed with destruction and can offer a return or internal processing. Please note, although the law is a German one we do apply it to all our warehouses across Europe.

Please confirm how you would like to proceed within the next 3 working days.
Should we not hear from you until then, we will assume your tacit consent that we may proceed by processing the articles internally.
Therefore, if you do not wish to accept this and would like to take back the articles for further processing on your side, please reach out to us within the deadline set.

If you have further questions, please do not hesitate to reach out.

Thank you and kind regards,
{user_name}
"""
    else:
        return "Unable to determine email template."

# Main logic
if uploaded_file:
    st.success(f"File uploaded successfully: {uploaded_file.name}")
    with st.spinner("Converting .doc to .docx using CloudConvert API..."):
        try:
            docx_bytes = convert_doc_to_docx_cloudconvert(uploaded_file)
            st.success("✅ Conversion successful!")
            text = extract_text_from_docx_bytes(docx_bytes)
            template = detect_template(text)
            supplier = extract_supplier(text)
            invoice = extract_invoice(text)
            table = extract_table(text)
            breakdown = extract_breakdown(text)
            email_output = generate_email(template, supplier, invoice, table, breakdown, user_name)
            st.markdown("### 📧 Email Preview")
            st.code(email_output, language="markdown")
        except Exception as e:
            st.error(f"Something went wrong: {e}")
