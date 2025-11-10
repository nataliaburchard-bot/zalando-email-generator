import streamlit as st
import requests
import base64
import os
import re
from docx import Document
from datetime import datetime

st.set_page_config(page_title="💌 Gemini Email Generator")

# Title
st.title("💌 Gemini Email Generator")

# Input field for user name
user_name = st.text_input("Your name for sign-off (e.g., Natalia Burchard)")

# Upload DOC or DOCX file
uploaded_file = st.file_uploader("Upload .doc or .docx Jira Ticket", type=["doc", "docx"])

# Conversion function using CloudConvert
CLOUDCONVERT_API_KEY = os.getenv("CLOUDCONVERT_API_KEY")

def convert_doc_to_docx_cloudconvert(uploaded_file):
    upload_response = requests.post(
        "https://api.cloudconvert.com/v2/import/upload",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}"}
    )
    import_id = upload_response.json()["data"]["id"]
    upload_url = upload_response.json()["data"]["url"]

    requests.put(upload_url, data=uploaded_file.getvalue())

    payload = {
        "tasks": {
            "import-my-file": {
                "operation": "import/upload"
            },
            "convert-my-file": {
                "operation": "convert",
                "input": "import-my-file",
                "input_format": "doc",
                "output_format": "docx"
            },
            "export-my-file": {
                "operation": "export/url",
                "input": "convert-my-file"
            }
        }
    }

    job_response = requests.post(
        "https://api.cloudconvert.com/v2/jobs",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}", "Content-Type": "application/json"},
        json=payload
    )

    job_id = job_response.json()["data"]["id"]

    import_task_response = requests.get(
        f"https://api.cloudconvert.com/v2/jobs/{job_id}",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}"}
    )

    tasks = import_task_response.json()["data"]["tasks"]
    export_task = next(task for task in tasks if task["name"] == "export-my-file")

    try:
        download_url = export_task["result"]["files"][0]["url"]
    except KeyError:
        st.error("❌ Could not retrieve download URL. Check CloudConvert job response.")
        st.json(export_task)
        return None

    docx_data = requests.get(download_url).content
    return docx_data

def extract_text_from_docx(docx_bytes):
    with open("temp.docx", "wb") as f:
        f.write(docx_bytes)
    doc = Document("temp.docx")
    return [p.text for p in doc.paragraphs if p.text.strip() != ""]

def extract_supplier(paragraphs):
    for p in paragraphs:
        match = re.search(r"Supplier:\s*(.*)", p)
        if match:
            return match.group(1).strip()
    return "[Supplier]"

def extract_invoice_number(paragraphs):
    for p in paragraphs:
        match = re.search(r"Supplier Invoice Number:\s*(.*)", p)
        if match:
            return match.group(1).strip()
    return "[Invoice Number]"

def extract_table(paragraphs):
    table_data = []
    for p in paragraphs:
        if "," in p and any(char.isdigit() for char in p):
            table_data.append(p)
    return "\n".join(table_data) if table_data else "[Table information]"

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

# MAIN LOGIC
if uploaded_file and user_name:
    st.info("Converting .doc to .docx using CloudConvert API...")
    converted = convert_doc_to_docx_cloudconvert(uploaded_file)

    if converted:
        st.success("✅ Conversion successful!")
        text = extract_text_from_docx(converted)
        supplier = extract_supplier(text)
        invoice = extract_invoice_number(text)
        table = extract_table(text)

        # Determine email type
        if "price" in uploaded_file.name.lower():
            email_body = generate_price_variance_email(supplier, invoice, table, user_name)
        else:
            sn_info = "[SN Info from .doc]"
            article_info = table
            number_info = "[Number/size breakdown]"
            email_body = generate_article_not_ordered_email(supplier, sn_info, article_info, number_info, user_name)

        st.markdown("**📧 Email Preview**")
        st.text_area("", email_body, height=500)

    else:
        st.error("❌ File conversion failed. Please try again.")
