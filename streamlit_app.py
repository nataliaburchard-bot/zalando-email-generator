import os
import io
import re
import streamlit as st
import mammoth

# Aspose SDK
import asposewordscloud
from asposewordscloud.apis.words_api import WordsApi
from asposewordscloud.models.requests import ConvertDocumentRequest


# ---------- UI ----------
st.set_page_config(page_title="💌 Zalando Email Generator")
st.title("💌 Zalando Email Generator")

user_name = st.text_input("Your name for sign-off (e.g., Natalia Burchard)")
uploaded_file = st.file_uploader("Upload .doc Jira Ticket", type=["doc"])


# ---------- Credentials ----------
def get_aspose_creds():
    """Reads credentials from Streamlit secrets or environment variables."""
    cid = None
    sec = None
    try:
        cid = st.secrets.get("ASPOSE_CLIENT_ID", None)
        sec = st.secrets.get("ASPOSE_CLIENT_SECRET", None)
    except Exception:
        pass
    cid = cid or os.getenv("ASPOSE_CLIENT_ID")
    sec = sec or os.getenv("ASPOSE_CLIENT_SECRET")
    return (str(cid).strip() if cid else None,
            str(sec).strip() if sec else None)

ASPOSE_CLIENT_ID, ASPOSE_CLIENT_SECRET = get_aspose_creds()


# ---------- Aspose converter ----------
@st.cache_resource
def get_words_api():
    if not ASPOSE_CLIENT_ID or not ASPOSE_CLIENT_SECRET:
        return None
    return WordsApi(ASPOSE_CLIENT_ID, ASPOSE_CLIENT_SECRET)


def convert_doc_to_docx_aspose(file_bytes: bytes, filename: str) -> bytes:
    """Uses Aspose Words API to convert .doc -> .docx in memory."""
    api = get_words_api()
    if api is None:
        raise RuntimeError("Aspose credentials missing.")
    file_stream = io.BytesIO(file_bytes)
    file_stream.name = filename
    request = ConvertDocumentRequest(document=file_stream, format="docx")
    result = api.convert_document(request)
    return result


# ---------- Parsing helpers ----------
def extract_text_from_docx_bytes(docx_bytes: bytes):
    with io.BytesIO(docx_bytes) as memf:
        result = mammoth.extract_raw_text(memf)
        text = result.value
    return [line.strip() for line in text.split("\n") if line.strip()]


def extract_supplier(paragraphs):
    for p in paragraphs:
        m = re.search(r"Brand:\s*(.*)", p, flags=re.I)
        if m:
            brand = m.group(1).strip()
            code_match = re.search(r"Supplier Number:\s*(\S+)", " ".join(paragraphs))
            return f"{brand} ({code_match.group(1)})" if code_match else brand
    return "[Supplier]"


def extract_invoice_number(paragraphs):
    for p in paragraphs:
        m = re.search(r"Supplier Invoice Number:\s*(.*)", p, flags=re.I)
        if m:
            return m.group(1).strip()
    return "[Invoice Number]"


def extract_sn_number(paragraphs):
    for p in paragraphs:
        m = re.search(r"(?:Shipping Notice|Zalando Shipping Notice Number):\s*(\S+)", p, flags=re.I)
        if m:
            return m.group(1).strip()
    return "[SN Number]"


def extract_article_table(paragraphs):
    """Extract article details table for Article Not Ordered emails."""
    table_rows = []
    inside_table = False
    for p in paragraphs:
        if re.search(r"Config SKU", p, re.I):
            inside_table = True
            continue
        if inside_table:
            if not p.strip():
                break
            table_rows.append(p)
    return "\n".join(table_rows) if table_rows else "[Article information from ticket]"


def extract_price_table(paragraphs):
    """Extracts the table and formats it as Jira-style HTML."""
    table_lines = []
    header_found = False
    for p in paragraphs:
        if re.search(r"Config.?SKU", p, re.I):
            header_found = True
            header = [
                "Config-SKU", "Supp. Article #", "Supp. Color",
                "PO Price (after discount)", "Invoiced Price (after discount)",
                "% Deviation", "Total Position Difference", "Invoiced Qty"
            ]
            table_lines.append(header)
            continue
        if header_found:
            if not re.search(r"\d", p):
                break
            row = re.split(r"[\t|,; ]{2,}", p.strip())
            if len(row) >= 8:
                table_lines.append(row[:8])
            else:
                table_lines.append(row)

    if not table_lines or len(table_lines) <= 1:
        return "[Table information]"

    # Build HTML table
    html = '<table border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; text-align:center; width:100%;">'
    html += '<tr style="background-color:#f2f2f2; font-weight:bold;">'
    for head in table_lines[0]:
        html += f"<td>{head}</td>"
    html += "</tr>"
    for row in table_lines[1:]:
        html += "<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>"
    html += "</table>"
    return html


# ---------- Email Templates ----------
def generate_price_variance_email(supplier, invoice, table_html, name):
    return f"""Dear {supplier},<br><br>
I hope this email finds you well.<br>
Your invoice {invoice} is currently in clarification due to a price discrepancy.<br><br>
{table_html}<br><br>
Please get back to us within the next 3 working days, otherwise we will have to deduct the difference automatically via debit note.<br>
If you have further questions, please do not hesitate to reach out.<br><br>
Thank you and kind regards,<br>
{name}"""


def generate_article_not_ordered_email(supplier, sn_info, article_info, number_info, name):
    return f"""Dear {supplier},<br><br>
I hope this email finds you well.<br><br>
We have (an) item/s in quarantine storage as it looks like we have not ordered it and therefore we can not receive it/them. 
The articles have been delivered with {sn_info}.<br>
When items cannot be directly received due to specific issues they are sidelined and stored in our quarantine storage area 
(= a separate area in our warehouse). This additional clarification process is causing capacity losses and unforeseen costs.<br><br>
{article_info}<br><br>
We have received {number_info} units of this style.<br><br>
We have two options:<br><br>
1. Return (please provide the address, return label, and return authorization number).<br><br>
2. You can agree that we shall process the goods internally at Zalando's own discretion and thereby relinquish any and all rights 
that may have been reserved in the items (items will be processed further internally and then sold in bulk).<br><br>
Please confirm how you would like to proceed within the next 3 working days.<br>
Should we not hear from you until then, we will assume your tacit consent that we may proceed by processing the articles internally.<br><br>
If you have further questions, please do not hesitate to reach out.<br><br>
Thank you and kind regards,<br>
{name}"""


# ---------- MAIN ----------
if uploaded_file and user_name:
    try:
        with st.status("Processing file…", expanded=False) as status:
            status.update(label="Converting .doc → .docx with Aspose")
            file_bytes = uploaded_file.read()
            docx_bytes = convert_doc_to_docx_aspose(file_bytes, uploaded_file.name)

            status.update(label="Extracting text")
            paragraphs = extract_text_from_docx_bytes(docx_bytes)

        supplier = extract_supplier(paragraphs)
        invoice = extract_invoice_number(paragraphs)
        sn_info = extract_sn_number(paragraphs)

        if "price" in uploaded_file.name.lower():
            table_html = extract_price_table(paragraphs)
            email_body = generate_price_variance_email(supplier, invoice, table_html, user_name)
        else:
            article_info = extract_article_table(paragraphs)
            number_info = "[Number/size breakdown]"
            email_body = generate_article_not_ordered_email(supplier, sn_info, article_info, number_info, user_name)

        st.success("✅ File processed successfully!")
        st.markdown("**📧 Email Preview**")
        st.markdown(email_body, unsafe_allow_html=True)

        if st.button("📋 Copy Email to Clipboard"):
            st.code(email_body, language="html")
            st.info("You can now press Ctrl+C to copy the text manually.")

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")

