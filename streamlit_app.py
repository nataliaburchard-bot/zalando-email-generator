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
    cid, sec = None, None
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


def convert_doc_to_docx_aspose(file_bytes: bytes, uploaded_file_name: str) -> bytes:
    """Uses Aspose Words API to convert .doc → .docx in memory."""
    api = get_words_api()
    if api is None:
        raise RuntimeError("Aspose credentials missing.")
    file_stream = io.BytesIO(file_bytes)
    file_stream.name = uploaded_file_name
    request = ConvertDocumentRequest(document=file_stream, format="docx")
    return api.convert_document(request)


# ---------- Parsing helpers ----------
def extract_text_from_docx_bytes(docx_bytes: bytes):
    with io.BytesIO(docx_bytes) as memf:
        result = mammoth.extract_raw_text(memf)
        text = result.value
    return [line.strip() for line in text.split("\n") if line.strip()]


def extract_supplier(paragraphs):
    brand, supplier = None, None
    for i, p in enumerate(paragraphs):
        low = p.lower()
        if "supplier number" in low:
            supplier = paragraphs[i + 1].strip() if i + 1 < len(paragraphs) else p
        elif "brand" in low:
            brand = paragraphs[i + 1].strip() if i + 1 < len(paragraphs) else p
    if brand and supplier:
        return f"{brand} ({supplier})"
    return brand or supplier or "[Supplier]"


def extract_sn_info(paragraphs):
    for i, p in enumerate(paragraphs):
        if "shipping notice" in p.lower():
            return paragraphs[i + 1].strip() if i + 1 < len(paragraphs) else p
    return "[SN Info]"


def extract_article_info(paragraphs):
    for i, p in enumerate(paragraphs):
        if "example sku" in p.lower():
            sku = paragraphs[i + 1].strip() if i + 1 < len(paragraphs) else p
            return f"Example SKU {sku}"
    return "[Article Info]"


def extract_number_info(paragraphs):
    """
    Matches Jira’s pattern:
    EAN (13 digits)
    next line -> size
    next line -> quantity
    next line -> SKU
    """
    lines = []
    for i, p in enumerate(paragraphs):
        if re.fullmatch(r"\d{13}", p):  # found an EAN
            ean = p
            size = paragraphs[i + 1].strip() if i + 1 < len(paragraphs) else ""
            qty = paragraphs[i + 2].strip() if i + 2 < len(paragraphs) else ""
            sku = paragraphs[i + 3].strip() if i + 3 < len(paragraphs) else ""
            lines.append(f"{ean}  {size}  {qty}  {sku}")
    return "\n".join(lines) if lines else "[Number/size breakdown]"


# ---------- Email templates ----------
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
We have (an) item/s in quarantine storage as it looks like we have not ordered it and therefore we cannot receive it/them. The articles have been delivered with {sn_info}.

When items cannot be directly received due to specific issues they are sidelined and stored in our quarantine storage area (= a separate area in our warehouse). This additional clarification process is causing capacity losses and unforeseen costs.

{article_info}

We have received:
{number_info}

We have two options:

1. Return (please provide the address, return label, and return authorization number).
2. You can agree that we shall process the goods internally at Zalando's own discretion and thereby relinquish any and all rights that may have been reserved in the items (items will be processed further internally and then sold in bulk).

Please confirm how you would like to proceed within the next 3 working days.
Should we not hear from you until then, we will assume your tacit consent that we may proceed by processing the articles internally.

If you have further questions, please do not hesitate to reach out.

Thank you and kind regards,
{name}"""


# ---------- MAIN ----------
if uploaded_file and user_name:
    try:
        if not (ASPOSE_CLIENT_ID and ASPOSE_CLIENT_SECRET):
            st.error("Missing Aspose credentials.")
            st.stop()

        with st.status("Processing file…", expanded=False) as status:
            status.update(label="Converting .doc → .docx with Aspose")
            file_bytes = uploaded_file.read()
            docx_bytes = convert_doc_to_docx_aspose(file_bytes, uploaded_file.name)

            status.update(label="Extracting text")
            paragraphs = extract_text_from_docx_bytes(docx_bytes)

        # Optional debug viewer
        if st.checkbox("🔍 Show extracted text for debugging"):
            for p in paragraphs:
                st.write(p)

        supplier = extract_supplier(paragraphs)
        sn_info = extract_sn_info(paragraphs)
        article_info = extract_article_info(paragraphs)
        number_info = extract_number_info(paragraphs)

        if "price" in uploaded_file.name.lower():
            invoice = "[Invoice Number]"
            table = "[Table information]"
            email_body = generate_price_variance_email(supplier, invoice, table, user_name)
        else:
            email_body = generate_article_not_ordered_email(
                supplier, sn_info, article_info, number_info, user_name
            )

        st.success("✅ File processed successfully!")
        st.markdown("**📧 Email Preview**")
        st.text_area("", email_body, height=500)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
