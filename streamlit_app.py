import os
import io
import re
import json
import streamlit as st
from streamlit.components.v1 import html
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

def convert_doc_to_docx_aspose(file_bytes: bytes, filename: str) -> bytes:
    api = get_words_api()
    if api is None:
        raise RuntimeError("Aspose credentials missing.")
    file_stream = io.BytesIO(file_bytes)
    file_stream.name = filename
    req = ConvertDocumentRequest(document=file_stream, format="docx")
    return api.convert_document(req)

# ---------- Parsing helpers ----------
def extract_text_from_docx_bytes(docx_bytes: bytes):
    with io.BytesIO(docx_bytes) as memf:
        result = mammoth.extract_raw_text(memf)
    text = result.value
    return [line.strip() for line in text.split("\n") if line.strip()]

def detect_price_variance(paragraphs):
    blob = " ".join(p.lower() for p in paragraphs)
    return (
        ("po price" in blob and "invoiced price" in blob and "% deviation" in blob) or
        ("total position difference" in blob and "invoiced qty" in blob)
    )

def extract_supplier_combo(paragraphs):
    name = None
    code = None
    for i, p in enumerate(paragraphs):
        low = p.lower()
        if "supplier name" in low or re.search(r"^supplier:?$", low):
            name = paragraphs[i+1].strip() if i + 1 < len(paragraphs) else None
        elif "brand" in low and not name:
            name = paragraphs[i+1].strip() if i + 1 < len(paragraphs) else None
        m = re.search(r"supplier number[:\s]*([A-Z0-9]+)", p, re.I)
        if m:
            code = m.group(1)
    if name and code:
        return f"{name} ({code})"
    if name:
        return name
    if code:
        return f"({code})"
    return "[Supplier]"

def extract_invoice_number(paragraphs):
    for i, p in enumerate(paragraphs):
        if "supplier invoice number" in p.lower():
            return (paragraphs[i+1].strip() if i+1 < len(paragraphs) else p).strip()
        m = re.search(r"supplier invoice number[:\s]*(\S+)", p, re.I)
        if m:
            return m.group(1).strip()
    return "[Invoice Number]"

# ---- Price variance table to HTML (Jira style) ----
CONFIG_RE = re.compile(r"^[A-Z0-9]{3,}-[A-Z0-9]{2,}$")

def extract_price_table_html(paragraphs):
    # find header
    header_idx = None
    for i, p in enumerate(paragraphs):
        low = p.lower()
        if ("config" in low and "sku" in low) and ("po price" in " ".join(paragraphs[i:i+10]).lower()):
            header_idx = i
            break
    if header_idx is None:
        return "[Table information]"

    # canonical header
    headers = [
        "Config-SKU", "Supp. Article #", "Supp. Color",
        "PO Price (after discount)", "Invoiced Price (after discount)",
        "% Deviation", "Total Position Difference", "Invoiced Qty"
    ]

    rows = []
    i = header_idx + 1
    # walk until footer
    while i < len(paragraphs):
        line = paragraphs[i].strip()
        if "generated at" in line.lower():
            break
        if CONFIG_RE.match(line):  # start of a row
            cols = [line]
            # Jira/Mammoth typically flattens each cell onto the next lines; grab next 7 lines
            for k in range(7):
                j = i + 1 + k
                if j < len(paragraphs):
                    cols.append(paragraphs[j].strip())
                else:
                    cols.append("")
            # normalise column count
            cols = (cols + [""] * 8)[:8]
            rows.append(cols)
            i += 8
            continue
        i += 1

    if not rows:
        return "[Table information]"

    # Build HTML table
    html_parts = []
    html_parts.append('<table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse; text-align:center; width:100%;">')
    html_parts.append('<tr style="background-color:#f2f2f2; font-weight:bold;">' + "".join(f"<td>{h}</td>" for h in headers) + "</tr>")
    for r in rows:
        html_parts.append("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>")
    html_parts.append("</table>")
    return "".join(html_parts)

# ---------- Templates ----------
def tpl_price_variance(supplier, invoice, table_html, name):
    return f"""Dear {supplier},<br><br>
I hope this email finds you well.<br>
Your invoice {invoice} is currently in clarification due to a price discrepancy.<br><br>
{table_html}<br><br>
Please get back to us within the next 3 working days, otherwise we will have to deduct the difference automatically via debit note.<br>
If you have further questions, please do not hesitate to reach out.<br><br>
Thank you and kind regards,<br>
{name}"""

def tpl_article_not_ordered(supplier, sn_info, article_info, number_info, name):
    return f"""Dear {supplier},<br><br>
I hope this email finds you well.<br><br>
We have (an) item/s in quarantine storage as it looks like we have not ordered it and therefore we can not receive it/them. The articles have been delivered with {sn_info}.<br>
When items cannot be directly received due to specific issues they are sidelined and stored in our quarantine storage area (= a separate area in our warehouse). This additional clarification process is causing capacity losses and unforeseen costs.<br><br>
{article_info}<br><br>
We have received {number_info} units of this style.<br><br>
We have two options:<br><br>
1. Return (please provide the address, return label, and return authorization number).<br><br>
2. You can agree that we shall process the goods internally at Zalando's own discretion and thereby relinquish any and all rights that may have been reserved in the items (items will be processed further internally and then sold in bulk).<br><br>
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
            fbytes = uploaded_file.read()
            docx_bytes = convert_doc_to_docx_aspose(fbytes, uploaded_file.name)

            status.update(label="Extracting text")
            paragraphs = extract_text_from_docx_bytes(docx_bytes)

        # Optional debug
        # if st.checkbox("🔍 Show text"):
        #     for p in paragraphs: st.write(p)

        supplier = extract_supplier_combo(paragraphs)
        is_price = detect_price_variance(paragraphs)

        if is_price:
            invoice = extract_invoice_number(paragraphs)
            table_html = extract_price_table_html(paragraphs)
            email_html = tpl_price_variance(supplier, invoice, table_html, user_name)
        else:
            # basic fallbacks (we didn’t wire ANO parsing here again; use your previous ANO version if needed)
            sn_info = "[SN Number]"
            article_info = "[Article information from ticket]"
            number_info = "[Number/size breakdown]"
            email_html = tpl_article_not_ordered(supplier, sn_info, article_info, number_info, user_name)

        st.success("✅ File processed successfully!")
        st.markdown("**📧 Email Preview**")
        st.markdown(email_html, unsafe_allow_html=True)

        # Copy HTML to clipboard button
        email_js = json.dumps(email_html)
        html(f"""
        <div style='margin-top:10px;'>
          <button id="copyBtn" style="padding:0.6rem 1rem;border-radius:8px;border:1px solid #ddd;cursor:pointer;background:#f4f4f4;">
            📋 Copy Email (HTML) to Clipboard
          </button>
        </div>
        <script>
          const EMAIL_HTML = {email_js};
          document.getElementById('copyBtn').addEventListener('click', async () => {{
            try {{
              await navigator.clipboard.writeText(EMAIL_HTML);
              alert('Email HTML copied to clipboard!');
            }} catch (e) {{
              const ta = document.createElement('textarea');
              ta.value = EMAIL_HTML;
              document.body.appendChild(ta);
              ta.select();
              document.execCommand('copy');
              document.body.removeChild(ta);
              alert('Email HTML copied to clipboard!');
            }}
          }});
        </script>
        """, height=70)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
