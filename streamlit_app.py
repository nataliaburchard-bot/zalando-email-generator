
import streamlit as st
import requests
from docx import Document
import os

st.set_page_config(page_title="Gemini Email Generator", layout="centered")

st.markdown("## 📩 Gemini Email Generator")
st.markdown("**Upload .doc or .docx Jira Ticket**")

uploaded_file = st.file_uploader("Drag and drop file here", type=["doc", "docx"])

if uploaded_file:
    file_name = uploaded_file.name
    st.success(f"✅ File uploaded successfully: {file_name}")

    with open(file_name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    if file_name.endswith(".doc"):
        st.info("Converting .doc to .docx...")
        api_key = os.getenv("CLOUDCONVERT_API_KEY")
        if not api_key:
            st.error("CloudConvert API key not set.")
        else:
            convert_url = "https://api.cloudconvert.com/v2/import/upload"
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            # Full CloudConvert logic goes here...
            st.warning("Further processing logic goes here...")
    else:
        # Directly handle .docx
        try:
            doc = Document(file_name)
            text = "
".join([para.text for para in doc.paragraphs])
            st.text_area("📨 Email Preview", text, height=300)
        except Exception as e:
            st.error(f"Failed to read .docx: {e}")
