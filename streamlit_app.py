import streamlit as st
import requests
import os
from docx import Document
import time

st.set_page_config(page_title="Gemini Email Generator", layout="centered")
st.markdown("## 📩 Gemini Email Generator")
st.markdown("**Upload .doc or .docx Jira Ticket**")

uploaded_file = st.file_uploader("Drag and drop file here", type=["doc", "docx"])

if uploaded_file:
    file_name = uploaded_file.name
    st.success(f"✅ File uploaded successfully: {file_name}")

    with open(file_name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # DOCX handling directly
    if file_name.endswith(".docx"):
        try:
            doc = Document(file_name)
            text = "\n".join([para.text for para in doc.paragraphs])
            st.text_area("📨 Email Preview", text, height=300)
        except Exception as e:
            st.error(f"❌ Failed to read .docx: {e}")

    # DOC needs conversion
    elif file_name.endswith(".doc"):
        st.info("🔄 Converting .doc to .docx using CloudConvert...")

        api_key = os.getenv("CLOUDCONVERT_API_KEY")
        if not api_key:
            st.error("❌ CloudConvert API key not set. Please check your environment variables.")
        else:
            headers = {"Authorization": f"Bearer {api_key}"}

            # STEP 1: Create job
            try:
                job_resp = requests.post(
                    "https://api.cloudconvert.com/v2/jobs",
                    headers=headers,
                    json={
                        "tasks": {
                            "import-upload": {
                                "operation": "import/upload"
                            },
                            "convert-doc": {
                                "operation": "convert",
                                "input": "import-upload",
                                "input_format": "doc",
                                "output_format": "docx"
                            },
                            "export-url": {
                                "operation": "export/url",
                                "input": "convert-doc"
                            }
                        }
                    },
                )
                job_data = job_resp.json()
                job_id = job_data.get("data", {}).get("id")

                if not job_id:
                    st.error("❌ Could not get job ID. Check API key or usage limits.")
                    st.json(job_data)
                    st.stop()

                upload_task = [
                    t for t in job_data["data"]["tasks"]
                    if t["name"] == "import-upload"
                ][0]
                upload_url = upload_task["result"]["form"]["url"]
                upload_params = upload_task["result"]["form"]["parameters"]

                # STEP 2: Upload file to CloudConvert
                with open(file_name, "rb") as file_stream:
                    upload_payload = {**upload_params, "file": file_stream}
                    upload_resp = requests.post(upload_url, files=upload_payload)

                if upload_resp.status_code != 201:
                    st.error(f"❌ Upload failed with status {upload_resp.status_code}")
                    st.text(upload_resp.text)
                    st.stop()

                # STEP 3: Wait for conversion to complete
                time.sleep(6)
                job_status_resp = requests.get(f"https://api.cloudconvert.com/v2/jobs/{job_id}", headers=headers)
                job_status_data = job_status_resp.json()

                export_task = [
                    t for t in job_status_data["data"]["tasks"]
                    if t["name"] == "export-url"
                ][0]

                file_url = export_task["result"]["files"][0]["url"]
                output_file = file_name.replace(".doc", ".docx")

                # STEP 4: Download converted file
                converted_file = requests.get(file_url)
                with open(output_file, "wb") as f_out:
                    f_out.write(converted_file.content)

                # STEP 5: Show result
                doc = Document(output_file)
                text = "\n".join([para.text for para in doc.paragraphs])
                st.success("✅ Conversion successful!")
                st.text_area("📨 Email Preview", text, height=300)

            except Exception as e:
                st.error(f"❌ Conversion failed: {e}")
