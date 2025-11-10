import streamlit as st
import requests
import os
from docx import Document

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
        st.info("🔄 Converting .doc to .docx using CloudConvert...")

        api_key = os.getenv("CLOUDCONVERT_API_KEY")
        if not api_key:
            st.error("❌ CloudConvert API key not set. Please check your environment variables.")
        else:
            # STEP 1: Get upload URL
            import_url = "https://api.cloudconvert.com/v2/import/upload"
            headers = {"Authorization": f"Bearer {api_key}"}
            import_resp = requests.post(import_url, headers=headers).json()

            try:
                upload_url = import_resp["data"]["url"]
                upload_file = {'file': open(file_name, 'rb')}
                upload_result = requests.post(upload_url, files=upload_file)

                if upload_result.status_code != 200:
                    raise Exception("Upload failed.")

                import_id = import_resp["data"]["id"]

                # STEP 2: Create conversion job
                job_url = "https://api.cloudconvert.com/v2/jobs"
                job_data = {
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

                job_resp = requests.post(job_url, headers=headers, json=job_data).json()
                job_id = job_resp["data"]["id"]

                # Wait for job to finish
                import time
                time.sleep(5)
                job_status_url = f"https://api.cloudconvert.com/v2/jobs/{job_id}"
                job_result = requests.get(job_status_url, headers=headers).json()

                export_url = job_result["data"]["tasks"][-1]["result"]["files"][0]["url"]
                output_name = file_name.replace(".doc", ".docx")

                # Download the converted file
                download_resp = requests.get(export_url)
                with open(output_name, "wb") as out_file:
                    out_file.write(download_resp.content)

                st.success("✅ Conversion successful. Processing the .docx file now...")

                # Extract text from converted docx
                doc = Document(output_name)
                text = "\n".join([para.text for para in doc.paragraphs])
                st.text_area("📨 Email Preview", text, height=300)

            except Exception as e:
                st.error(f"❌ Conversion failed: {e}")

    else:
        # Directly handle .docx files
        try:
            doc = Document(file_name)
            text = "\n".join([para.text for para in doc.paragraphs])
            st.text_area("📨 Email Preview", text, height=300)
        except Exception as e:
            st.error(f"❌ Failed to read .docx file: {e}")
