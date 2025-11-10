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
        st.info("Converting .doc to .docx using CloudConvert...")
        api_key = os.getenv("CLOUDCONVERT_API_KEY")
        if not api_key:
            st.error("CloudConvert API key not set.")
        else:
            try:
                # STEP 1: Create the import task
                import_response = requests.post(
                    "https://api.cloudconvert.com/v2/import/upload",
                    headers={"Authorization": f"Bearer {api_key}"},
                )
                import_task = import_response.json()["data"]
                upload_url = import_task["result"]["form"]["url"]
                upload_params = import_task["result"]["form"]["parameters"]

                # STEP 2: Upload the .doc file
                with open(file_name, "rb") as f:
                    files = {'file': (file_name, f)}
                    upload_params["file"] = files["file"]
                    upload_response = requests.post(upload_url, data=upload_params, files=files)

                # STEP 3: Create the convert task
                convert_response = requests.post(
                    "https://api.cloudconvert.com/v2/convert",
                    headers={"Authorization": f"Bearer {api_key}"},
                    json={
                        "input": "import-upload",
                        "file": file_name,
                        "input_format": "doc",
                        "output_format": "docx"
                    }
                )
                convert_task = convert_response.json()["data"]
                task_id = convert_task["id"]

                # STEP 4: Wait for conversion to finish
                status = ""
                while status != "finished":
                    task_response = requests.get(
                        f"https://api.cloudconvert.com/v2/tasks/{task_id}",
                        headers={"Authorization": f"Bearer {api_key}"}
                    )
                    status = task_response.json()["data"]["status"]

                # STEP 5: Download the .docx result
                export_url = task_response.json()["data"]["result"]["files"][0]["url"]
                new_file_name = file_name.replace(".doc", ".docx")
                r = requests.get(export_url)
                with open(new_file_name, "wb") as f:
                    f.write(r.content)

                # STEP 6: Parse and preview
                doc = Document(new_file_name)
                text = "\n".join([para.text for para in doc.paragraphs])
                st.text_area("📨 Email Preview", text, height=300)

            except Exception as e:
                st.error(f"Conversion failed: {e}")

    else:
        try:
            doc = Document(file_name)
            text = "\n".join([para.text for para in doc.paragraphs])
            st.text_area("📨 Email Preview", text, height=300)
        except Exception as e:
            st.error(f"Failed to read .docx: {e}")
