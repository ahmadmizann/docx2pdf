import streamlit as st
import tempfile
import os
import subprocess
from pathlib import Path

# ---------------------------------------------------------
# Streamlit App Title
# ---------------------------------------------------------
st.title("📄 DOCX → PDF Converter")

# ---------------------------------------------------------
# File uploader: allow user to upload multiple .docx files
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload DOCX files",           # Instruction shown to user
    accept_multiple_files=True,    # Allow multiple files at once
    type="docx"                    # Restrict uploads to only DOCX
)

# ---------------------------------------------------------
# If the user has uploaded files, process them
# ---------------------------------------------------------
if uploaded_files:
    for uploaded_file in uploaded_files:
        # -------------------------------------------------
        # Create a temporary folder to store the uploaded
        # file and the converted PDF. This folder is
        # automatically deleted when finished.
        # -------------------------------------------------
        with tempfile.TemporaryDirectory() as tmpdir:

            # Path for the uploaded DOCX
            docx_path = os.path.join(tmpdir, uploaded_file.name)

            # Path for the converted PDF (same name, .pdf extension)
            pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

            # -------------------------------------------------
            # Save the uploaded DOCX file into our temp folder
            # -------------------------------------------------
            with open(docx_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # -------------------------------------------------
            # Use LibreOffice (must be installed on the system)
            # to convert the DOCX file into a PDF
            # -------------------------------------------------
            try:
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf",
                    "--outdir", tmpdir, docx_path
                ], check=True)
            except Exception as e:
                st.error(f"⚠️ Conversion failed for {uploaded_file.name}: {e}")
                continue  # Skip this file and move to the next one

            # -------------------------------------------------
            # If conversion succeeded, show success message
            # -------------------------------------------------
            st.success(f"✅ Converted {uploaded_file.name} to PDF")

            # -------------------------------------------------
            # Provide a download button for the converted PDF
            # -------------------------------------------------
            with open(pdf_path, "rb") as pdf_file:
                st.download_button(
                    label=f"📥 Download {Path(pdf_path).name}",
                    data=pdf_file,
                    file_name=Path(pdf_path).name,
                    mime="application/pdf",
                )

            # -------------------------------------------------
            # Show a PDF preview directly inside the app
            # Using Streamlit's official st.pdf() API
            # -------------------------------------------------
            st.pdf(pdf_path, height="stretch")

# ---------------------------------------------------------
# If no file uploaded, display instructions
# ---------------------------------------------------------
else:
    st.write("📂 Please upload one or more DOCX files to begin")
