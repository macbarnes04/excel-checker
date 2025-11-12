import os
import tempfile
import shutil
import streamlit as st
from app import analyze_excel_folder, create_pdf_report  # make sure create_pdf_report is imported

st.title(" AI Detection for LBO Submissions 🔍📊")
st.write("Select all LBO Excel files from Canvas downloads")

uploaded_files = st.file_uploader(
    "Select all LBO Excel files from Canvas downloads",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:
    temp_dir = tempfile.mkdtemp()
    for file in uploaded_files:
        file_path = os.path.join(temp_dir, file.name)
        with open(file_path, "wb") as f:
            f.write(file.getbuffer())

    st.info("Running analysis... ⏳")

    report_text = analyze_excel_folder(temp_dir)

    st.success("Analysis complete! ✅")

    st.subheader("🧾 Report Summary")
    st.text(report_text)

    # Create PDF
    pdf_path = create_pdf_report(report_text, os.path.join(temp_dir, "report.pdf"))

    st.download_button(
        "Download PDF Report",
        data=open(pdf_path, "rb"),
        file_name="LBO_AI_Report.pdf",
        mime="application/pdf"
    )

    shutil.rmtree(temp_dir)
