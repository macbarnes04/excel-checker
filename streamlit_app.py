import streamlit as st
from app import analyze_excel_folder  # your main logic function
import tempfile
import shutil
import os

st.title("Excel Submission Checker 🧮")

st.write("Upload a folder of `.xlsx` submissions to analyze:")

uploaded_files = st.file_uploader("Select .xlsx files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    temp_dir = tempfile.mkdtemp()
    for file in uploaded_files:
        file_path = os.path.join(temp_dir, file.name)
        with open(file_path, "wb") as f:
            f.write(file.getbuffer())

    st.success(f"Uploaded {len(uploaded_files)} files. Running analysis...")

    report = analyze_excel_folder(temp_dir)

    st.subheader("🧾 Report Summary")
    st.markdown(report)

    # optionally allow download
    st.download_button("Download Report", report, file_name="report.txt")

    shutil.rmtree(temp_dir)
