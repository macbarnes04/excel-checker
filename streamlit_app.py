import streamlit as st
from app import analyze_excel_folder  # your main logic function
import tempfile
import shutil
import os

st.title(" AI Detection for LBO Submissions - PE Methods 🔍📊")

st.write("Upload all `.xlsx` submissions to analyze:")

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

    # Run analysis with spinner
    with st.spinner(f"🧠 Running analysis on {len(uploaded_files)} submissions..."):
        report = analyze_excel_folder(temp_dir)

    st.success("✅ Analysis complete!")

    st.subheader("🧾 Report Summary")
    st.markdown(report)

    st.download_button("Download Report", report, file_name="report.txt")

    shutil.rmtree(temp_dir)
