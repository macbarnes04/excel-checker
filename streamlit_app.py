import os
import tempfile
import shutil
import streamlit as st
from app import analyze_excel_folder, create_pdf_report  # make sure create_pdf_report is imported

st.title("AI Detection - LBO Submissions 🔍")

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

    # --- Run analysis and unpack results ---
    results = analyze_excel_folder(temp_dir)

    report_text = results["report_text"]
    submissions_df = results["df"]
    text_duplicates = results["text_dups"]
    formula_duplicates = results["formula_dups"]
    metadata_anomalies = results["metadata_flags"]

    # ✅ New: unpack relative similarity + clusters
    formula_dups_relative = results.get("formula_dups_relative", [])
    clusters = results.get("clusters", [])

    st.success("Analysis complete! ✅")

    st.subheader("🧾 Report Summary")
    st.text(report_text)

   # --- Create PDF ---
    output_path = os.path.join(temp_dir, "LBO_AI_Report.pdf")

    pdf_path = create_pdf_report(
        df=submissions_df,
        text_dups=text_duplicates,
        formula_dups=formula_duplicates,
        formula_dups_relative=formula_dups_relative,
        clusters=clusters,
        metadata_flags=metadata_anomalies,
        output_path=output_path,
    )

    # --- Download button ---
    st.download_button(
        "Download PDF Report",
        data=open(pdf_path, "rb"),
        file_name="LBO_AI_Report.pdf",
        mime="application/pdf"
    )

    # Clean up temporary directory
    shutil.rmtree(temp_dir)
