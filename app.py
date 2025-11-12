import os
import re
import json
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.cluster import AgglomerativeClustering
from datetime import datetime
from collections import defaultdict
from fpdf import FPDF

# ========== CONFIG ==========
SUBMISSIONS_DIR = "submissions"   # folder with all Excel files
REPORT_FILE = "ai_detection_report.csv"
SIMILARITY_THRESHOLD = 0.9        # flag if two files are >90% similar
# ============================


def extract_excel_data(filepath):
    """
    Extract metadata, text, and formulas from an Excel file.
    Returns a dict of structured data.
    """
    data = {
        "filename": os.path.basename(filepath),
        "creator": None,
        "lastModifiedBy": None,
        "created": None,
        "modified": None,
        "text_content": "",
        "formula_content": "",
        "num_sheets": 0
    }

    try:
        wb = load_workbook(filepath, data_only=False)
        props = wb.properties

        data["creator"] = props.creator
        data["lastModifiedBy"] = props.lastModifiedBy
        data["created"] = props.created
        data["modified"] = props.modified
        data["num_sheets"] = len(wb.sheetnames)

        text_cells = []
        formula_cells = []

        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    if cell.data_type == "s" and cell.value:
                        text_cells.append(str(cell.value))
                    elif cell.data_type == "f" and cell.value:
                        formula_cells.append(str(cell.value))

        data["text_content"] = " ".join(text_cells)
        data["formula_content"] = " ".join(formula_cells)

    except Exception as e:
        data["error"] = str(e)

    return data


def compute_text_similarity(texts):
    """
    Compute cosine similarity between text contents of all files.
    """
    vectorizer = TfidfVectorizer(stop_words="english")
    tfidf = vectorizer.fit_transform(texts)
    sim_matrix = cosine_similarity(tfidf)
    return sim_matrix


def compute_formula_similarity(formulas):
    """
    Compute cosine similarity between formula contents of all files.
    """
    vectorizer = TfidfVectorizer(token_pattern=r"[\w\+\-\*/\(\)]+")
    tfidf = vectorizer.fit_transform(formulas)
    sim_matrix = cosine_similarity(tfidf)
    return sim_matrix


def find_duplicates(sim_matrix, filenames, threshold=0.9):
    """
    Return pairs of files with similarity above threshold.
    """
    duplicates = []
    n = len(filenames)
    for i in range(n):
        for j in range(i + 1, n):
            if sim_matrix[i, j] >= threshold:
                duplicates.append((filenames[i], filenames[j], round(sim_matrix[i, j], 3)))
    return duplicates


def detect_metadata_anomalies(df):
    """
    Detect metadata anomalies like identical authors, timestamps near submission, etc.
    """
    anomalies = []

    # Identical author or lastModifiedBy
    grouped = df.groupby(["creator", "lastModifiedBy"]).size()
    for idx, count in grouped.items():
        if count > 1:
            anomalies.append(f"{count} files share identical metadata: {idx}")

    # Keyword-based anomalies
    for _, row in df.iterrows():
        if row["creator"]:
            if any(k in str(row["creator"]).lower() for k in ["copilot", "chatgpt", "claude", "openai"]):
                anomalies.append(f"{row['filename']} references AI in metadata ({row['creator']})")

    return anomalies

def cluster_submissions(sim_matrix, filenames):
    from collections import defaultdict
    import numpy as np
    
    clusters = defaultdict(list)
    
    if len(filenames) < 2:
        # Only one file → just return it as a single cluster
        clusters[0] = filenames
        return clusters
    
    # safe to compute distance matrix and cluster
    dist_matrix = 1 - sim_matrix
    dist_matrix = np.nan_to_num(dist_matrix)
    
    clustering = AgglomerativeClustering(
        n_clusters=None,
        distance_threshold=1 - 0.9,  # SIMILARITY_THRESHOLD
        metric='precomputed',
        linkage='average'
    )
    labels = clustering.fit_predict(dist_matrix)
    
    for file, label in zip(filenames, labels):
        clusters[label].append(file)
    
    return clusters



def analyze_excel_folder(submissions_dir):
    """
    Run full Excel submission analysis for the given directory.
    Returns a formatted report string and saves a CSV report.
    """
    records = []
    for filename in os.listdir(submissions_dir):
        if filename.endswith(".xlsx"):
            path = os.path.join(submissions_dir, filename)
            record = extract_excel_data(path)
            record["student_name"] = record.get("lastModifiedBy") or record.get("creator") or "Unknown"
            records.append(record)

    df = pd.DataFrame(records)

    if len(df) == 0:
        return "No valid Excel files found."

    text_sim = compute_text_similarity(df["text_content"].fillna(""))
    formula_sim = compute_formula_similarity(df["formula_content"].fillna(""))

    text_dups = find_duplicates(text_sim, df["filename"], SIMILARITY_THRESHOLD)
    formula_dups = find_duplicates(formula_sim, df["filename"], SIMILARITY_THRESHOLD)

    metadata_flags = detect_metadata_anomalies(df)
    clusters = cluster_submissions(formula_sim, df["filename"])

    suspicious_scores = []
    for i, row in df.iterrows():
        score = 0
        if any(row["filename"] in dup for dup in formula_dups):
            score += 3
        if any(row["filename"] in dup for dup in text_dups):
            score += 2
        if row["creator"] and any(k in str(row["creator"]).lower() for k in ["copilot", "chatgpt", "claude", "openai"]):
            score += 5
        suspicious_scores.append(score)

    df["suspicious_score"] = suspicious_scores
    df = df.sort_values("suspicious_score", ascending=False)

    # Save report CSV
    report_path = os.path.join(submissions_dir, REPORT_FILE)
    df.to_csv(report_path, index=False)

    # Build readable text summary
    summary = []
    summary.append(f"✅ Total submissions: {len(df)}")
    summary.append(f"⚠️ Text duplicates: {len(text_dups)} | Formula duplicates: {len(formula_dups)}")
    summary.append(f"📁 Clusters: {sum(len(v) > 1 for v in clusters.values())}")
    summary.append(f"🔎 Metadata anomalies: {len(metadata_flags)}")

    summary.append("\nTop suspicious submissions:")
    for _, r in df.head(5).iterrows():
        summary.append(f" - {r['student_name']} ({r['filename']}): score {r['suspicious_score']}")

    if metadata_flags:
        summary.append("\nMetadata flags:")
        for f in metadata_flags:
            summary.append(f" - {f}")

    return "\n".join(summary)


def create_pdf_report(report_text, output_path="report.pdf"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)

    # Add report lines
    for line in report_text.split("\n"):
        pdf.multi_cell(0, 6, line)
    
    pdf.output(output_path)
    return output_path


# Optional: allow running directly from terminal
if __name__ == "__main__":
    print(analyze_excel_folder(SUBMISSIONS_DIR))

