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
    """
    Use hierarchical clustering to group similar submissions.
    """
    clustering = AgglomerativeClustering(
        n_clusters=None,
        distance_threshold=1 - SIMILARITY_THRESHOLD,
        affinity='precomputed',
        linkage='complete'
    )
    dist_matrix = 1 - sim_matrix
    labels = clustering.fit_predict(dist_matrix)

    clusters = defaultdict(list)
    for file, label in zip(filenames, labels):
        clusters[label].append(file)
    return clusters


def main():
    print("🔍 Scanning Excel submissions...")

    # Step 1: Extract data from each Excel file
    records = []
    for filename in os.listdir(SUBMISSIONS_DIR):
        if filename.endswith(".xlsx"):
            path = os.path.join(SUBMISSIONS_DIR, filename)
            record = extract_excel_data(path)
            # Add a student name field for convenience
            record["student_name"] = record.get("lastModifiedBy") or record.get("creator") or "Unknown"
            records.append(record)

    df = pd.DataFrame(records)

    # Step 2: Compute similarities
    print("🧠 Computing similarity matrices...")
    text_sim = compute_text_similarity(df["text_content"].fillna(""))
    formula_sim = compute_formula_similarity(df["formula_content"].fillna(""))

    # Step 3: Detect duplicates
    text_dups = find_duplicates(text_sim, df["filename"], SIMILARITY_THRESHOLD)
    formula_dups = find_duplicates(formula_sim, df["filename"], SIMILARITY_THRESHOLD)

    # Step 4: Metadata anomalies
    metadata_flags = detect_metadata_anomalies(df)

    # Step 5: Clustering
    clusters = cluster_submissions(formula_sim, df["filename"])

    # Step 6: Build report with suspiciousness score
    print("📊 Generating report...")
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

    # Select final report columns
    report_cols = [
        "filename", "student_name", "creator", "lastModifiedBy", 
        "created", "modified", "suspicious_score"
    ]
    df[report_cols].to_csv(REPORT_FILE, index=False)

    # Step 7: Console summary
    print("\n✅ DONE. Results summary:")
    print(f"  - Total submissions: {len(df)}")
    print(f"  - Possible duplicate text pairs: {len(text_dups)}")
    print(f"  - Possible duplicate formula pairs: {len(formula_dups)}")
    print(f"  - Metadata anomalies: {len(metadata_flags)}")
    print(f"  - Report saved to {REPORT_FILE}")

    # Group and show duplicates with student names
    def display_dup_list(dups, label):
        if dups:
            print(f"\n⚠️  {label} similarities ({len(dups)} pairs):")
            for f1, f2, sim in dups:
                s1 = df.loc[df["filename"] == f1, "student_name"].values[0]
                s2 = df.loc[df["filename"] == f2, "student_name"].values[0]
                print(f"   {s1} ↔ {s2} ({sim*100:.1f}% similar)")
        else:
            print(f"\n✅ No significant {label} similarities found.")

    display_dup_list(text_dups, "text")
    display_dup_list(formula_dups, "formula")

    # Clusters summary
    print("\n📁 Clusters (similar groups):")
    for cluster_id, files in clusters.items():
        if len(files) > 1:
            names = [df.loc[df["filename"] == f, "student_name"].values[0] for f in files]
            print(f"  Cluster {cluster_id}: {names}")

    if metadata_flags:
        print("\n⚠️ Metadata Flags:")
        for f in metadata_flags:
            print("  -", f)
