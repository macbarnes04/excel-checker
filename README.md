# LBO AI Detection 🧮🤖

**Automated Excel Submission Analyzer for LBO Modeling Exercises**

https://excel-checker-pe-methods-fall-2025.streamlit.app/ 

This Streamlit app allows instructors to quickly check Excel submissions for potential AI-generated content, formula duplication, and metadata anomalies.

---

## ✅ Features / What it checks

1. **Formula similarity**  
   - Detects submissions with highly similar formulas (>90% similarity by default).  
   - Flags potential copying between students.  

2. **Text similarity**  
   - Analyzes textual content in Excel cells (notes, comments, input values).  
   - Flags submissions with unusually high similarity.

3. **Metadata analysis**  
   - Checks Excel document metadata (`creator`, `lastModifiedBy`, timestamps).  
   - Flags submissions referencing AI tools (e.g., Copilot, ChatGPT, Claude, OpenAI).  
   - Flags identical metadata across multiple files (may indicate reuse or copying).

4. **Suspiciousness scoring**  
   - Each submission receives a score based on formula/text duplication and metadata flags.  
   - Higher scores indicate a higher likelihood of potential AI assistance or copying.

5. **Clustering**  
   - Groups submissions with similar formulas into clusters for easy inspection.  

6. **Report generation**  
   - Produces a CSV report with file names, student identifiers (from metadata), suspiciousness scores, and flags.  
   - Can be viewed in the browser and downloaded.

---

## ⚠️ Limitations / What it does **not** check

- Does **not** detect manually rewritten AI-generated content if formulas and text are changed.  
- Does **not** detect plagiarism outside the Excel file (e.g., from PDFs, Google Docs, or external websites).  
- Metadata analysis is limited to what Excel exposes; users may manually edit or remove metadata.  
- The suspiciousness score is a heuristic — flagged files **require human review**.  

---

## 📂 How to use

1. Download all student `.xlsx` submissions from Canvas.  
2. Open the Streamlit app.  
3. Upload all Excel files simultaneously.  
4. Wait for the analysis to finish (spinner + success message).  
5. Review the report in the app or download as CSV.  

---

## ⚡ Notes

- Works best with **Excel files downloaded directly from Canvas**.  
- Designed for **LBO modeling exercises** but can work with other Excel-based assignments.  
- Minimal setup: Python 3.11 recommended, install requirements via `pip install -r requirements.txt`.

