import streamlit as st
import zipfile
import pdfplumber
import pandas as pd
import re
import os
from pathlib import Path
from tempfile import TemporaryDirectory

# Keywords to extract
KEYWORDS = [
    "Loan Account No", "Balance Amount", "Principal Outstanding",
    "Interest accrued since last due date", "Overdue Installments",
    "Overdue Principal", "Overdue Interest", "Interest Till Locking Period",
    "Unbilled Installment", "Other Dues", "Pre Payment Interest",
    "LPI Amount till date", "LPC Amount till date", "Amount Receivable (Rs.)",
    "Unrealized Amount (Rs.)", "Advance EMI Payable (Rs.)",
    "Total Amount Receivable (Rs.)", "Excess Amount"
]

def extract_text_lines(pdf_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))
    return lines

def extract_loan_data(lines):
    data = {k: "" for k in KEYWORDS}
    for i, line in enumerate(lines):
        for keyword in KEYWORDS:
            if keyword in line:
                if keyword == "Loan Account No":
                    match = re.search(r"Loan Account No[:\s]+([A-Z0-9\-]+)", line)
                    if match:
                        data[keyword] = match.group(1).strip()
                else:
                    match = re.search(rf"{re.escape(keyword)}\s*[:\-]?\s*([\d,.\w]+)", line)
                    if match:
                        data[keyword] = match.group(1).strip()
        if keyword == "Loan Account No" and not data[keyword]:
            next_line = lines[i + 1] if i + 1 < len(lines) else ""
            match = re.search(r"([A-Z0-9\-]+)", next_line)
            if match:
                data[keyword] = match.group(1).strip()
    return data

def process_pdfs(pdf_dir):
    pdf_files = list(Path(pdf_dir).rglob("*.pdf"))
    all_data = []
    for pdf_file in pdf_files:
        lines = extract_text_lines(pdf_file)
        record = extract_loan_data(lines)
        record["File Name"] = pdf_file.name
        all_data.append(record)
    return pd.DataFrame(all_data)

# ---------- Streamlit App ----------
st.title("📄 Loan PDF Extractor Tool")
uploaded_zip = st.file_uploader("Upload ZIP of PDF Loan Statements", type="zip")

if uploaded_zip:
    with TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, "input.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.getbuffer())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        st.success("✅ ZIP extracted. Processing PDFs...")

        df = process_pdfs(temp_dir)

        st.success("✅ Extraction complete!")
        st.dataframe(df)

        # Download link
        output_path = os.path.join(temp_dir, "extracted_loan_data.xlsx")
        df.to_excel(output_path, index=False)
        with open(output_path, "rb") as f:
            st.download_button("📥 Download Excel", f, file_name="loan_data.xlsx")