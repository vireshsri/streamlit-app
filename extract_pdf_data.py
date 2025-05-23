import zipfile
import pdfplumber
import pandas as pd
import re
from pathlib import Path

# ------------------------- Configuration -------------------------
ZIP_FILE = "FCL (2).zip"  # <--- Change this to your actual ZIP file path
EXTRACT_DIR = "extracted_pdfs"
OUTPUT_FILE = "Extracted_Loan_Details.xlsx"
# -----------------------------------------------------------------

# Define keywords to extract
KEYWORDS = [
    "Loan Account No",
    "Balance Amount",
    "Principal Outstanding",
    "Interest accrued since last due date",
    "Overdue Installments",
    "Overdue Principal",
    "Overdue Interest",
    "Interest Till Locking Period",
    "Unbilled Installment",
    "Other Dues",
    "Pre Payment Interest",
    "LPI Amount till date",
    "LPC Amount till date",
    "Amount Receivable (Rs.)",
    "Unrealized Amount (Rs.)",
    "Advance EMI Payable (Rs.)",
    "Total Amount Receivable (Rs.)",
    "Excess Amount"
]

def extract_zip(zip_path, extract_to):
    """Extract the ZIP file to a directory."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"✅ Extracted ZIP to: {extract_to}")

def extract_text_lines(pdf_path):
    """Extract all text lines from a PDF."""
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))
    return lines

def extract_loan_data(lines):
    """Extract loan-related fields from lines."""
    data = {k: "" for k in KEYWORDS}
    for i, line in enumerate(lines):
        for keyword in KEYWORDS:
            if keyword in line:
                if keyword == "Loan Account No":
                    # Extract full alphanumeric loan number
                    match = re.search(r"Loan Account No[:\s]+([A-Z0-9\-]+)", line)
                    if match:
                        data[keyword] = match.group(1).strip()
                    else:
                        # Try next line
                        if i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            match_next = re.search(r"([A-Z0-9\-]+)", next_line)
                            if match_next:
                                data[keyword] = match_next.group(1).strip()
                else:
                    # Generic numeric match pattern
                    match = re.search(rf"{re.escape(keyword)}\s*[:\-]?\s*([\d,.\w]+)", line)
                    if match:
                        data[keyword] = match.group(1).strip()
                    else:
                        # Try next line if value not found
                        if i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            if re.match(r"^[\d,.\w]+$", next_line):
                                data[keyword] = next_line
                break
    return data

def process_pdf_folder(folder_path, output_path):
    """Main function to process PDFs and export results."""
    pdf_files = list(Path(folder_path).rglob("*.pdf"))
    all_data = []

    for pdf_file in pdf_files:
        lines = extract_text_lines(pdf_file)
        record = extract_loan_data(lines)
        record["File Name"] = pdf_file.name
        all_data.append(record)

    df = pd.DataFrame(all_data)
    df.to_excel(output_path, index=False)
    print(f"📄 Data extracted and saved to: {output_path}")

# ------------------------ Run Everything ------------------------

def main():
    extract_zip(ZIP_FILE, EXTRACT_DIR)
    process_pdf_folder(EXTRACT_DIR, OUTPUT_FILE)

if __name__ == "__main__":
    main()
