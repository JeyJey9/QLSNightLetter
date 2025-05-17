import os
import pandas as pd
from parsing import parse_pdf_values, extract_broken_pdf
import pdfplumber

def process_pdf(pdf_path, root_folder, output_root, log_func, update_prog):
    """
    Parses a PDF and writes its extracted values to XLSX.
    Uses broken-parser for PDFs >5 pages.
    """
    rel = os.path.relpath(pdf_path, root_folder)
    out_xlsx = os.path.join(output_root, os.path.splitext(rel)[0] + ".xlsx")
    os.makedirs(os.path.dirname(out_xlsx), exist_ok=True)

    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)

    if page_count > 5:
        log_func(f"[INFO] Broken PDF detected (>5 pages), using fallback parser: {os.path.basename(pdf_path)}\n")
        vals = extract_broken_pdf(pdf_path)
    else:
        vals = parse_pdf_values(pdf_path)

    if not vals:
        log_func(f"[WARN] Skipped {os.path.basename(pdf_path)} (insufficient data)\n")
        update_prog()
        return False

    pd.DataFrame([vals]).to_excel(out_xlsx, sheet_name="Data", index=False)
    log_func(f"[OK] Extracted Data for {os.path.basename(pdf_path)}\n")
    update_prog()
    return True
