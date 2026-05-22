"""
Auth Radar - Extraction Router

Decides HOW to extract a file based on its type and content:
  - CSV / XLSX  → structured parser (no OCR needed)
  - Text-based PDF → direct text extraction via pdfplumber/PyPDF2
  - Scanned PDF / image → OCR via Tesseract
  - Any extraction with messy output → optional LLM normalization (future)

The router does NOT replace the existing PDFExtractor — it wraps it
and adds routing for new file types (CSV, XLSX, images).
"""

import os
import pathlib

from extraction.schema import ExtractionResult


def classify_file(file_path: str) -> str:
    """
    Return a file-type key based on extension.
    Keys: 'csv', 'xlsx', 'pdf', 'image', 'unknown'
    """
    ext = pathlib.Path(file_path).suffix.lower()
    if ext == ".csv":
        return "csv"
    if ext == ".xlsx":
        return "xlsx"
    if ext == ".pdf":
        return "pdf"
    if ext in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}:
        return "image"
    return "unknown"


def pdf_has_text(file_path: str) -> bool:
    """
    Quick check: does the PDF contain any extractable text?
    Used to decide between text extraction and OCR.
    """
    try:
        import pdfplumber
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if text.strip():
                    return True
    except Exception:
        pass

    try:
        import PyPDF2
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            if reader.is_encrypted:
                from config import PDF_PASSWORD
                reader.decrypt(PDF_PASSWORD)
            for page in reader.pages:
                text = page.extract_text() or ""
                if text.strip():
                    return True
    except Exception:
        pass

    return False


def route_file(file_path: str) -> ExtractionResult:
    """
    Top-level entry point: given a file, extract data using
    the best available method and return an ExtractionResult.
    """
    if not os.path.isfile(file_path):
        return ExtractionResult.from_error(file_path, f"File not found: {file_path}")

    file_type = classify_file(file_path)

    if file_type == "csv":
        from extraction.structured_extractor import extract_csv
        return extract_csv(file_path)

    if file_type == "xlsx":
        from extraction.excel_extractor import extract_xlsx
        return extract_xlsx(file_path)

    if file_type == "image":
        from extraction.ocr_extractor import extract_image
        return extract_image(file_path)

    if file_type == "pdf":
        return _route_pdf(file_path)

    return ExtractionResult.from_error(
        file_path, f"Unsupported file type: {pathlib.Path(file_path).suffix}"
    )


def _route_pdf(file_path: str) -> ExtractionResult:
    """
    Route a PDF to either text extraction or OCR.
    Delegates to the existing PDFExtractor from auth_extractor.py
    which already handles the pdfplumber → PyPDF2 → OCR fallback chain.
    
    This wrapper exists so the router can report back the method used
    in a normalized ExtractionResult.
    """
    from extraction.pdf_text_extractor import extract_pdf
    return extract_pdf(file_path)
