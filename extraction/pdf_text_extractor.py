"""
Auth Radar - PDF Text Extractor

Primary extraction pipeline:
  1. PyMuPDF (fitz) — native text extraction with built-in encryption support.
     Handles password-protected PDFs directly without writing temp files.
  2. Tesseract OCR — fallback for scanned / image-only PDFs.
  3. PDFExtractor — field parsing (regex, name matching, date cleaning, fallbacks).

PyMuPDF is used as the primary extractor because:
  - It decrypts in-memory using the configured password (no temp file)
  - It is significantly faster than pdfplumber / PyPDF2
  - It produces clean per-page text suitable for the existing regex patterns
"""

import os
import sys
import pathlib
from datetime import datetime

# fitz (PyMuPDF) is imported lazily inside _extract_text_pymupdf so that
# the module loads cleanly even on machines where PyMuPDF is not yet installed.

# Add project root to path so auth_extractor can be imported
_project_root = str(pathlib.Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from config import FIELDS, PDF_PASSWORD, POPPLER_PATH
from extraction.schema import ExtractionResult


# ---------------------------------------------------------------------------
# Text extraction helpers
# ---------------------------------------------------------------------------

def _extract_text_pymupdf(pdf_path: str) -> tuple[list[str], bool]:
    """
    Extract per-page text with PyMuPDF.  Handles encryption natively.

    Returns:
        (page_texts, was_encrypted)
        page_texts is empty when the PDF contains only images (scanned).

    Raises:
        ValueError if the PDF is encrypted and the password fails.
    """
    import fitz  # lazy import — PyMuPDF

    doc = fitz.open(pdf_path)
    was_encrypted = doc.is_encrypted

    if was_encrypted:
        if not doc.authenticate(PDF_PASSWORD):
            doc.close()
            raise ValueError(
                f"Failed to decrypt PDF with configured password: {os.path.basename(pdf_path)}"
            )

    page_texts = []
    for page in doc:
        text = page.get_text("text")
        if text.strip():
            page_texts.append(text)

    doc.close()
    return page_texts, was_encrypted


def _extract_text_ocr(pdf_path: str) -> list[str]:
    """
    OCR fallback for scanned PDFs.  Dual-pass Tesseract:
      PSM 3  — full page layout (captures general structure)
      PSM 11 — sparse text (best for isolated form values like auth # and dates)
    Both passes are concatenated so all regex patterns can match.
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract
        import auth_extractor as _ae

        tesseract_path = _ae.find_tesseract()
        if not tesseract_path:
            return []
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

        poppler_path = str(POPPLER_PATH) if POPPLER_PATH and pathlib.Path(str(POPPLER_PATH)).exists() else None
        images = convert_from_path(
            pdf_path, poppler_path=poppler_path, dpi=400, userpw=PDF_PASSWORD
        )

        cfg_psm3 = "--psm 3 --oem 3 -c tessedit_do_invert=0"
        cfg_psm11 = "--psm 11 --oem 3 -c tessedit_do_invert=0"

        page_texts = []
        for img in images:
            text3 = pytesseract.image_to_string(img, config=cfg_psm3)
            text11 = pytesseract.image_to_string(img, config=cfg_psm11)
            combined = text3
            if text11.strip():
                combined = text3 + "\n--- PSM11 ---\n" + text11
            if combined.strip():
                page_texts.append(combined)

        return page_texts

    except Exception as e:
        print(f"OCR error on {os.path.basename(pdf_path)}: {e}")
        return []


# ---------------------------------------------------------------------------
# Module-level PDFExtractor cache (avoids re-initialising Tesseract every call)
# ---------------------------------------------------------------------------

_extractor_instance = None


def _extractor():
    global _extractor_instance
    if _extractor_instance is None:
        from auth_extractor import PDFExtractor
        _extractor_instance = PDFExtractor()
    return _extractor_instance


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def extract_pdf(file_path: str) -> ExtractionResult:
    """
    Extract authorization data from a PDF file.

    Flow:
      PyMuPDF text extraction (handles encrypted PDFs natively)
        → Tesseract OCR fallback (for scanned / image-only PDFs)
        → PDFExtractor field parsing (regex, date cleaning, name matching)

    Returns an ExtractionResult with normalized fields.
    """
    file_path = str(file_path)
    if not os.path.isfile(file_path):
        return ExtractionResult.from_error(file_path, "File not found")

    # ---- Step 1: extract text ------------------------------------------------
    page_texts: list[str] = []
    method = "none"

    try:
        page_texts, was_encrypted = _extract_text_pymupdf(file_path)
        if page_texts:
            method = "pymupdf_decrypted" if was_encrypted else "pymupdf"
    except ValueError as e:
        # Password authentication failed — hard error, cannot proceed
        return ExtractionResult.from_error(file_path, str(e))
    except Exception as e:
        print(f"PyMuPDF error on {os.path.basename(file_path)}: {e}")

    if not page_texts:
        page_texts = _extract_text_ocr(file_path)
        if page_texts:
            method = "ocr"

    if not page_texts:
        return ExtractionResult.from_error(
            file_path,
            "Could not extract text — PDF may be a scanned image without OCR support "
            "or is encrypted with an unknown password.",
        )

    # ---- Step 2: locate the auth form page and parse fields ------------------
    try:
        ext = _extractor()
    except Exception as e:
        return ExtractionResult.from_error(file_path, f"Could not initialise extractor: {e}")

    auth_text, all_text, page_num = ext.find_auth_page(page_texts)

    # Build a result dict in the same format process_pdf() produces so that
    # all downstream helper methods (validate_and_fix_dates, etc.) work as-is.
    result: dict = {
        "file": os.path.basename(file_path),
        "extracted_at": datetime.now().isoformat(),
        "extraction_method": method,
        "auth_page": page_num if page_num > 0 else "all",
    }
    for f in FIELDS:
        result[f] = None

    if not auth_text.strip():
        result["error"] = "Could not extract text from the authorization form page."
    else:
        result["raw_text_preview"] = auth_text[:800].replace("\n", " | ")

        # Smart regex (primary)
        for f, val in ext.extract_smart(auth_text).items():
            if val and not result[f]:
                result[f] = val
                result[f"{f}_method"] = "regex_smart"

        # Pattern fallback for any remaining empty fields
        for f in FIELDS:
            if not result[f]:
                val = ext.extract_field(auth_text, f)
                if val:
                    result[f] = val
                    result[f"{f}_method"] = "regex_pattern"

        # Date validation / cross-correction
        result = ext.validate_and_fix_dates(result)

    # ---- Step 3: map to ExtractionResult schema ------------------------------
    fields = {f: result[f] for f in FIELDS if result.get(f)}
    confidence = {f: result[f"{f}_confidence"] for f in FIELDS if result.get(f"{f}_confidence")}
    warnings = [result["error"]] if result.get("error") else []

    doc_type = "pdf_ocr" if method == "ocr" else "pdf_text"

    return ExtractionResult(
        source_file=file_path,
        document_type=doc_type,
        fields=fields,
        warnings=warnings,
        confidence=confidence,
        raw_text=result.get("raw_text_preview", ""),
        extraction_method=method,
        status="error" if result.get("error") else "ready_for_review",
        page_number=page_num,
    )
