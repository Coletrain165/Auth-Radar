"""
Auth Radar - PDF Text Extractor

Wraps the existing PDFExtractor.process_pdf() logic and returns
a normalized ExtractionResult.  This is intentionally a thin adapter —
all the battle-tested regex, OCR fallback, and date-cleaning logic
stays in auth_extractor.PDFExtractor until we refactor it piece by piece.
"""

import os
import sys
import pathlib

# Add project root to path so auth_extractor can be imported
_project_root = str(pathlib.Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from extraction.schema import ExtractionResult


def _get_extractor():
    """
    Lazily import and instantiate the existing PDFExtractor.
    We import late to avoid circular-import issues and so the heavy
    Tesseract/pdfplumber setup only happens when actually needed.
    """
    from auth_extractor import PDFExtractor
    return PDFExtractor()


# Module-level cache so we don't reinitialize Tesseract every call
_extractor_instance = None


def _extractor():
    global _extractor_instance
    if _extractor_instance is None:
        _extractor_instance = _get_extractor()
    return _extractor_instance


def extract_pdf(file_path: str) -> ExtractionResult:
    """
    Extract data from a PDF using the existing PDFExtractor pipeline
    (pdfplumber → PyPDF2 → OCR fallback → regex + ML).

    Returns an ExtractionResult with normalized fields.
    """
    file_path = str(file_path)
    if not os.path.isfile(file_path):
        return ExtractionResult.from_error(file_path, "File not found")

    try:
        ext = _extractor()
        raw = ext.process_pdf(file_path)
    except Exception as e:
        return ExtractionResult.from_error(file_path, str(e))

    # Map the raw dict returned by process_pdf into our schema
    from config import FIELDS

    fields = {}
    confidence = {}
    warnings = []

    for f in FIELDS:
        val = raw.get(f)
        if val:
            fields[f] = val
        conf = raw.get(f"{f}_confidence")
        if conf is not None:
            confidence[f] = conf
        method = raw.get(f"{f}_method")
        if method and method.startswith("fallback"):
            warnings.append(f"{f} found via fallback ({method})")

    if raw.get("error"):
        warnings.append(raw["error"])

    extraction_method = raw.get("extraction_method", "unknown")
    doc_type = "pdf_ocr" if "ocr" in extraction_method else "pdf_text"

    return ExtractionResult(
        source_file=file_path,
        document_type=doc_type,
        fields=fields,
        warnings=warnings,
        confidence=confidence,
        raw_text=raw.get("raw_text_preview", ""),
        extraction_method=extraction_method,
        status="error" if raw.get("error") else "ready_for_review",
        page_number=raw.get("auth_page", 0),
    )
