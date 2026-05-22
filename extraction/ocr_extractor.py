"""
Auth Radar - OCR Extractor

Handles OCR for standalone image files (PNG, JPG, JPEG).
For scanned PDFs, the existing PDFExtractor already handles OCR internally;
this module is specifically for image-only inputs.
"""

import os
import pathlib

from extraction.schema import ExtractionResult


def extract_image(file_path: str) -> ExtractionResult:
    """
    Extract text from an image file using Tesseract OCR,
    then run the standard regex extraction on the result.
    """
    file_path = str(file_path)
    if not os.path.isfile(file_path):
        return ExtractionResult.from_error(file_path, "File not found")

    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return ExtractionResult.from_error(
            file_path, "OCR libraries not available (pytesseract / Pillow)"
        )

    # Find tesseract
    from auth_extractor import find_tesseract
    tesseract_path = find_tesseract()
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    try:
        from auth_extractor import preprocess_image_for_ocr

        img = Image.open(file_path)
        processed = preprocess_image_for_ocr(img)

        # Run OCR with two page-segmentation modes (same as PDF pipeline)
        config_psm3 = "--psm 3 --oem 3 -c tessedit_do_invert=0"
        config_psm11 = "--psm 11 --oem 3 -c tessedit_do_invert=0"

        text_psm3 = pytesseract.image_to_string(processed, config=config_psm3)
        text_psm11 = pytesseract.image_to_string(processed, config=config_psm11)

        combined_text = text_psm3
        if text_psm11.strip():
            combined_text += "\n--- PSM11 ---\n" + text_psm11

        if not combined_text.strip():
            return ExtractionResult.from_error(file_path, "OCR produced no text")

        # Use the existing smart extraction on the OCR text
        from auth_extractor import PDFExtractor
        ext = PDFExtractor()
        smart = ext.extract_smart(combined_text)

        from config import FIELDS
        fields = {f: smart[f] for f in FIELDS if smart.get(f)}

        return ExtractionResult(
            source_file=file_path,
            document_type="image_ocr",
            fields=fields,
            raw_text=combined_text[:800],
            extraction_method="tesseract_image",
            status="ready_for_review" if fields else "error",
            error="" if fields else "No fields extracted from image",
        )

    except Exception as e:
        return ExtractionResult.from_error(file_path, f"Image OCR failed: {e}")
