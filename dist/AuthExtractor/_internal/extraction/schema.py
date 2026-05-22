"""
Auth Radar - Common Extraction Result Schema

All extractors return an ExtractionResult so downstream code
(review UI, database upload, audit logger) can treat every
source file identically regardless of extraction method.
"""

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ExtractionResult:
    """Normalized output from any extraction method."""

    source_file: str = ""
    document_type: str = ""        # e.g. "pdf_text", "pdf_ocr", "csv", "xlsx", "image_ocr"
    fields: dict = field(default_factory=dict)    # {"Patient Name": "...", "Auth #": "...", ...}
    warnings: list = field(default_factory=list)  # human-readable warnings
    confidence: dict = field(default_factory=dict) # {"Patient Name": 0.95, ...}
    raw_text: str = ""
    extraction_method: str = ""    # "pdfplumber", "ocr", "csv_parse", "excel_parse", ...
    status: str = "ready_for_review"
    error: str = ""
    page_number: int = 0

    # ----- helpers -----

    @property
    def is_valid(self) -> bool:
        """Has at least one non-empty field and no fatal error."""
        return bool(self.fields) and not self.error

    def to_dict(self) -> dict:
        return {
            "source_file": self.source_file,
            "document_type": self.document_type,
            "fields": self.fields,
            "warnings": self.warnings,
            "confidence": self.confidence,
            "raw_text": self.raw_text,
            "extraction_method": self.extraction_method,
            "status": self.status,
            "error": self.error,
            "page_number": self.page_number,
        }

    @classmethod
    def from_error(cls, source_file: str, error: str) -> "ExtractionResult":
        return cls(
            source_file=source_file,
            status="error",
            error=error,
        )
