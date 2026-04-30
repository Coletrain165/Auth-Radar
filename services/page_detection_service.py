"""
Auth Radar - Page Detection Service

Identifies the page in a PDF most likely to contain the Treatment Authorization Form.
Renders candidate pages to high-resolution images for OpenAI extraction.
"""

import logging
import os
import tempfile
from pathlib import Path

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)

# Keywords that indicate an authorization form page
AUTH_PAGE_KEYWORDS = [
    "TREATMENT AUTHORIZATION FORM",
    "Auth #",
    "Date Approved",
    "Date Auth. Expire",
    "MEMBER (Participant)",
    "Participant's Name",
    "Participant ID",
]


class PageDetectionService:
    """Detects the authorization page and renders it to a high-res image."""

    # Render at 300 DPI for good OCR/vision quality
    DEFAULT_DPI = 300

    def __init__(self, dpi: int = 0):
        self.dpi = dpi or self.DEFAULT_DPI
        # PyMuPDF zoom factor: 72 DPI is default, so zoom = target_dpi / 72
        self.zoom = self.dpi / 72.0

    def detect_auth_page(self, pdf_path: str) -> int | None:
        """
        Find the page number (0-indexed) most likely containing the authorization form.

        Strategy:
        1. Try text extraction on each page and look for keywords.
        2. Return the page with the highest keyword match score.
        3. If no text is available (scanned PDF), return None to signal
           that image-based detection should be used.

        Returns:
            0-indexed page number, or None if no text-based detection was possible.
        """
        doc = fitz.open(pdf_path)
        best_page = None
        best_score = 0
        any_text_found = False

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text") or ""

            if text.strip():
                any_text_found = True
                score = self._score_page(text)
                if score > best_score:
                    best_score = score
                    best_page = page_num

        doc.close()

        if not any_text_found:
            logger.info("No extractable text found in PDF (likely scanned): %s", os.path.basename(pdf_path))
            return None

        if best_page is not None and best_score >= 2:
            logger.info(
                "Auth page detected: page %d (score=%d) in %s",
                best_page + 1, best_score, os.path.basename(pdf_path),
            )
            return best_page

        # Fallback: if only 1 keyword matched, still use best page
        if best_page is not None:
            logger.info(
                "Weak auth page match: page %d (score=%d) in %s",
                best_page + 1, best_score, os.path.basename(pdf_path),
            )
            return best_page

        logger.warning("No auth page detected in %s", os.path.basename(pdf_path))
        return None

    def get_candidate_pages(self, pdf_path: str, max_pages: int = 3) -> list[int]:
        """
        Return a list of candidate page numbers (0-indexed) to send for extraction.

        If text detection finds a clear winner, returns just that page.
        For scanned PDFs with no text, returns the first `max_pages` pages.
        """
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        doc.close()

        detected = self.detect_auth_page(pdf_path)

        if detected is not None:
            return [detected]

        # Scanned PDF fallback: send first N pages (auth form is usually near the front)
        candidates = list(range(min(max_pages, total_pages)))
        logger.info(
            "Scanned PDF - sending first %d pages for extraction: %s",
            len(candidates), os.path.basename(pdf_path),
        )
        return candidates

    def render_page_to_image(self, pdf_path: str, page_num: int) -> str:
        """
        Render a single PDF page to a high-resolution PNG image.

        Args:
            pdf_path: Path to the (unlocked) PDF.
            page_num: 0-indexed page number.

        Returns:
            Path to the temporary PNG file. Caller is responsible for cleanup.
        """
        doc = fitz.open(pdf_path)
        page = doc[page_num]

        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        tmp = tempfile.NamedTemporaryFile(
            suffix=".png", prefix=f"auth_page{page_num + 1}_", delete=False
        )
        tmp_path = tmp.name
        tmp.close()

        pix.save(tmp_path)
        doc.close()

        logger.info(
            "Rendered page %d to image (%dx%d px): %s",
            page_num + 1, pix.width, pix.height, os.path.basename(tmp_path),
        )
        return tmp_path

    def render_pages_to_images(self, pdf_path: str, page_nums: list[int]) -> list[tuple[int, str]]:
        """
        Render multiple pages to images.

        Returns:
            List of (page_num, image_path) tuples.
        """
        results = []
        for page_num in page_nums:
            img_path = self.render_page_to_image(pdf_path, page_num)
            results.append((page_num, img_path))
        return results

    @staticmethod
    def cleanup_images(image_paths: list[str]):
        """Remove temporary image files."""
        for path in image_paths:
            try:
                if path and os.path.isfile(path):
                    os.unlink(path)
            except OSError as e:
                logger.warning("Could not clean up image %s: %s", path, e)

    def _score_page(self, text: str) -> int:
        """Score a page's text against auth form keywords."""
        text_upper = text.upper()
        score = 0
        for keyword in AUTH_PAGE_KEYWORDS:
            if keyword.upper() in text_upper:
                score += 1
        return score
