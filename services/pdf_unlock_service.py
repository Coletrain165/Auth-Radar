"""
Auth Radar - PDF Unlock Service

Handles decryption of password-protected PDFs using PyMuPDF (fitz).
Returns unlocked PDF bytes or saves to a temporary file.
"""

import logging
import tempfile
import os

import fitz  # PyMuPDF

from config import PDF_PASSWORD

logger = logging.getLogger(__name__)


class PdfUnlockService:
    """Unlocks encrypted PDFs using the configured password."""

    def __init__(self, password: str = ""):
        self.password = password or PDF_PASSWORD

    def unlock(self, pdf_path: str) -> tuple[str, bool]:
        """
        Unlock a PDF if it is encrypted.

        Args:
            pdf_path: Path to the (possibly encrypted) PDF file.

        Returns:
            Tuple of (path_to_unlocked_pdf, was_encrypted).
            If the PDF was not encrypted, returns the original path.
            If it was encrypted and successfully unlocked, returns a temp file path.

        Raises:
            ValueError: If the PDF is encrypted and the password fails.
            FileNotFoundError: If the input file does not exist.
        """
        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        doc = fitz.open(pdf_path)

        if not doc.is_encrypted:
            logger.info("PDF is not encrypted: %s", os.path.basename(pdf_path))
            doc.close()
            return pdf_path, False

        # Attempt to decrypt
        if not doc.authenticate(self.password):
            doc.close()
            raise ValueError(
                f"Failed to unlock PDF with configured password: {os.path.basename(pdf_path)}"
            )

        logger.info("PDF unlocked successfully: %s", os.path.basename(pdf_path))

        # Save unlocked copy to a temp file
        tmp = tempfile.NamedTemporaryFile(
            suffix=".pdf", prefix="unlocked_", delete=False
        )
        tmp_path = tmp.name
        tmp.close()

        # Save without encryption
        doc.save(tmp_path, garbage=3, deflate=True)
        doc.close()

        return tmp_path, True

    def unlock_to_bytes(self, pdf_path: str) -> tuple[bytes, bool]:
        """
        Unlock a PDF and return its bytes.

        Returns:
            Tuple of (pdf_bytes, was_encrypted).
        """
        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        doc = fitz.open(pdf_path)

        if not doc.is_encrypted:
            doc.close()
            with open(pdf_path, "rb") as f:
                return f.read(), False

        if not doc.authenticate(self.password):
            doc.close()
            raise ValueError(
                f"Failed to unlock PDF with configured password: {os.path.basename(pdf_path)}"
            )

        pdf_bytes = doc.tobytes(garbage=3, deflate=True)
        doc.close()
        return pdf_bytes, True

    @staticmethod
    def cleanup(path: str):
        """Remove a temporary unlocked PDF file."""
        try:
            if path and os.path.isfile(path) and "unlocked_" in os.path.basename(path):
                os.unlink(path)
                logger.debug("Cleaned up temp file: %s", path)
        except OSError as e:
            logger.warning("Could not clean up temp file %s: %s", path, e)
