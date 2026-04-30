"""
Auth Radar - Auth Page Extraction Pipeline

Orchestrates the PDF processing flow:
  DropboxService (existing)
  → PdfUnlockService  (decrypt protected PDFs)
  → PageDetectionService  (find the Treatment Authorization Form page)
  → Save auth page as high-res PNG image

The output is a folder of auth page images (one per PDF) that can be
dragged into ChatGPT for manual data extraction.
"""

import logging
import os
import sys
import shutil
import pathlib
from datetime import datetime

from config import APP_DIR, PDF_PASSWORD

from services.pdf_unlock_service import PdfUnlockService
from services.page_detection_service import PageDetectionService

logger = logging.getLogger(__name__)


class AuthPagePipeline:
    """
    Unlocks PDFs and extracts the authorization page as a high-res image.

    Usage:
        pipeline = AuthPagePipeline()
        output_folder = pipeline.process_folder("/path/to/pdfs", "/path/to/output")
    """

    def __init__(self, progress_callback=None, log_callback=None):
        """
        Args:
            progress_callback: Optional callable(current, total, filename)
            log_callback: Optional callable(message) for status logging.
        """
        self.progress_callback = progress_callback
        self.log_callback = log_callback

        # Initialize services
        self.unlock_service = PdfUnlockService()
        self.page_detection = PageDetectionService()

    def log(self, msg: str):
        """Log a message to both Python logging and the optional UI callback."""
        logger.info(msg)
        if self.log_callback:
            self.log_callback(msg)

    def process_folder(self, folder_path: str, output_folder: str) -> list[dict]:
        """
        Process all PDF files: unlock, detect auth page, save as image.

        Args:
            folder_path: Path to folder containing PDF files.
            output_folder: Path where auth page images will be saved.

        Returns:
            List of result dicts (one per PDF) with status info.
        """
        folder = pathlib.Path(folder_path)
        pdf_files = sorted(folder.glob("*.pdf"))

        if not pdf_files:
            self.log("No PDF files found in folder.")
            return []

        # Create output folder
        out_path = pathlib.Path(output_folder)
        out_path.mkdir(parents=True, exist_ok=True)

        self.log(f"Found {len(pdf_files)} PDF files to process.")
        self.log(f"Auth page images will be saved to: {output_folder}")
        results = []

        for i, pdf_file in enumerate(pdf_files):
            if self.progress_callback:
                self.progress_callback(i + 1, len(pdf_files), pdf_file.name)

            result = self.process_single_pdf(str(pdf_file), output_folder)
            results.append(result)

        # Summary
        success_count = sum(1 for r in results if r.get("status") == "success")
        error_count = sum(1 for r in results if r.get("status") == "error")
        self.log(f"\nDone! {success_count} auth page images saved, {error_count} errors.")

        return results

    def process_single_pdf(self, pdf_path: str, output_folder: str) -> dict:
        """
        Process a single PDF: unlock, render ALL pages as images.

        Args:
            pdf_path: Path to the PDF file.
            output_folder: Where to save the output images.

        Returns:
            Dict with 'source_file', 'status', 'pages_saved', 'notes'.
        """
        filename = os.path.basename(pdf_path)
        stem = pathlib.Path(filename).stem  # filename without extension
        self.log(f"Processing: {filename}")

        unlocked_path = None

        try:
            # --- Step 1: Unlock PDF ---
            unlocked_path, was_encrypted = self.unlock_service.unlock(pdf_path)
            if was_encrypted:
                self.log(f"  Unlocked (was encrypted).")
            else:
                self.log(f"  Not encrypted.")

            # --- Step 2: Get total page count ---
            import fitz
            doc = fitz.open(unlocked_path)
            total_pages = len(doc)
            doc.close()
            self.log(f"  {total_pages} page(s) - rendering all...")

            # --- Step 3: Render ALL pages to high-res images ---
            # Create a subfolder per PDF to keep things organized
            pdf_output_folder = os.path.join(output_folder, stem)
            os.makedirs(pdf_output_folder, exist_ok=True)

            saved_pages = 0
            for page_num in range(total_pages):
                tmp_image_path = self.page_detection.render_page_to_image(
                    unlocked_path, page_num
                )
                output_image_name = f"page_{page_num + 1}.png"
                final_image_path = os.path.join(pdf_output_folder, output_image_name)
                shutil.move(tmp_image_path, final_image_path)
                saved_pages += 1

            self.log(f"  Saved {saved_pages} page images to: {stem}/")

            return {
                "source_file": filename,
                "status": "success",
                "output_folder": pdf_output_folder,
                "pages_saved": saved_pages,
                "notes": "",
            }

        except Exception as e:
            self.log(f"  ERROR: {e}")
            return {
                "source_file": filename,
                "status": "error",
                "output_image": "",
                "page_number": 0,
                "notes": str(e),
            }

        finally:
            # Cleanup unlocked temp file
            if unlocked_path and unlocked_path != pdf_path:
                PdfUnlockService.cleanup(unlocked_path)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    """Run the pipeline from the command line."""
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    parser = argparse.ArgumentParser(
        description="Auth Radar - Extract auth pages as images from PDFs"
    )
    parser.add_argument(
        "input_folder",
        help="Folder containing PDF files to process.",
    )
    parser.add_argument(
        "-o", "--output",
        default="Auth_Pages",
        help="Output folder for auth page images (default: Auth_Pages).",
    )
    args = parser.parse_args()

    if not os.path.isdir(args.input_folder):
        print(f"Error: Input folder does not exist: {args.input_folder}")
        sys.exit(1)

    pipeline = AuthPagePipeline(log_callback=print)
    results = pipeline.process_folder(args.input_folder, args.output)

    if results:
        success = sum(1 for r in results if r["status"] == "success")
        print(f"\n{success} auth page images saved to: {os.path.abspath(args.output)}")
        print("You can now drag these images into ChatGPT for extraction.")


if __name__ == "__main__":
    main()

