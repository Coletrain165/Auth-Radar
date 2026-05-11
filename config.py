"""
Auth Radar - Centralized Configuration

Loads settings from environment variables (.env file) with sensible defaults.
Moves hardcoded credentials out of source code for security.
"""

import os
import sys
import pathlib

# Load .env file if python-dotenv is available
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


def get_app_dir():
    """Get the application directory - works for both script and frozen exe."""
    if getattr(sys, 'frozen', False):
        return pathlib.Path(sys.executable).parent
    else:
        return pathlib.Path(__file__).parent


APP_DIR = get_app_dir()

# ---------------------------------------------------------------------------
# Caspio API
# ---------------------------------------------------------------------------
CASPIO_ACCOUNT_ID = os.getenv("CASPIO_ACCOUNT_ID", "")
CASPIO_CLIENT_ID = os.getenv("CASPIO_CLIENT_ID", "")
CASPIO_CLIENT_SECRET = os.getenv("CASPIO_CLIENT_SECRET", "")
CASPIO_TABLE_NAME = os.getenv("CASPIO_TABLE_NAME", "a_Authorizations")

# ---------------------------------------------------------------------------
# Dropbox
# ---------------------------------------------------------------------------
DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET", "")
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN", "")
DROPBOX_ROOT_FOLDER = os.getenv("DROPBOX_ROOT_FOLDER", "")

# ---------------------------------------------------------------------------
# PDF / OCR
# ---------------------------------------------------------------------------
PDF_PASSWORD = os.getenv("PDF_PASSWORD", "")

# Poppler path for pdf2image (Windows)
_poppler_locations = [
    APP_DIR / "poppler" / "poppler-24.08.0" / "Library" / "bin",
    APP_DIR / "poppler" / "Library" / "bin",
    APP_DIR / "poppler",
]
POPPLER_PATH = None
for _loc in _poppler_locations:
    if _loc.exists() and (_loc / "pdftoppm.exe").exists():
        POPPLER_PATH = _loc
        break
if POPPLER_PATH is None:
    for _loc in _poppler_locations:
        if _loc.exists():
            POPPLER_PATH = _loc
            break
if POPPLER_PATH is None:
    POPPLER_PATH = _poppler_locations[0]

# ---------------------------------------------------------------------------
# Data paths
# ---------------------------------------------------------------------------
DATA_DIR = APP_DIR / "data"
PATIENT_NAMES_FILE = DATA_DIR / "patient_names.json"
GMAIL_ACCOUNTS_FILE = DATA_DIR / "gmail_accounts.json"
CREDENTIALS_FILE = APP_DIR / "credentials.json"
AUDIT_DB_FILE = DATA_DIR / "audit_log.json"

# ---------------------------------------------------------------------------
# Extraction fields
# ---------------------------------------------------------------------------
FIELDS = [
    "Patient Name",
    "Auth #",
    "Date Approved",
    "Date Auth Expire",
    "Patient ID",
    "Service_Type_Identifier",
]

# Supported file types for extraction
SUPPORTED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".csv", ".xlsx"}
