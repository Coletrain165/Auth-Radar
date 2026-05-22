"""
Auth Radar - Dropbox Integration Service

Provides authenticated access to a Dropbox folder:
  - List files recursively
  - Filter by supported file types
  - Download files to a temp directory for processing
  - Optionally move processed files to a "processed" subfolder

Requires the `dropbox` package:  pip install dropbox
"""

import os
import tempfile
import pathlib
from datetime import datetime

try:
    import dropbox
    from dropbox.files import FileMetadata, FolderMetadata
    from dropbox.exceptions import ApiError, AuthError
    DROPBOX_AVAILABLE = True
except ImportError:
    DROPBOX_AVAILABLE = False

from config import (
    DROPBOX_APP_KEY,
    DROPBOX_APP_SECRET,
    DROPBOX_REFRESH_TOKEN,
    DROPBOX_ROOT_FOLDER,
    SUPPORTED_EXTENSIONS,
)


class DropboxService:
    """
    Wraps the Dropbox API for Auth Radar.

    Usage:
        svc = DropboxService()
        svc.connect()
        files = svc.list_supported_files("/Auth Forms")
        local_path = svc.download_file(files[0])
    """

    def __init__(
        self,
        app_key: str = "",
        app_secret: str = "",
        refresh_token: str = "",
        root_folder: str = "",
    ):
        self.app_key = app_key or DROPBOX_APP_KEY
        self.app_secret = app_secret or DROPBOX_APP_SECRET
        self.refresh_token = refresh_token or DROPBOX_REFRESH_TOKEN
        self.root_folder = root_folder or DROPBOX_ROOT_FOLDER
        self.dbx = None
        self._temp_dir = None

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def connect(self):
        """
        Authenticate with Dropbox using a long-lived refresh token.
        Raises RuntimeError if the dropbox SDK is missing or creds are empty.
        """
        if not DROPBOX_AVAILABLE:
            raise RuntimeError(
                "The 'dropbox' package is not installed. Run: pip install dropbox"
            )

        if not self.app_key or not self.refresh_token:
            raise RuntimeError(
                "Dropbox credentials are not configured. "
                "Set DROPBOX_APP_KEY, DROPBOX_APP_SECRET, and DROPBOX_REFRESH_TOKEN in your .env file."
            )

        self.dbx = dropbox.Dropbox(
            app_key=self.app_key,
            app_secret=self.app_secret,
            oauth2_refresh_token=self.refresh_token,
        )

        # Verify the connection
        try:
            account = self.dbx.users_get_current_account()
            print(f"[Dropbox] Connected as {account.name.display_name}")
        except AuthError as e:
            self.dbx = None
            raise RuntimeError(f"Dropbox authentication failed: {e}")

    @property
    def is_connected(self) -> bool:
        return self.dbx is not None

    # ------------------------------------------------------------------
    # Listing files
    # ------------------------------------------------------------------

    def list_files(self, folder_path: str = "", recursive: bool = True):
        """
        List all files under *folder_path* (defaults to root_folder).
        Returns a list of dropbox.files.FileMetadata objects.
        """
        self._ensure_connected()
        folder = folder_path or self.root_folder
        # Dropbox API expects "" for root, otherwise a path starting with /
        if folder and not folder.startswith("/"):
            folder = "/" + folder

        entries = []
        result = self.dbx.files_list_folder(folder, recursive=recursive)

        while True:
            for entry in result.entries:
                if isinstance(entry, FileMetadata):
                    entries.append(entry)
            if not result.has_more:
                break
            result = self.dbx.files_list_folder_continue(result.cursor)

        return entries

    def list_supported_files(self, folder_path: str = "", recursive: bool = True):
        """
        List only files whose extension is in SUPPORTED_EXTENSIONS.
        Returns list of FileMetadata.
        """
        all_files = self.list_files(folder_path, recursive=recursive)
        return [
            f
            for f in all_files
            if pathlib.Path(f.name).suffix.lower() in SUPPORTED_EXTENSIONS
        ]

    def list_folders(self, folder_path: str = "") -> list:
        """
        List immediate sub-folders under *folder_path* (defaults to root_folder).
        Returns a sorted list of path_display strings.
        """
        self._ensure_connected()
        folder = folder_path or self.root_folder
        if folder and not folder.startswith("/"):
            folder = "/" + folder

        entries = []
        result = self.dbx.files_list_folder(folder, recursive=False)
        while True:
            for entry in result.entries:
                if isinstance(entry, FolderMetadata):
                    entries.append(entry.path_display)
            if not result.has_more:
                break
            result = self.dbx.files_list_folder_continue(result.cursor)

        return sorted(entries)

    # ------------------------------------------------------------------
    # Downloading
    # ------------------------------------------------------------------

    def download_file(self, file_meta, dest_dir: str = "") -> str:
        """
        Download a single Dropbox file to dest_dir.
        file_meta may be a FileMetadata object or a Dropbox path string.
        The file is placed flat in dest_dir (no subfolder nesting).
        Returns the local file path.
        """
        self._ensure_connected()

        if not dest_dir:
            dest_dir = self._get_temp_dir()

        # Accept either a FileMetadata object or a plain path string
        if isinstance(file_meta, str):
            dbx_path = file_meta
            filename = dbx_path.split("/")[-1]
        else:
            dbx_path = file_meta.path_lower
            filename = file_meta.name

        os.makedirs(dest_dir, exist_ok=True)
        local_path = os.path.join(dest_dir, filename)
        self.dbx.files_download_to_file(local_path, dbx_path)
        return local_path

    def download_files(self, file_metas, dest_dir: str = "", progress_callback=None):
        """
        Download multiple files. Returns list of (dropbox_path, local_path) tuples.
        progress_callback(current, total, filename) is called per file.
        """
        dest = dest_dir or self._get_temp_dir()
        results = []
        total = len(file_metas)

        for i, meta in enumerate(file_metas):
            if progress_callback:
                progress_callback(i + 1, total, meta.name)
            local = self.download_file(meta, dest_dir=dest)
            results.append((meta.path_display, local))

        return results

    # ------------------------------------------------------------------
    # Post-processing helpers
    # ------------------------------------------------------------------

    def move_to_processed(self, file_meta, processed_folder: str = ""):
        """
        Move a file into a 'processed' subfolder within its parent directory.
        Useful for marking files as done.
        """
        self._ensure_connected()
        parent = "/".join(file_meta.path_display.split("/")[:-1])
        dest_folder = processed_folder or f"{parent}/processed"
        dest_path = f"{dest_folder}/{file_meta.name}"

        try:
            # Create the folder if it doesn't exist (Dropbox creates on move)
            self.dbx.files_move_v2(file_meta.path_lower, dest_path.lower())
            return dest_path
        except ApiError as e:
            print(f"[Dropbox] Could not move {file_meta.name}: {e}")
            return None

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _ensure_connected(self):
        if not self.is_connected:
            raise RuntimeError("Not connected to Dropbox. Call connect() first.")

    def _get_temp_dir(self) -> str:
        if self._temp_dir is None:
            self._temp_dir = tempfile.mkdtemp(prefix="auth_radar_")
        return self._temp_dir

    def cleanup_temp(self):
        """Remove the temporary download directory."""
        if self._temp_dir and os.path.isdir(self._temp_dir):
            import shutil
            shutil.rmtree(self._temp_dir, ignore_errors=True)
            self._temp_dir = None


# ---------------------------------------------------------------------------
# OAuth helper — one-time setup to get a refresh token
# ---------------------------------------------------------------------------

def run_dropbox_oauth_flow(app_key: str = ""):
    """
    Interactive helper to obtain a Dropbox refresh token.
    Run this once from the command line, then paste the token into .env.

    Usage:
        python -c "from integrations.dropbox_service import run_dropbox_oauth_flow; run_dropbox_oauth_flow()"
    """
    key = app_key or DROPBOX_APP_KEY
    if not key:
        key = input("Enter your Dropbox App Key: ").strip()

    if not DROPBOX_AVAILABLE:
        print("Install the dropbox package first: pip install dropbox")
        return

    auth_flow = dropbox.DropboxOAuth2FlowNoRedirect(
        key,
        use_pkce=True,
        token_access_type="offline",
    )
    authorize_url = auth_flow.start()
    print(f"\n1. Go to: {authorize_url}")
    print("2. Click 'Allow' (you may need to log in first)")
    print("3. Copy the authorization code.\n")
    auth_code = input("Enter the authorization code: ").strip()

    try:
        oauth_result = auth_flow.finish(auth_code)
        print(f"\nRefresh token: {oauth_result.refresh_token}")
        print("\nAdd this to your .env file as:")
        print(f"DROPBOX_REFRESH_TOKEN={oauth_result.refresh_token}")
    except Exception as e:
        print(f"Error: {e}")
