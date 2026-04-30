"""
Auth PDF Downloader - Simple Dropbox PDF Retrieval Tool

Connects to Dropbox, lets you search/filter authorization PDFs,
downloads them, unlocks encrypted ones, and saves to a local folder.

Usage: python app.py
"""

import os
import sys
import pathlib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread
from datetime import datetime

# Load configuration
from config import (
    DROPBOX_APP_KEY,
    DROPBOX_APP_SECRET,
    DROPBOX_REFRESH_TOKEN,
    DROPBOX_ROOT_FOLDER,
    PDF_PASSWORD,
    APP_DIR,
)
from integrations.dropbox_service import DropboxService
from services.pdf_unlock_service import PdfUnlockService


class AuthDownloaderApp:
    """Simple app to download and unlock auth PDFs from Dropbox."""

    def __init__(self, root):
        self.root = root
        self.root.title("Auth PDF Downloader")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        # State
        self.dropbox_service = None
        self.all_files = []        # All files from Dropbox listing
        self.filtered_files = []   # After keyword filter
        self.is_busy = False

        self._build_ui()
        self._auto_connect()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        """Build the simple interface."""
        # --- Top: Connection status ---
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="Not connected")
        ttk.Label(top_frame, text="Dropbox:").pack(side=tk.LEFT)
        ttk.Label(top_frame, textvariable=self.status_var, foreground="gray").pack(side=tk.LEFT, padx=(5, 15))
        ttk.Button(top_frame, text="Connect", command=self._connect).pack(side=tk.LEFT)

        # --- Filter section ---
        filter_frame = ttk.LabelFrame(self.root, text="Search / Filter", padding=10)
        filter_frame.pack(fill=tk.X, padx=10, pady=(5, 0))

        row1 = ttk.Frame(filter_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Keyword:").pack(side=tk.LEFT)
        self.keyword_var = tk.StringVar()
        keyword_entry = ttk.Entry(row1, textvariable=self.keyword_var, width=40)
        keyword_entry.pack(side=tk.LEFT, padx=(5, 15))
        keyword_entry.bind("<Return>", lambda e: self._list_files())

        ttk.Label(row1, text="Modified after:").pack(side=tk.LEFT)
        self.date_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.date_var, width=12).pack(side=tk.LEFT, padx=(5, 15))
        ttk.Label(row1, text="(YYYY-MM-DD)", foreground="gray").pack(side=tk.LEFT)

        row2 = ttk.Frame(filter_frame)
        row2.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(row2, text="Search Files", command=self._list_files).pack(side=tk.LEFT)
        self.file_count_var = tk.StringVar(value="")
        ttk.Label(row2, textvariable=self.file_count_var, foreground="blue").pack(side=tk.LEFT, padx=10)

        # --- File list ---
        list_frame = ttk.LabelFrame(self.root, text="Files Found", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Treeview with scrollbar
        tree_container = ttk.Frame(list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_container, columns=("name", "date", "size"), show="headings", selectmode="extended")
        self.tree.heading("name", text="Filename")
        self.tree.heading("date", text="Modified")
        self.tree.heading("size", text="Size")
        self.tree.column("name", width=400)
        self.tree.column("date", width=120)
        self.tree.column("size", width=80)

        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Select all / none
        sel_frame = ttk.Frame(list_frame)
        sel_frame.pack(fill=tk.X, pady=(3, 0))
        ttk.Button(sel_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(sel_frame, text="Select None", command=self._select_none).pack(side=tk.LEFT)

        # --- Bottom: Download ---
        bottom_frame = ttk.Frame(self.root, padding=10)
        bottom_frame.pack(fill=tk.X)

        ttk.Label(bottom_frame, text="Save to:").pack(side=tk.LEFT)
        self.output_var = tk.StringVar(value=str(APP_DIR / "Downloaded Auths"))
        ttk.Entry(bottom_frame, textvariable=self.output_var, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="Browse...", command=self._browse_output).pack(side=tk.LEFT, padx=(0, 15))

        self.download_btn = ttk.Button(bottom_frame, text="Download & Unlock", command=self._download)
        self.download_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Progress
        self.progress_var = tk.StringVar(value="")
        ttk.Label(bottom_frame, textvariable=self.progress_var, foreground="gray").pack(side=tk.LEFT, padx=10)

    # ------------------------------------------------------------------
    # Dropbox connection
    # ------------------------------------------------------------------

    def _auto_connect(self):
        """Auto-connect if credentials are configured."""
        if DROPBOX_APP_KEY and DROPBOX_REFRESH_TOKEN:
            self._connect()

    def _connect(self):
        """Connect to Dropbox."""
        try:
            self.dropbox_service = DropboxService()
            self.dropbox_service.connect()
            self.status_var.set("Connected ✓")
        except Exception as e:
            self.status_var.set("Connection failed")
            messagebox.showerror("Dropbox Error", str(e))

    # ------------------------------------------------------------------
    # File listing
    # ------------------------------------------------------------------

    def _list_files(self):
        """List and filter files from Dropbox."""
        if not self.dropbox_service or not self.dropbox_service.is_connected:
            messagebox.showwarning("Not Connected", "Connect to Dropbox first.")
            return

        self.is_busy = True
        self.file_count_var.set("Searching...")
        Thread(target=self._list_files_thread, daemon=True).start()

    def _list_files_thread(self):
        """Background thread for file listing."""
        try:
            # Get all PDF files from Dropbox
            all_files = self.dropbox_service.list_files()
            # Filter to PDFs only
            self.all_files = [
                f for f in all_files
                if f.name.lower().endswith(".pdf")
            ]

            # Apply keyword filter
            keyword = self.keyword_var.get().strip().lower()
            date_filter = self.date_var.get().strip()

            self.filtered_files = self.all_files

            if keyword:
                # Support multiple keywords separated by comma
                keywords = [k.strip() for k in keyword.split(",") if k.strip()]
                self.filtered_files = [
                    f for f in self.filtered_files
                    if any(k in f.name.lower() for k in keywords)
                ]

            if date_filter:
                try:
                    filter_date = datetime.strptime(date_filter, "%Y-%m-%d")
                    self.filtered_files = [
                        f for f in self.filtered_files
                        if f.server_modified >= filter_date
                    ]
                except ValueError:
                    pass  # Ignore bad date format

            # Update UI on main thread
            self.root.after(0, self._populate_tree)

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.is_busy = False

    def _populate_tree(self):
        """Fill the file list tree."""
        self.tree.delete(*self.tree.get_children())

        for f in self.filtered_files:
            size_kb = f.size / 1024
            size_str = f"{size_kb:.0f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
            date_str = f.server_modified.strftime("%Y-%m-%d")
            self.tree.insert("", tk.END, values=(f.name, date_str, size_str))

        self.file_count_var.set(f"{len(self.filtered_files)} files found")

        # Auto-select all
        self._select_all()

    def _select_all(self):
        for item in self.tree.get_children():
            self.tree.selection_add(item)

    def _select_none(self):
        self.tree.selection_remove(*self.tree.get_children())

    # ------------------------------------------------------------------
    # Download
    # ------------------------------------------------------------------

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Choose download folder")
        if folder:
            self.output_var.set(folder)

    def _download(self):
        """Download selected files, unlock, save to output folder."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Files Selected", "Select files to download.")
            return

        output_folder = self.output_var.get().strip()
        if not output_folder:
            messagebox.showwarning("No Folder", "Choose an output folder.")
            return

        # Get the Dropbox file metadata for selected items
        selected_names = set()
        for item in selected:
            values = self.tree.item(item, "values")
            selected_names.add(values[0])  # filename

        files_to_download = [f for f in self.filtered_files if f.name in selected_names]

        self.download_btn.config(state=tk.DISABLED)
        self.progress_var.set(f"Downloading 0/{len(files_to_download)}...")
        Thread(target=self._download_thread, args=(files_to_download, output_folder), daemon=True).start()

    def _download_thread(self, files, output_folder):
        """Background download, unlock, and save."""
        os.makedirs(output_folder, exist_ok=True)
        unlock = PdfUnlockService()
        total = len(files)
        success = 0
        errors = []

        for i, file_meta in enumerate(files):
            self.root.after(0, lambda c=i+1, t=total, n=file_meta.name:
                           self.progress_var.set(f"Downloading {c}/{t}: {n}"))
            try:
                # Download to output folder
                local_path = self.dropbox_service.download_file(file_meta, dest_dir=output_folder)

                # Unlock if encrypted
                try:
                    unlocked_path, was_encrypted = unlock.unlock(local_path)
                    if was_encrypted:
                        # Replace encrypted file with unlocked version
                        import shutil
                        shutil.move(unlocked_path, local_path)
                except ValueError:
                    # Password failed - keep the encrypted file anyway
                    pass
                except Exception:
                    pass  # Non-fatal, keep the file as-is

                success += 1

            except Exception as e:
                errors.append(f"{file_meta.name}: {e}")

        # Done
        def _show_done():
            self.download_btn.config(state=tk.NORMAL)
            self.progress_var.set(f"Done! {success}/{total} files saved.")
            msg = f"Downloaded {success} of {total} files.\n\nSaved to:\n{output_folder}"
            if errors:
                msg += f"\n\nErrors ({len(errors)}):\n" + "\n".join(errors[:5])
            messagebox.showinfo("Download Complete", msg)
            # Open the folder
            os.startfile(output_folder)

        self.root.after(0, _show_done)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    app = AuthDownloaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
