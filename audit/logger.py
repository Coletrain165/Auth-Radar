"""
Auth Radar - Audit Logger

Tracks the lifecycle of every processed file:
  - source path (local or Dropbox)
  - download / extraction / review / upload timestamps
  - extraction status and method
  - upload status
  - errors encountered
  - reviewer edits

Stores records in a JSON-lines file (one JSON object per line)
for simplicity and append-only safety.  Can be swapped for a
database table later without changing the public API.
"""

import json
import os
from datetime import datetime, timezone
from typing import Optional

from config import AUDIT_DB_FILE


class AuditLogger:
    """Append-only audit trail for processed files."""

    def __init__(self, log_path: str = ""):
        self.log_path = log_path or str(AUDIT_DB_FILE)
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def log_extraction(
        self,
        source_file: str,
        extraction_method: str,
        status: str,
        fields_extracted: int = 0,
        warnings: Optional[list] = None,
        error: str = "",
    ) -> dict:
        """Record an extraction event."""
        record = self._base_record(source_file)
        record.update({
            "event": "extraction",
            "extraction_method": extraction_method,
            "status": status,
            "fields_extracted": fields_extracted,
            "warnings": warnings or [],
            "error": error,
        })
        self._append(record)
        return record

    def log_review(
        self,
        source_file: str,
        reviewer: str = "",
        edits: Optional[dict] = None,
    ) -> dict:
        """Record that a user reviewed (and optionally edited) a result."""
        record = self._base_record(source_file)
        record.update({
            "event": "review",
            "reviewer": reviewer,
            "edits": edits or {},
        })
        self._append(record)
        return record

    def log_upload(
        self,
        source_file: str,
        table_name: str,
        record_count: int = 1,
        status: str = "success",
        error: str = "",
    ) -> dict:
        """Record a database upload event."""
        record = self._base_record(source_file)
        record.update({
            "event": "upload",
            "table_name": table_name,
            "record_count": record_count,
            "status": status,
            "error": error,
        })
        self._append(record)
        return record

    def log_event(self, source_file: str, event: str, **details) -> dict:
        """Generic event logger for anything that doesn't fit above."""
        record = self._base_record(source_file)
        record["event"] = event
        record.update(details)
        self._append(record)
        return record

    # ------------------------------------------------------------------
    # Querying
    # ------------------------------------------------------------------

    def get_all(self) -> list[dict]:
        """Read the full audit log."""
        if not os.path.isfile(self.log_path):
            return []
        records = []
        with open(self.log_path, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line:
                    try:
                        records.append(json.loads(line))
                    except json.JSONDecodeError:
                        continue
        return records

    def get_for_file(self, source_file: str) -> list[dict]:
        """Return all events for a given source file."""
        return [r for r in self.get_all() if r.get("source_file") == source_file]

    def was_processed(self, source_file: str) -> bool:
        """Check if a file has already been successfully extracted."""
        for r in self.get_for_file(source_file):
            if r.get("event") == "extraction" and r.get("status") == "ready_for_review":
                return True
        return False

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    @staticmethod
    def _base_record(source_file: str) -> dict:
        return {
            "source_file": source_file,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }

    def _append(self, record: dict):
        with open(self.log_path, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, default=str) + "\n")
