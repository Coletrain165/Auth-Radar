"""
Auth Radar - Structured (CSV) Extractor

Parses CSV files that contain authorization data in tabular form.
Returns an ExtractionResult per row (or a single result for the whole file).
"""

import os
import csv
import io

from extraction.schema import ExtractionResult

# Column-name aliases we look for in CSV headers.
# Maps canonical field name → list of possible CSV column headers.
_HEADER_MAP = {
    "Patient Name": [
        "patient name", "participant name", "member name", "name",
        "patient_name", "participant_name",
    ],
    "Auth #": [
        "auth #", "auth number", "authorization number", "auth_number",
        "auth#", "authorization #", "auth_num",
    ],
    "Date Approved": [
        "date approved", "approved date", "date_approved", "start date",
        "effective date", "start_date",
    ],
    "Date Auth Expire": [
        "date auth expire", "expiration date", "expire date",
        "date_auth_expire", "end date", "end_date", "expiry",
    ],
    "Patient ID": [
        "patient id", "participant id", "member id", "patient_id",
        "participant_id", "member_id", "policy #", "policy number",
    ],
    "Service_Type_Identifier": [
        "service type", "service_type_identifier", "service type identifier",
        "service_type", "type",
    ],
}


def _match_header(col_name: str) -> str | None:
    """Return the canonical field name if *col_name* matches any alias."""
    lower = col_name.strip().lower()
    for field, aliases in _HEADER_MAP.items():
        if lower in aliases:
            return field
    return None


def extract_csv(file_path: str) -> ExtractionResult:
    """
    Parse a CSV and return ONE ExtractionResult per data row.
    If the file has multiple rows, we return a result for the first row;
    callers that need batch processing can use extract_csv_rows() instead.
    """
    rows = extract_csv_rows(file_path)
    if not rows:
        return ExtractionResult.from_error(file_path, "CSV contained no extractable rows")
    return rows[0]


def extract_csv_rows(file_path: str) -> list[ExtractionResult]:
    """Parse every data row in the CSV into its own ExtractionResult."""
    file_path = str(file_path)
    if not os.path.isfile(file_path):
        return [ExtractionResult.from_error(file_path, "File not found")]

    try:
        with open(file_path, "r", encoding="utf-8-sig") as fh:
            sample = fh.read(4096)
        dialect = csv.Sniffer().sniff(sample)
    except csv.Error:
        dialect = csv.excel  # fallback

    results = []
    try:
        with open(file_path, "r", encoding="utf-8-sig") as fh:
            reader = csv.DictReader(fh, dialect=dialect)
            if reader.fieldnames is None:
                return [ExtractionResult.from_error(file_path, "CSV has no header row")]

            # Build mapping: canonical field → CSV column name
            col_map = {}
            for col in reader.fieldnames:
                canonical = _match_header(col)
                if canonical:
                    col_map[canonical] = col

            if not col_map:
                return [
                    ExtractionResult.from_error(
                        file_path,
                        f"No recognized columns. Found headers: {reader.fieldnames}",
                    )
                ]

            for row_num, row in enumerate(reader, start=2):
                fields = {}
                for canonical, csv_col in col_map.items():
                    val = (row.get(csv_col) or "").strip()
                    if val:
                        fields[canonical] = val

                results.append(
                    ExtractionResult(
                        source_file=file_path,
                        document_type="csv",
                        fields=fields,
                        extraction_method="csv_parse",
                        status="ready_for_review" if fields else "error",
                        error="" if fields else f"Row {row_num} was empty",
                        warnings=[],
                    )
                )

    except Exception as e:
        return [ExtractionResult.from_error(file_path, f"CSV parse error: {e}")]

    return results
