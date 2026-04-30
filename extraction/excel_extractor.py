"""
Auth Radar - Excel (XLSX) Extractor

Parses .xlsx files that contain authorization data in tabular form.
Uses the same column-alias logic as the CSV extractor.
"""

import os
from extraction.schema import ExtractionResult

# Reuse the header-matching logic from the CSV extractor
from extraction.structured_extractor import _match_header


def extract_xlsx(file_path: str) -> ExtractionResult:
    """
    Parse an XLSX and return the first data row as an ExtractionResult.
    For multi-row results, use extract_xlsx_rows().
    """
    rows = extract_xlsx_rows(file_path)
    if not rows:
        return ExtractionResult.from_error(file_path, "XLSX contained no extractable rows")
    return rows[0]


def extract_xlsx_rows(file_path: str) -> list[ExtractionResult]:
    """Parse every data row in the first sheet into its own ExtractionResult."""
    file_path = str(file_path)
    if not os.path.isfile(file_path):
        return [ExtractionResult.from_error(file_path, "File not found")]

    try:
        import openpyxl
    except ImportError:
        return [ExtractionResult.from_error(file_path, "openpyxl not installed")]

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        if ws is None:
            return [ExtractionResult.from_error(file_path, "XLSX has no active sheet")]

        rows_iter = ws.iter_rows(values_only=True)

        # First row = header
        header_row = next(rows_iter, None)
        if header_row is None:
            wb.close()
            return [ExtractionResult.from_error(file_path, "XLSX has no header row")]

        headers = [str(h).strip() if h else "" for h in header_row]

        # Build mapping: canonical field → column index
        col_map: dict[str, int] = {}
        for idx, col_name in enumerate(headers):
            canonical = _match_header(col_name)
            if canonical:
                col_map[canonical] = idx

        if not col_map:
            wb.close()
            return [
                ExtractionResult.from_error(
                    file_path,
                    f"No recognized columns. Found headers: {headers}",
                )
            ]

        results = []
        for row_num, row in enumerate(rows_iter, start=2):
            fields = {}
            for canonical, idx in col_map.items():
                if idx < len(row):
                    val = str(row[idx]).strip() if row[idx] is not None else ""
                    if val:
                        fields[canonical] = val

            results.append(
                ExtractionResult(
                    source_file=file_path,
                    document_type="xlsx",
                    fields=fields,
                    extraction_method="excel_parse",
                    status="ready_for_review" if fields else "error",
                    error="" if fields else f"Row {row_num} was empty",
                    warnings=[],
                )
            )

        wb.close()
        return results

    except Exception as e:
        return [ExtractionResult.from_error(file_path, f"XLSX parse error: {e}")]
