# excel/writer.py
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

RESULT_COLS = ["targetField", "targetEntity", "matchType", "confidence", "aiReason"]


def write_results(input_path: Path, raw_rows: list[dict], results: list[dict]) -> Path:
    """Write match results to *_matched_YYYYMMDD_HHMMSS.xlsx beside the input file."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = input_path.parent / f"{input_path.stem}_matched_{ts}.xlsx"

    input_cols = [k for k in (raw_rows[0].keys() if raw_rows else []) if k != "rowIndex"]
    headers = input_cols + RESULT_COLS

    result_by_row: dict[int, dict] = {r.get("rowIndex", -1): r for r in results}

    wb = openpyxl.Workbook()
    ws = wb.active

    for col_i, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_i, value=h)
        cell.font = Font(bold=True)

    for row_i, raw in enumerate(raw_rows, start=2):
        row_idx = raw.get("rowIndex", -1)
        matched = result_by_row.get(row_idx, {})
        for col_i, col in enumerate(input_cols, start=1):
            ws.cell(row=row_i, column=col_i, value=raw.get(col, ""))
        offset = len(input_cols)
        for col_i, col in enumerate(RESULT_COLS, start=offset + 1):
            ws.cell(row=row_i, column=col_i, value=matched.get(col, ""))

    wb.save(out_path)
    return out_path
