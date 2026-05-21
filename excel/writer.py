# excel/writer.py
from datetime import datetime
from pathlib import Path

import openpyxl

# Maps CAP result camelCase keys → output_cols snake_case keys
_RESULT_KEY_MAP: dict[str, str] = {
    "fieldText":   "field_name",
    "tableId":     "table_id",
    "fieldId":     "field_id",
    "dataType":    "data_type",
    "notes":       "notes",
    "sampleValue": "sample_value",
    "matchScore":  "match_score",
    "matchSource": "match_source",
    "obligatory":  "obligatory",
    "verified":    "verify",
}


def write_results(
    input_path: Path,
    workbook: openpyxl.Workbook,
    results: list[dict],
    output_cols: dict[str, str],
    sheet_data: str = "IFマッピング定義",
) -> Path:
    """Write match results into workbook's data sheet, save as processed_*.xlsx.

    output_cols maps snake_case field names to Excel column letters,
    e.g. {"table_id": "W", "field_id": "X", ...}.
    """
    ws = workbook[sheet_data]
    result_map: dict[int, dict] = {r["rowIndex"]: r for r in results if "rowIndex" in r}

    for row_idx, result in result_map.items():
        for cap_key, col_key in _RESULT_KEY_MAP.items():
            if cap_key not in result or col_key not in output_cols:
                continue
            col_letter = output_cols[col_key]
            ws[f"{col_letter}{row_idx}"] = result[cap_key]

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = input_path.parent / f"processed_{ts}_{input_path.name}"
    workbook.save(out_path)
    return out_path
