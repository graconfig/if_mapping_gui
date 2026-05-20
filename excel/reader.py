# excel/reader.py
import openpyxl
from pathlib import Path
from dataclasses import dataclass

COLUMN_ALIASES: dict[str, str] = {
    "sourcefield": "sourceField",
    "field_name": "sourceField",
    "フィールド名": "sourceField",
    "sourcedesc": "sourceDesc",
    "description": "sourceDesc",
    "説明": "sourceDesc",
    "描述": "sourceDesc",
    "sourcetable": "sourceTable",
    "table_name": "sourceTable",
    "テーブル名": "sourceTable",
}


@dataclass
class InterfaceFieldInput:
    sourceField: str
    sourceDesc: str
    rowIndex: int
    sourceTable: str = ""

    def to_dict(self) -> dict:
        return {
            "sourceField": self.sourceField,
            "sourceDesc": self.sourceDesc,
            "sourceTable": self.sourceTable,
            "rowIndex": self.rowIndex,
        }


class ExcelReadError(Exception):
    pass


def read_fields(path: Path) -> tuple[list[InterfaceFieldInput], list[dict]]:
    """Return (fields, raw_rows) where raw_rows preserve original header names."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows:
        raise ExcelReadError(f"{path.name}: empty file")

    # Build column index map: canonical name -> column index
    original_headers = [str(c).strip() if c is not None else "" for c in all_rows[0]]
    col_map: dict[str, int] = {}
    for i, h in enumerate(original_headers):
        canonical = COLUMN_ALIASES.get(h.lower(), h.lower())
        col_map[canonical] = i

    if "sourceField" not in col_map:
        raise ExcelReadError(f"{path.name}: missing required column 'sourceField' (got: {original_headers})")
    if "sourceDesc" not in col_map:
        raise ExcelReadError(f"{path.name}: missing required column 'sourceDesc' (got: {original_headers})")

    fields: list[InterfaceFieldInput] = []
    raw_rows: list[dict] = []

    for row_offset, row in enumerate(all_rows[1:], start=2):
        sf = row[col_map["sourceField"]] if col_map["sourceField"] < len(row) else None
        if not sf:
            continue
        sd = row[col_map["sourceDesc"]] if col_map["sourceDesc"] < len(row) else None
        st_idx = col_map.get("sourceTable")
        st = str(row[st_idx]).strip() if st_idx is not None and st_idx < len(row) and row[st_idx] else ""

        fields.append(InterfaceFieldInput(
            sourceField=str(sf).strip(),
            sourceDesc=str(sd).strip() if sd else "",
            rowIndex=row_offset,
            sourceTable=st,
        ))
        raw_row = {original_headers[i]: (str(v).strip() if v is not None else "") for i, v in enumerate(row)}
        raw_row["rowIndex"] = row_offset
        raw_rows.append(raw_row)

    return fields, raw_rows
