# tests/test_excel_reader.py
import pytest
import openpyxl
from pathlib import Path
from excel.reader import read_fields, InterfaceFieldInput, ExcelReadError

def _make_xlsx(tmp_path, headers: list, rows: list[list]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    path = tmp_path / "test.xlsx"
    wb.save(path)
    return path

def test_reads_standard_columns(tmp_path):
    path = _make_xlsx(tmp_path,
        ["sourceField", "sourceDesc", "sourceTable"],
        [["EKPO-MATNR", "品目コード", "EKPO"], ["EKKO-BUKRS", "会社コード", "EKKO"]],
    )
    fields, raw = read_fields(path)
    assert len(fields) == 2
    assert fields[0].sourceField == "EKPO-MATNR"
    assert fields[0].sourceDesc == "品目コード"
    assert fields[0].sourceTable == "EKPO"
    assert fields[0].rowIndex == 2
    assert fields[1].rowIndex == 3

def test_reads_alias_columns(tmp_path):
    path = _make_xlsx(tmp_path,
        ["field_name", "description"],
        [["MARA-MATNR", "Material No."]],
    )
    fields, _ = read_fields(path)
    assert fields[0].sourceField == "MARA-MATNR"
    assert fields[0].sourceDesc == "Material No."

def test_reads_japanese_alias_columns(tmp_path):
    path = _make_xlsx(tmp_path,
        ["フィールド名", "説明", "テーブル名"],
        [["MARC-WERKS", "プラント", "MARC"]],
    )
    fields, _ = read_fields(path)
    assert fields[0].sourceField == "MARC-WERKS"
    assert fields[0].sourceTable == "MARC"

def test_sourcetable_optional(tmp_path):
    path = _make_xlsx(tmp_path,
        ["sourceField", "sourceDesc"],
        [["MARA-MATNR", "Material"]],
    )
    fields, _ = read_fields(path)
    assert fields[0].sourceTable == ""

def test_skips_blank_rows(tmp_path):
    path = _make_xlsx(tmp_path,
        ["sourceField", "sourceDesc"],
        [["MATNR", "Material"], [None, None], ["WERKS", "Plant"]],
    )
    fields, _ = read_fields(path)
    assert len(fields) == 2
    assert fields[1].sourceField == "WERKS"

def test_raises_on_missing_sourcefield(tmp_path):
    path = _make_xlsx(tmp_path, ["sourceDesc"], [["Material"]])
    with pytest.raises(ExcelReadError, match="sourceField"):
        read_fields(path)

def test_raises_on_missing_sourcedesc(tmp_path):
    path = _make_xlsx(tmp_path, ["sourceField"], [["MATNR"]])
    with pytest.raises(ExcelReadError, match="sourceDesc"):
        read_fields(path)

def test_raw_rows_preserve_original_headers(tmp_path):
    path = _make_xlsx(tmp_path,
        ["フィールド名", "説明"],
        [["MATNR", "Material"]],
    )
    _, raw = read_fields(path)
    assert raw[0]["フィールド名"] == "MATNR"
    assert "rowIndex" in raw[0]
