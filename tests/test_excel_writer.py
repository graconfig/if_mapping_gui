# tests/test_excel_writer.py
import openpyxl
import pytest
from pathlib import Path
from excel.writer import write_results

def test_output_filename_contains_matched_and_timestamp(tmp_path):
    input_path = tmp_path / "IF_MM_001.xlsx"
    raw_rows = [{"sourceField": "MATNR", "sourceDesc": "Material", "rowIndex": 2}]
    results = [{"rowIndex": 2, "targetField": "Material", "targetEntity": "C_PurchaseOrderItemTP",
                "matchType": "exact", "confidence": 1.0, "aiReason": ""}]
    out = write_results(input_path, raw_rows, results)
    assert out.stem.startswith("IF_MM_001_matched_")
    assert out.suffix == ".xlsx"
    assert out.parent == tmp_path

def test_output_has_all_input_columns_plus_result_columns(tmp_path):
    input_path = tmp_path / "test.xlsx"
    raw_rows = [{"sourceField": "MATNR", "sourceDesc": "Material", "rowIndex": 2}]
    results = [{"rowIndex": 2, "targetField": "Material", "targetEntity": "CDS_View",
                "matchType": "ai", "confidence": 0.9, "aiReason": "名前が一致"}]
    out = write_results(input_path, raw_rows, results)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    assert "sourceField" in headers
    assert "sourceDesc" in headers
    assert "targetField" in headers
    assert "targetEntity" in headers
    assert "matchType" in headers
    assert "confidence" in headers
    assert "aiReason" in headers

def test_output_data_row_values_are_correct(tmp_path):
    input_path = tmp_path / "test.xlsx"
    raw_rows = [{"sourceField": "MATNR", "sourceDesc": "Material", "rowIndex": 2}]
    results = [{"rowIndex": 2, "targetField": "Material", "targetEntity": "CDS_View",
                "matchType": "ai", "confidence": 0.9, "aiReason": "名前が一致"}]
    out = write_results(input_path, raw_rows, results)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    data = {headers[c]: ws.cell(2, c + 1).value for c in range(len(headers))}
    assert data["sourceField"] == "MATNR"
    assert data["targetField"] == "Material"
    assert data["matchType"] == "ai"
    assert abs(data["confidence"] - 0.9) < 0.001
    assert data["aiReason"] == "名前が一致"

def test_unmatched_row_has_empty_result_columns(tmp_path):
    input_path = tmp_path / "test.xlsx"
    raw_rows = [{"sourceField": "UNKNOWN", "sourceDesc": "?", "rowIndex": 2}]
    results = [{"rowIndex": 2, "targetField": "", "targetEntity": "",
                "matchType": "unmatched", "confidence": 0.0, "aiReason": ""}]
    out = write_results(input_path, raw_rows, results)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    match_type_col = headers.index("matchType") + 1
    assert ws.cell(2, match_type_col).value == "unmatched"
