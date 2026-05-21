# Excel Config & Direction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace hardcoded Excel column/sheet constants with a configurable two-direction system (normal vs SAP), expand `InterfaceFieldInput` to the full field set, and add a settings GUI for editing all column mappings.

**Architecture:** Config is loaded from `config.json` with deep-merge defaults; `read_fields` detects direction from a configurable cell; `write_results` receives `output_cols` from config. Settings frame gains a `CTkTabview` with editable column entries for both directions.

**Tech Stack:** Python 3.11+, customtkinter, openpyxl, pytest

---

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `config.py` | Modify | Add `EXCEL_DEFAULTS`, `_deep_merge()`, update `load()` |
| `excel/reader.py` | Modify | Expand dataclass; accept `excel_cfg`; detect direction; return `(fields, wb, direction)` |
| `excel/writer.py` | Modify | Accept `output_cols` dict; use `_RESULT_KEY_MAP` |
| `gui/frames/settings_frame.py` | Modify | Add scrollable Excel config section with CTkTabview |
| `gui/frames/match_frame.py` | Modify | Pass `excel_cfg`; store `direction` per file; pass `output_cols` to writer |
| `tests/test_config.py` | Modify | Add deep_merge and excel defaults tests |
| `tests/test_excel_reader.py` | Modify | Update to new 3-value return; add direction + new fields tests |
| `tests/test_excel_writer.py` | Modify | Update to pass `output_cols`; remove `OUTPUT_COLS` import |
| `tests/test_match_worker.py` | Modify | Update to new `match_worker` signature |

---

## Task 1: config.py — EXCEL_DEFAULTS and deep_merge

**Files:**
- Modify: `config.py`
- Modify: `tests/test_config.py`

- [ ] **Step 1: Add failing tests for deep_merge and excel defaults**

Append to `tests/test_config.py`:

```python
def test_load_returns_excel_defaults_when_no_file():
    cfg = config.load()
    assert "excel" in cfg
    assert cfg["excel"]["sheet_head"] == "対象IF"
    assert cfg["excel"]["sheet_data"] == "IFマッピング定義"
    assert cfg["excel"]["header_row"] == 6
    assert cfg["excel"]["start_row"] == 5
    assert cfg["excel"]["detection"]["col"] == "F"
    assert cfg["excel"]["detection"]["row"] == 6
    assert cfg["excel"]["detection"]["keyword"] == "SAP"
    assert cfg["excel"]["directions"]["normal"]["input_row_cols"]["field_name"] == "C"
    assert cfg["excel"]["directions"]["sap"]["input_row_cols"]["field_name"] == "S"

def test_load_deep_merges_partial_excel_config():
    config.CONFIG_PATH.write_text(json.dumps({
        "excel": {"sheet_head": "MySheet", "directions": {"normal": {"input_row_cols": {"field_name": "B"}}}}
    }))
    cfg = config.load()
    assert cfg["excel"]["sheet_head"] == "MySheet"        # overridden
    assert cfg["excel"]["sheet_data"] == "IFマッピング定義"  # default preserved
    assert cfg["excel"]["directions"]["normal"]["input_row_cols"]["field_name"] == "B"  # overridden
    assert cfg["excel"]["directions"]["normal"]["input_row_cols"]["field_text"] == "L"  # default preserved
    assert "sap" in cfg["excel"]["directions"]             # other direction preserved
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_config.py -v 2>&1 | tail -20
```

Expected: `FAILED` on the two new tests.

- [ ] **Step 3: Implement EXCEL_DEFAULTS, _deep_merge, and update load()**

Replace `config.py` entirely:

```python
# config.py
import json
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.json"

DEFAULTS: dict = {
    "server_url": "http://localhost:4004",
    "provider": "claude",
    "language": "ja",
    "last_input_dir": "",
}

EXCEL_DEFAULTS: dict = {
    "sheet_head": "対象IF",
    "sheet_data": "IFマッピング定義",
    "header_row": 6,
    "start_row": 5,
    "detection": {"col": "F", "row": 6, "keyword": "SAP"},
    "directions": {
        "normal": {
            "input_header_cols": {"module": "D", "if_name": "C", "if_desc": "E"},
            "input_row_cols": {
                "field_name": "C", "is_append": "D", "key_flag": "E", "obligatory": "F",
                "data_type": "I", "table_id": "G", "field_id": "H",
                "length_total": "J", "length_dec": "K",
                "field_text": "L", "sample_value": "N", "remark": "M", "verify": "AF",
            },
            "output_cols": {
                "field_name": "S", "is_append": "T", "key_flag": "U", "obligatory": "V",
                "table_id": "W", "field_id": "X", "data_type": "Y",
                "length_total": "Z", "length_dec": "AA",
                "notes": "AB", "sample_value": "AC", "match_score": "AE",
                "match_source": "AF", "verify": "AG",
            },
        },
        "sap": {
            "input_header_cols": {"module": "D", "if_name": "C", "if_desc": "E"},
            "input_row_cols": {
                "field_name": "S", "is_append": "T", "key_flag": "U", "obligatory": "V",
                "table_id": "W", "field_id": "X", "data_type": "Y",
                "length_total": "Z", "length_dec": "AA",
                "field_text": "AB", "sample_value": "AC", "remark": "AE", "verify": "AF",
            },
            "output_cols": {
                "field_name": "C", "is_append": "D", "key_flag": "E", "obligatory": "F",
                "table_id": "G", "field_id": "H", "data_type": "I",
                "length_total": "J", "length_dec": "K",
                "notes": "M", "sample_value": "N", "match_score": "AE",
                "match_source": "AF", "verify": "AG",
            },
        },
    },
}


def _deep_merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def load() -> dict:
    if not CONFIG_PATH.exists():
        return {**DEFAULTS, "excel": EXCEL_DEFAULTS}
    with CONFIG_PATH.open(encoding="utf-8") as f:
        data = json.load(f)
    excel_override = data.pop("excel", {})
    merged = {**DEFAULTS, **data}
    merged["excel"] = _deep_merge(EXCEL_DEFAULTS, excel_override)
    return merged


def save(cfg: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
```

- [ ] **Step 4: Run all config tests**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_config.py -v
```

Expected: all 6 tests PASS.

- [ ] **Step 5: Commit**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && git add config.py tests/test_config.py && git commit -m "feat: add EXCEL_DEFAULTS and deep_merge to config"
```

---

## Task 2: excel/reader.py — expand fields, accept excel_cfg, detect direction

**Files:**
- Modify: `excel/reader.py`
- Modify: `tests/test_excel_reader.py`

- [ ] **Step 1: Rewrite tests/test_excel_reader.py**

Replace the entire file:

```python
# tests/test_excel_reader.py
import pytest
import openpyxl
from pathlib import Path
from excel.reader import read_fields, read_kb_fields, InterfaceFieldInput, ExcelReadError
import config


def _default_excel_cfg() -> dict:
    return config.load()["excel"]


def _make_if_xlsx(
    tmp_path: Path,
    *,
    module: str = "MM",
    if_name: str = "TEST_IF",
    if_desc: str = "Test interface",
    data_rows: list[dict] | None = None,
    detection_value: str = "",
    excel_cfg: dict | None = None,
) -> Path:
    cfg = excel_cfg or _default_excel_cfg()
    sheet_head = cfg["sheet_head"]
    sheet_data = cfg["sheet_data"]
    header_row = cfg["header_row"]

    # Direction is always "normal" in helpers unless overridden
    direction = "sap" if detection_value and cfg["detection"]["keyword"].upper() in detection_value.upper() else "normal"
    hcols = cfg["directions"][direction]["input_header_cols"]
    rcols = cfg["directions"][direction]["input_row_cols"]
    start_row = cfg["start_row"]
    det = cfg["detection"]

    wb = openpyxl.Workbook()
    wb.active.title = sheet_head
    ws_head = wb[sheet_head]
    ws_head[f"{hcols['module']}{header_row}"] = module
    ws_head[f"{hcols['if_name']}{header_row}"] = if_name
    ws_head[f"{hcols['if_desc']}{header_row}"] = if_desc
    if detection_value:
        ws_head[f"{det['col']}{det['row']}"] = detection_value

    ws_data = wb.create_sheet(sheet_data)
    if data_rows is None:
        data_rows = [{"field_name": "MATNR", "field_text": "品目コード", "sample_value": "100", "remark": ""}]
    for i, r in enumerate(data_rows):
        row_num = start_row + i
        for col_key, col_letter in rcols.items():
            if col_key in r:
                ws_data[f"{col_letter}{row_num}"] = r[col_key]

    path = tmp_path / "test_if.xlsx"
    wb.save(path)
    return path


# ── read_fields: normal direction ────────────────────────────────────────────

def test_reads_interface_fields(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, data_rows=[
        {"field_name": "MATNR", "field_text": "品目コード", "sample_value": "100", "remark": "備考"},
        {"field_name": "WERKS", "field_text": "プラント",   "sample_value": "1000", "remark": ""},
    ])
    fields, wb, direction = read_fields(path, excel_cfg)
    assert direction == "normal"
    assert len(fields) == 2
    assert fields[0].fieldName == "MATNR"
    assert fields[0].fieldText == "品目コード"
    assert fields[0].sampleValue == "100"
    assert fields[0].remark == "備考"
    assert fields[1].fieldName == "WERKS"
    wb.close()


def test_reads_header_from_head_sheet(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, module="SD", if_name="IF_SD_001", if_desc="Sales order")
    fields, wb, direction = read_fields(path, excel_cfg)
    assert fields[0].module == "SD"
    assert fields[0].ifName == "IF_SD_001"
    assert fields[0].ifDesc == "Sales order"
    wb.close()


def test_rowindex_matches_sheet_row(tmp_path):
    excel_cfg = _default_excel_cfg()
    start = excel_cfg["start_row"]
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, data_rows=[
        {"field_name": "MATNR"}, {"field_name": "WERKS"},
    ])
    fields, wb, _ = read_fields(path, excel_cfg)
    assert fields[0].rowIndex == start
    assert fields[1].rowIndex == start + 1
    wb.close()


def test_skips_blank_rows(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, data_rows=[
        {"field_name": "MATNR"}, {"field_name": None}, {"field_name": ""}, {"field_name": "WERKS"},
    ])
    fields, wb, _ = read_fields(path, excel_cfg)
    assert len(fields) == 2
    assert fields[1].fieldName == "WERKS"
    wb.close()


def test_raises_on_missing_head_sheet(tmp_path):
    excel_cfg = _default_excel_cfg()
    wb = openpyxl.Workbook()
    wb.active.title = excel_cfg["sheet_data"]
    path = tmp_path / "no_head.xlsx"
    wb.save(path)
    with pytest.raises(ExcelReadError, match=excel_cfg["sheet_head"]):
        read_fields(path, excel_cfg)


def test_raises_on_missing_data_sheet(tmp_path):
    excel_cfg = _default_excel_cfg()
    wb = openpyxl.Workbook()
    wb.active.title = excel_cfg["sheet_head"]
    path = tmp_path / "no_data.xlsx"
    wb.save(path)
    with pytest.raises(ExcelReadError, match=excel_cfg["sheet_data"]):
        read_fields(path, excel_cfg)


def test_to_dict_includes_all_new_fields(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, data_rows=[{
        "field_name": "MATNR", "field_text": "品目", "sample_value": "X", "remark": "R",
        "table_id": "EKKO", "field_id": "EBELN", "key_flag": "○", "obligatory": "必須",
        "data_type": "CHAR", "length_total": "10", "length_dec": "0",
        "is_append": "Y", "verify": "○",
    }])
    fields, wb, _ = read_fields(path, excel_cfg)
    d = fields[0].to_dict()
    expected_keys = {
        "rowIndex", "module", "ifName", "ifDesc",
        "fieldName", "fieldText", "sampleValue", "remark",
        "tableId", "fieldId", "keyFlag", "obligatory",
        "dataType", "lengthTotal", "lengthDec", "isAppend", "verify",
    }
    assert set(d.keys()) == expected_keys
    assert d["tableId"] == "EKKO"
    assert d["fieldId"] == "EBELN"
    wb.close()


# ── direction detection ───────────────────────────────────────────────────────

def test_detects_sap_direction(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, detection_value="SAP S/4HANA",
                          data_rows=[{"field_name": "MATNR"}])
    fields, wb, direction = read_fields(path, excel_cfg)
    assert direction == "sap"
    wb.close()


def test_detects_normal_direction_when_no_keyword(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, detection_value="External System",
                          data_rows=[{"field_name": "MATNR"}])
    fields, wb, direction = read_fields(path, excel_cfg)
    assert direction == "normal"
    wb.close()


def test_detects_normal_direction_when_detection_cell_empty(tmp_path):
    excel_cfg = _default_excel_cfg()
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, data_rows=[{"field_name": "MATNR"}])
    fields, wb, direction = read_fields(path, excel_cfg)
    assert direction == "normal"
    wb.close()


def test_sap_direction_reads_from_sap_columns(tmp_path):
    excel_cfg = _default_excel_cfg()
    # SAP direction: field_name is in column S (not C)
    path = _make_if_xlsx(tmp_path, excel_cfg=excel_cfg, detection_value="SAP",
                          data_rows=[{"field_name": "MATNR_SAP", "field_text": "SAP品目"}])
    fields, wb, direction = read_fields(path, excel_cfg)
    assert direction == "sap"
    assert fields[0].fieldName == "MATNR_SAP"
    assert fields[0].fieldText == "SAP品目"
    wb.close()


# ── read_kb_fields (unchanged) ────────────────────────────────────────────────

def _make_kb_xlsx(tmp_path: Path, headers: list, rows: list[list]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    path = tmp_path / "kb.xlsx"
    wb.save(path)
    return path


def test_kb_reads_standard_columns(tmp_path):
    path = _make_kb_xlsx(tmp_path,
        ["sourceField", "sourceDesc", "sourceTable", "targetField", "targetTable"],
        [["MATNR", "品目コード", "EKPO", "Material", "C_PurchaseOrderItemTP"]],
    )
    records = read_kb_fields(path)
    assert len(records) == 1
    assert records[0]["sourceField"] == "MATNR"
    assert records[0]["targetField"] == "Material"


def test_kb_skips_blank_rows(tmp_path):
    path = _make_kb_xlsx(tmp_path,
        ["sourceField", "sourceDesc"],
        [["MATNR", "品目"], [None, None], ["WERKS", "プラント"]],
    )
    records = read_kb_fields(path)
    assert len(records) == 2


def test_kb_raises_on_missing_sourcefield(tmp_path):
    path = _make_kb_xlsx(tmp_path, ["sourceDesc"], [["Material"]])
    with pytest.raises(ExcelReadError, match="sourceField"):
        read_kb_fields(path)
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_excel_reader.py -v 2>&1 | tail -20
```

Expected: multiple FAILED (old function signature).

- [ ] **Step 3: Rewrite excel/reader.py**

Replace the entire file:

```python
# excel/reader.py
import warnings
import openpyxl
from pathlib import Path
from dataclasses import dataclass

warnings.filterwarnings(
    "ignore",
    message="DrawingML support is incomplete.*",
    category=UserWarning,
)

# KB upload aliases (unchanged)
KB_ALIASES: dict[str, str] = {
    "ifname":      "ifName",
    "sourcedesc":  "sourceDesc",
    "sourcetable": "sourceTable",
    "sourcefield": "sourceField",
    "targetdesc":  "targetDesc",
    "targettable": "targetTable",
    "targetfield": "targetField",
    "notes":       "notes",
}

KB_FIELDS = ["ifName", "sourceDesc", "sourceTable", "sourceField",
             "targetDesc", "targetTable", "targetField", "notes"]


@dataclass
class InterfaceFieldInput:
    rowIndex:    int
    module:      str = ""
    ifName:      str = ""
    ifDesc:      str = ""
    fieldName:   str = ""
    fieldText:   str = ""
    sampleValue: str = ""
    remark:      str = ""
    tableId:     str = ""
    fieldId:     str = ""
    keyFlag:     str = ""
    obligatory:  str = ""
    dataType:    str = ""
    lengthTotal: str = ""
    lengthDec:   str = ""
    isAppend:    str = ""
    verify:      str = ""

    def to_dict(self) -> dict:
        return {
            "rowIndex":    self.rowIndex,
            "module":      self.module,
            "ifName":      self.ifName,
            "ifDesc":      self.ifDesc,
            "fieldName":   self.fieldName,
            "fieldText":   self.fieldText,
            "sampleValue": self.sampleValue,
            "remark":      self.remark,
            "tableId":     self.tableId,
            "fieldId":     self.fieldId,
            "keyFlag":     self.keyFlag,
            "obligatory":  self.obligatory,
            "dataType":    self.dataType,
            "lengthTotal": self.lengthTotal,
            "lengthDec":   self.lengthDec,
            "isAppend":    self.isAppend,
            "verify":      self.verify,
        }


class ExcelReadError(Exception):
    pass


def _cell_str(ws, col: str, row: int) -> str:
    v = ws[f"{col}{row}"].value
    return str(v).strip() if v is not None else ""


def read_fields(
    path: Path, excel_cfg: dict
) -> tuple[list[InterfaceFieldInput], openpyxl.Workbook, str]:
    """Read an IF mapping Excel using excel_cfg column definitions.

    Returns (fields, workbook, direction) where direction is 'normal' or 'sap'.
    The workbook is kept open so write_results() can use it.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        raise ExcelReadError(f"{path.name}: cannot open — {e}") from e

    sheet_head = excel_cfg["sheet_head"]
    sheet_data = excel_cfg["sheet_data"]

    if sheet_head not in wb.sheetnames:
        raise ExcelReadError(
            f"{path.name}: sheet '{sheet_head}' not found (sheets: {wb.sheetnames})"
        )
    if sheet_data not in wb.sheetnames:
        raise ExcelReadError(
            f"{path.name}: sheet '{sheet_data}' not found (sheets: {wb.sheetnames})"
        )

    ws_head = wb[sheet_head]
    ws_data = wb[sheet_data]

    # Detect direction
    det = excel_cfg["detection"]
    det_value = _cell_str(ws_head, det["col"], det["row"])
    keyword = det["keyword"].upper()
    direction = "sap" if keyword in det_value.upper() else "normal"

    dir_cfg = excel_cfg["directions"][direction]
    hcols = dir_cfg["input_header_cols"]
    rcols = dir_cfg["input_row_cols"]
    header_row = excel_cfg["header_row"]
    start_row = excel_cfg["start_row"]

    module  = _cell_str(ws_head, hcols["module"],  header_row)
    if_name = _cell_str(ws_head, hcols["if_name"], header_row)
    if_desc = _cell_str(ws_head, hcols["if_desc"], header_row)

    fields: list[InterfaceFieldInput] = []
    max_row = ws_data.max_row or 1000

    for row in range(start_row, max_row + 1):
        fn_col = rcols["field_name"]
        field_name_raw = ws_data[f"{fn_col}{row}"].value
        if field_name_raw is None or str(field_name_raw).strip() in ("", "e"):
            continue

        def _get(key: str) -> str:
            col = rcols.get(key)
            return _cell_str(ws_data, col, row) if col else ""

        fields.append(InterfaceFieldInput(
            rowIndex=row,
            module=module,
            ifName=if_name,
            ifDesc=if_desc,
            fieldName=str(field_name_raw).strip(),
            fieldText=_get("field_text"),
            sampleValue=_get("sample_value"),
            remark=_get("remark"),
            tableId=_get("table_id"),
            fieldId=_get("field_id"),
            keyFlag=_get("key_flag"),
            obligatory=_get("obligatory"),
            dataType=_get("data_type"),
            lengthTotal=_get("length_total"),
            lengthDec=_get("length_dec"),
            isAppend=_get("is_append"),
            verify=_get("verify"),
        ))

    return fields, wb, direction


def read_kb_fields(path: Path) -> list[dict]:
    """Read a knowledge-base Excel (flat header row) → list of CustomFieldUploadInput dicts."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise ExcelReadError(f"{path.name}: cannot open — {e}") from e

    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows:
        raise ExcelReadError(f"{path.name}: empty file")

    headers = [str(c).strip() if c is not None else "" for c in all_rows[0]]
    col_map: dict[str, int] = {}
    for i, h in enumerate(headers):
        canonical = KB_ALIASES.get(h.lower(), h.lower())
        col_map[canonical] = i

    if "sourceField" not in col_map:
        raise ExcelReadError(
            f"{path.name}: missing required column 'sourceField' (got: {headers})"
        )

    records: list[dict] = []
    for row in all_rows[1:]:
        idx = col_map["sourceField"]
        sf = row[idx] if idx < len(row) else None
        if not sf:
            continue
        record: dict = {}
        for field in KB_FIELDS:
            i = col_map.get(field)
            record[field] = str(row[i]).strip() if i is not None and i < len(row) and row[i] else ""
        records.append(record)

    return records
```

- [ ] **Step 4: Run reader tests**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_excel_reader.py -v
```

Expected: all tests PASS.

- [ ] **Step 5: Commit**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && git add excel/reader.py tests/test_excel_reader.py && git commit -m "feat: expand InterfaceFieldInput and add direction detection to reader"
```

---

## Task 3: excel/writer.py — accept output_cols

**Files:**
- Modify: `excel/writer.py`
- Modify: `tests/test_excel_writer.py`

- [ ] **Step 1: Rewrite tests/test_excel_writer.py**

Replace the entire file:

```python
# tests/test_excel_writer.py
import openpyxl
import pytest
from pathlib import Path
from excel.writer import write_results
import config


def _default_output_cols(direction: str = "normal") -> dict:
    return config.load()["excel"]["directions"][direction]["output_cols"]


def _make_workbook(sheet_head: str = "対象IF", sheet_data: str = "IFマッピング定義",
                   data_rows: list[dict] | None = None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.active.title = sheet_head
    ws = wb.create_sheet(sheet_data)
    start = config.load()["excel"]["start_row"]
    if data_rows:
        for i, row in enumerate(data_rows):
            ws[f"C{start + i}"] = row.get("fieldName", "")
    return wb


def test_output_filename_format(tmp_path):
    output_cols = _default_output_cols()
    wb = _make_workbook(data_rows=[{"fieldName": "MATNR"}])
    input_path = tmp_path / "IF_MM_001.xlsx"
    start = config.load()["excel"]["start_row"]
    results = [{"rowIndex": start, "tableId": "T", "fieldId": "F",
                "dataType": "Char", "fieldText": "Material", "matchScore": 1.0,
                "matchSource": "custom", "notes": ""}]
    out = write_results(input_path, wb, results, output_cols)
    assert out.name.startswith("processed_")
    assert out.name.endswith("IF_MM_001.xlsx")
    assert out.parent == tmp_path


def test_writes_to_correct_columns(tmp_path):
    output_cols = _default_output_cols()
    wb = _make_workbook(data_rows=[{"fieldName": "MATNR"}])
    input_path = tmp_path / "test.xlsx"
    start = config.load()["excel"]["start_row"]
    results = [{
        "rowIndex": start,
        "tableId": "CDS_View", "fieldId": "Material",
        "dataType": "Char10", "fieldText": "品目コード",
        "matchScore": 0.95, "matchSource": "ai", "notes": "High confidence",
    }]
    out = write_results(input_path, wb, results, output_cols)

    sheet_data = config.load()["excel"]["sheet_data"]
    saved_ws = openpyxl.load_workbook(out)[sheet_data]
    assert saved_ws[f"{output_cols['table_id']}{start}"].value == "CDS_View"
    assert saved_ws[f"{output_cols['field_id']}{start}"].value == "Material"
    assert saved_ws[f"{output_cols['data_type']}{start}"].value == "Char10"
    assert saved_ws[f"{output_cols['field_name']}{start}"].value == "品目コード"
    assert saved_ws[f"{output_cols['match_source']}{start}"].value == "ai"
    assert saved_ws[f"{output_cols['notes']}{start}"].value == "High confidence"


def test_result_matched_by_rowindex(tmp_path):
    output_cols = _default_output_cols()
    wb = _make_workbook(data_rows=[{"fieldName": "MATNR"}, {"fieldName": "WERKS"}])
    input_path = tmp_path / "test.xlsx"
    start = config.load()["excel"]["start_row"]
    row1, row2 = start, start + 1
    results = [
        {"rowIndex": row2, "tableId": "T2", "fieldId": "F2",
         "dataType": "", "fieldText": "", "matchScore": 0.8, "matchSource": "ai", "notes": ""},
        {"rowIndex": row1, "tableId": "T1", "fieldId": "F1",
         "dataType": "", "fieldText": "", "matchScore": 1.0, "matchSource": "custom", "notes": ""},
    ]
    out = write_results(input_path, wb, results, output_cols)
    sheet_data = config.load()["excel"]["sheet_data"]
    ws = openpyxl.load_workbook(out)[sheet_data]
    assert ws[f"{output_cols['table_id']}{row1}"].value == "T1"
    assert ws[f"{output_cols['table_id']}{row2}"].value == "T2"


def test_rows_with_no_result_unchanged(tmp_path):
    output_cols = _default_output_cols()
    wb = _make_workbook(data_rows=[{"fieldName": "MATNR"}, {"fieldName": "WERKS"}])
    input_path = tmp_path / "test.xlsx"
    start = config.load()["excel"]["start_row"]
    row1, row2 = start, start + 1
    results = [{"rowIndex": row1, "tableId": "T1", "fieldId": "F1",
                "dataType": "", "fieldText": "", "matchScore": 1.0, "matchSource": "custom", "notes": ""}]
    out = write_results(input_path, wb, results, output_cols)
    sheet_data = config.load()["excel"]["sheet_data"]
    ws = openpyxl.load_workbook(out)[sheet_data]
    assert ws[f"{output_cols['table_id']}{row2}"].value is None


def test_sap_direction_writes_to_sap_output_cols(tmp_path):
    output_cols = _default_output_cols("sap")
    wb = _make_workbook(data_rows=[{"fieldName": "MATNR"}])
    input_path = tmp_path / "test_sap.xlsx"
    start = config.load()["excel"]["start_row"]
    results = [{"rowIndex": start, "tableId": "EKKO", "fieldId": "EBELN",
                "dataType": "CHAR10", "fieldText": "PO Number",
                "matchScore": 0.9, "matchSource": "ai", "notes": ""}]
    out = write_results(input_path, wb, results, output_cols)
    sheet_data = config.load()["excel"]["sheet_data"]
    ws = openpyxl.load_workbook(out)[sheet_data]
    # SAP direction: table_id → G, field_id → H
    assert ws[f"G{start}"].value == "EKKO"
    assert ws[f"H{start}"].value == "EBELN"
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_excel_writer.py -v 2>&1 | tail -20
```

Expected: FAILED (old function signature / missing `output_cols`).

- [ ] **Step 3: Rewrite excel/writer.py**

Replace the entire file:

```python
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
```

- [ ] **Step 4: Run writer tests**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_excel_writer.py -v
```

Expected: all 5 tests PASS.

- [ ] **Step 5: Commit**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && git add excel/writer.py tests/test_excel_writer.py && git commit -m "feat: accept output_cols in write_results, remove hardcoded OUTPUT_COLS"
```

---

## Task 4: match_frame.py — pass excel_cfg, store direction

**Files:**
- Modify: `gui/frames/match_frame.py`
- Modify: `tests/test_match_worker.py`

- [ ] **Step 1: Rewrite tests/test_match_worker.py**

Replace the entire file:

```python
# tests/test_match_worker.py
import queue
import threading
from pathlib import Path
from unittest.mock import patch
import openpyxl
import pytest

from gui.frames.match_frame import match_worker
import config


def _default_excel_cfg() -> dict:
    return config.load()["excel"]


def _make_if_xlsx(tmp_path: Path, name: str = "input.xlsx",
                  detection_value: str = "") -> Path:
    cfg = _default_excel_cfg()
    sheet_head = cfg["sheet_head"]
    sheet_data = cfg["sheet_data"]
    header_row = cfg["header_row"]
    start_row = cfg["start_row"]
    hcols = cfg["directions"]["normal"]["input_header_cols"]
    rcols = cfg["directions"]["normal"]["input_row_cols"]
    det = cfg["detection"]

    wb = openpyxl.Workbook()
    wb.active.title = sheet_head
    ws_head = wb[sheet_head]
    ws_head[f"{hcols['if_name']}{header_row}"] = "IF_TEST"
    ws_head[f"{hcols['module']}{header_row}"] = "MM"
    ws_head[f"{hcols['if_desc']}{header_row}"] = "Test"
    if detection_value:
        ws_head[f"{det['col']}{det['row']}"] = detection_value
    ws_data = wb.create_sheet(sheet_data)
    ws_data[f"{rcols['field_name']}{start_row}"] = "MATNR"
    ws_data[f"{rcols['field_text']}{start_row}"] = "品目コード"
    path = tmp_path / name
    wb.save(path)
    return path


def _drain(q: queue.Queue) -> list[dict]:
    msgs = []
    while not q.empty():
        msgs.append(q.get_nowait())
    return msgs


def test_worker_emits_done_with_results(tmp_path):
    xlsx = _make_if_xlsx(tmp_path)
    excel_cfg = _default_excel_cfg()
    q = queue.Queue()
    stop = threading.Event()
    start = excel_cfg["start_row"]
    mock_results = [{"rowIndex": start, "tableId": "C_PO", "fieldId": "Material",
                     "matchSource": "custom", "matchScore": 1.0, "notes": ""}]
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        MockClient.return_value.match.return_value = mock_results
        match_worker([str(xlsx)], "claude", "ja", "http://localhost:4004", excel_cfg, q, stop)
    msgs = _drain(q)
    assert "done" in [m["type"] for m in msgs]
    done_msg = next(m for m in msgs if m["type"] == "done")
    assert done_msg["results"] == mock_results


def test_worker_stores_direction_in_done_message(tmp_path):
    xlsx_normal = _make_if_xlsx(tmp_path, name="normal.xlsx")
    xlsx_sap    = _make_if_xlsx(tmp_path, name="sap.xlsx", detection_value="SAP")
    excel_cfg = _default_excel_cfg()
    q = queue.Queue()
    stop = threading.Event()
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        MockClient.return_value.match.return_value = []
        match_worker([str(xlsx_normal), str(xlsx_sap)], "claude", "ja",
                     "http://localhost:4004", excel_cfg, q, stop)
    msgs = _drain(q)
    done = next(m for m in msgs if m["type"] == "done")
    assert done["directions"][str(xlsx_normal)] == "normal"
    assert done["directions"][str(xlsx_sap)] == "sap"


def test_worker_emits_error_when_cap_unreachable(tmp_path):
    xlsx = _make_if_xlsx(tmp_path)
    excel_cfg = _default_excel_cfg()
    q = queue.Queue()
    stop = threading.Event()
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = False
        match_worker([str(xlsx)], "claude", "ja", "http://localhost:4004", excel_cfg, q, stop)
    msgs = _drain(q)
    assert any(m["type"] == "error" for m in msgs)


def test_worker_skips_bad_excel_and_continues(tmp_path):
    bad = tmp_path / "bad.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "WrongSheet"
    wb.save(bad)

    good = _make_if_xlsx(tmp_path, name="good.xlsx")
    excel_cfg = _default_excel_cfg()
    q = queue.Queue()
    stop = threading.Event()
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        MockClient.return_value.match.return_value = []
        match_worker([str(bad), str(good)], "claude", "ja", "http://localhost:4004", excel_cfg, q, stop)
    msgs = _drain(q)
    log_texts = [m["text"] for m in msgs if m["type"] == "log"]
    assert any("ERROR" in t for t in log_texts)
    assert any(m["type"] == "done" for m in msgs)


def test_worker_respects_stop_event(tmp_path):
    files = []
    for i in range(3):
        sub = tmp_path / f"f{i}"
        sub.mkdir()
        files.append(str(_make_if_xlsx(sub, name="input.xlsx")))
    excel_cfg = _default_excel_cfg()
    q = queue.Queue()
    stop = threading.Event()
    stop.set()
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        match_worker(files, "claude", "ja", "http://localhost:4004", excel_cfg, q, stop)
    msgs = _drain(q)
    log_texts = [m["text"] for m in msgs if m["type"] == "log"]
    assert any("停止" in t for t in log_texts)
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_match_worker.py -v 2>&1 | tail -20
```

Expected: FAILED (old `match_worker` signature).

- [ ] **Step 3: Rewrite gui/frames/match_frame.py**

Replace the entire file:

```python
# gui/frames/match_frame.py
import queue
import threading
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

import openpyxl
import customtkinter as ctk

import config
from api.cap_client import CapClient
from excel.reader import read_fields, ExcelReadError
from excel.writer import write_results
from gui.frames import BaseFrame

LOG_COLORS = {"info": "#22c55e", "warn": "#f59e0b", "error": "#ef4444", "step": "#7b8cde"}


def match_worker(
    files: list[str],
    provider: str,
    language: str,
    server_url: str,
    excel_cfg: dict,
    q: queue.Queue,
    stop: threading.Event,
) -> None:
    """Background worker: runs matching pipeline and puts messages on q."""

    def log(text, level="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        q.put({"type": "log", "text": f"[{ts}]  {level.upper():<5} {text}", "level": level})

    client = CapClient(server_url)
    if not client.ping():
        log(f"无法连接到 {server_url} — 请检查 CAP 是否已启动", "error")
        q.put({"type": "error", "msg": "CAP service unreachable"})
        return

    all_results: list[dict] = []
    all_wb: dict[str, openpyxl.Workbook] = {}
    all_row_indices: dict[str, set[int]] = {}
    all_directions: dict[str, str] = {}

    for i, file_str in enumerate(files):
        if stop.is_set():
            log("用户停止 — 处理中断", "warn")
            break

        path = Path(file_str)
        try:
            fields, workbook, direction = read_fields(path, excel_cfg)
            all_wb[file_str] = workbook
            all_row_indices[file_str] = {f.rowIndex for f in fields}
            all_directions[file_str] = direction
            log(f"解析 {path.name} — {len(fields)} 条字段 [{direction}]")
        except ExcelReadError as e:
            log(str(e), "error")
            continue

        try:
            results = client.match(
                [f.to_dict() for f in fields],
                provider=provider,
                language=language,
            )
            all_results.extend(results)
            custom = sum(1 for r in results if r.get("matchSource") == "custom")
            ai     = sum(1 for r in results if r.get("matchSource") == "ai")
            log(f"匹配完成: {len(results)} 条 (知识库 {custom} · AI {ai})", "step")
        except Exception as e:
            log(f"{path.name} 匹配失败: {e}", "error")

        q.put({"type": "progress", "pct": (i + 1) / len(files)})

    q.put({
        "type": "done",
        "results": all_results,
        "wb": all_wb,
        "row_indices": all_row_indices,
        "directions": all_directions,
        "files": files,
    })


class MatchFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._files: list[str] = []
        self._results: list[dict] = []
        self._wb: dict[str, openpyxl.Workbook] = {}
        self._row_indices: dict[str, set[int]] = {}
        self._directions: dict[str, str] = {}
        self._queue: queue.Queue = queue.Queue()
        self._stop_event = threading.Event()
        self._build()

    def _build(self):
        pad = {"padx": 16, "pady": 6}

        self._drop_zone = ctk.CTkButton(
            self, text="📂  点击选择 Excel 文件（可多选）\n支持 .xlsx / .xls",
            height=72, fg_color="#1e293b", hover_color="#334155",
            text_color="#64748b", command=self._pick_files,
        )
        self._drop_zone.pack(fill="x", **pad)

        self._file_list_frame = ctk.CTkScrollableFrame(self, height=80, fg_color="#1e293b")
        self._file_list_frame.pack(fill="x", padx=16, pady=(0, 6))

        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", padx=16, pady=(0, 6))
        ctk.CTkLabel(opts, text="Provider", font=("", 10), text_color="gray").grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=120).grid(row=1, column=0, padx=(0, 10))
        ctk.CTkLabel(opts, text="语言", font=("", 10), text_color="gray").grid(row=0, column=1, sticky="w")
        self._lang_var = ctk.StringVar(value=self.app.cfg.get("language", "ja"))
        ctk.CTkOptionMenu(opts, variable=self._lang_var,
                          values=["ja", "en", "zh"], width=100).grid(row=1, column=1, padx=(0, 10))
        self._start_btn = ctk.CTkButton(opts, text="▶ 开始匹配", width=100, command=self._start)
        self._start_btn.grid(row=1, column=2, padx=(10, 6))
        self._stop_btn = ctk.CTkButton(opts, text="■ 停止", width=80,
                                       fg_color="#1e293b", command=self._stop, state="disabled")
        self._stop_btn.grid(row=1, column=3)

        self._progress = ctk.CTkProgressBar(self)
        self._progress.set(0)
        self._progress.pack(fill="x", padx=16, pady=(0, 6))

        self._log = ctk.CTkTextbox(self, height=160, font=("Consolas", 10), state="disabled")
        self._log.pack(fill="both", expand=True, padx=16, pady=(0, 6))

        result_bar = ctk.CTkFrame(self, fg_color="transparent")
        result_bar.pack(fill="x", padx=16, pady=(0, 10))
        self._result_label = ctk.CTkLabel(result_bar, text="", font=("", 11), text_color="gray")
        self._result_label.pack(side="left")
        self._export_btn = ctk.CTkButton(result_bar, text="📥 导出结果 Excel",
                                          width=140, command=self._export, state="disabled")
        self._export_btn.pack(side="right")

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=self.app.cfg.get("last_input_dir") or None,
        )
        for p in paths:
            if p not in self._files:
                self._files.append(p)
        if paths:
            self.app.cfg["last_input_dir"] = str(Path(paths[0]).parent)
            config.save(self.app.cfg)
        self._refresh_file_list()

    def _refresh_file_list(self):
        for w in self._file_list_frame.winfo_children():
            w.destroy()
        for path_str in self._files:
            row = ctk.CTkFrame(self._file_list_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)
            ctk.CTkLabel(row, text=f"📄 {Path(path_str).name}", font=("", 11)).pack(side="left")
            ctk.CTkButton(row, text="✕", width=24, height=20,
                          command=lambda p=path_str: self._remove_file(p)).pack(side="right")

    def _remove_file(self, path_str: str):
        self._files.remove(path_str)
        self._refresh_file_list()

    def _log_append(self, text: str, level: str = "info"):
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _start(self):
        if not self._files:
            self._log_append("[ERROR] 请先选择 Excel 文件", "error")
            return
        self._results.clear()
        self._wb.clear()
        self._row_indices.clear()
        self._directions.clear()
        self._stop_event.clear()
        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._export_btn.configure(state="disabled")
        self._progress.set(0)
        self._progress.configure(progress_color=("#7b8cde", "#7b8cde"))

        excel_cfg = self.app.cfg.get("excel", config.EXCEL_DEFAULTS)
        t = threading.Thread(
            target=match_worker,
            args=(list(self._files), self._provider_var.get(), self._lang_var.get(),
                  self.app.cfg.get("server_url", "http://localhost:4004"),
                  excel_cfg, self._queue, self._stop_event),
            daemon=True,
        )
        t.start()
        self.after(100, self._poll)

    def _stop(self):
        self._stop_event.set()
        self._progress.configure(progress_color=("#f59e0b", "#f59e0b"))

    def _poll(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                if msg["type"] == "log":
                    self._log_append(msg["text"], msg.get("level", "info"))
                elif msg["type"] == "progress":
                    self._progress.set(msg["pct"])
                elif msg["type"] == "done":
                    self._results = msg["results"]
                    self._wb = msg["wb"]
                    self._row_indices = msg["row_indices"]
                    self._directions = msg["directions"]
                    self._on_done()
                    return
                elif msg["type"] == "error":
                    self._on_error()
                    return
        except Exception:
            pass
        self.after(100, self._poll)

    def _on_done(self):
        custom = sum(1 for r in self._results if r.get("matchSource") == "custom")
        ai     = sum(1 for r in self._results if r.get("matchSource") == "ai")
        self._result_label.configure(
            text=f"完成: {len(self._results)} 条 ｜ 知识库 {custom} ｜ AI {ai}"
        )
        self._progress.set(1.0)
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
        if self._results:
            self._export_btn.configure(state="normal")

    def _on_error(self):
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")

    def _export(self):
        excel_cfg = self.app.cfg.get("excel", config.EXCEL_DEFAULTS)
        for file_str in self._files:
            workbook = self._wb.get(file_str)
            if workbook is None:
                continue
            row_idxs = self._row_indices.get(file_str, set())
            file_results = [r for r in self._results if r.get("rowIndex") in row_idxs]
            direction = self._directions.get(file_str, "normal")
            output_cols = excel_cfg["directions"][direction]["output_cols"]
            sheet_data = excel_cfg.get("sheet_data", "IFマッピング定義")
            out = write_results(Path(file_str), workbook, file_results, output_cols, sheet_data)
            self._log_append(f"已导出: {out.name}", "info")
```

- [ ] **Step 4: Run match_worker tests**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest tests/test_match_worker.py -v
```

Expected: all 5 tests PASS.

- [ ] **Step 5: Run full test suite to confirm no regressions**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest -v
```

Expected: all tests PASS (test_app_smoke.py and test_cap_client.py should still pass).

- [ ] **Step 6: Commit**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && git add gui/frames/match_frame.py tests/test_match_worker.py && git commit -m "feat: pass excel_cfg to match_worker, store direction per file"
```

---

## Task 5: settings_frame.py — Excel config section

**Files:**
- Modify: `gui/frames/settings_frame.py`

No automated tests for UI. Manual verification described at end.

- [ ] **Step 1: Replace settings_frame.py**

Replace the entire file:

```python
# gui/frames/settings_frame.py
import threading
import customtkinter as ctk
import config
from config import EXCEL_DEFAULTS
from gui.frames import BaseFrame

# Labels for each field key shown in the UI
_FIELD_LABELS: dict[str, str] = {
    # header cols
    "module": "module", "if_name": "if_name", "if_desc": "if_desc",
    # row cols
    "field_name": "field_name", "is_append": "is_append", "key_flag": "key_flag",
    "obligatory": "obligatory", "data_type": "data_type", "table_id": "table_id",
    "field_id": "field_id", "length_total": "length_total", "length_dec": "length_dec",
    "field_text": "field_text", "sample_value": "sample_value", "remark": "remark",
    "verify": "verify",
    # output cols
    "notes": "notes", "match_score": "match_score", "match_source": "match_source",
}


def _build_col_grid(parent: ctk.CTkFrame, label: str, fields: dict[str, str],
                    entries: dict[str, ctk.CTkEntry]) -> None:
    """Build a labelled group of field→column letter entries in a 3-up grid."""
    ctk.CTkLabel(parent, text=label, font=("", 11, "bold"), text_color="#94a3b8").pack(
        anchor="w", padx=4, pady=(8, 2)
    )
    grid = ctk.CTkFrame(parent, fg_color="transparent")
    grid.pack(fill="x", padx=4)

    items = list(fields.items())
    for idx, (key, default_val) in enumerate(items):
        col = idx % 3
        row = idx // 3
        cell = ctk.CTkFrame(grid, fg_color="transparent")
        cell.grid(row=row, column=col, padx=(0, 12), pady=2, sticky="w")
        ctk.CTkLabel(cell, text=_FIELD_LABELS.get(key, key), font=("", 10),
                     text_color="#64748b", width=90, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(cell, width=52)
        entry.insert(0, default_val)
        entry.pack(side="left")
        entries[key] = entry


class SettingsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._excel_entries: dict[str, dict[str, ctk.CTkEntry]] = {
            "detection": {},
            "global": {},
            "normal_header": {}, "normal_row": {}, "normal_output": {},
            "sap_header": {},    "sap_row": {},    "sap_output": {},
        }
        self._build()

    def _build(self):
        outer = ctk.CTkScrollableFrame(self)
        outer.pack(fill="both", expand=True)
        pad = {"padx": 20, "pady": 8}

        ctk.CTkLabel(outer, text="设置", font=("", 16, "bold")).pack(anchor="w", padx=20, pady=(20, 4))

        # ── CAP URL ──────────────────────────────────────────────────────────
        url_frame = ctk.CTkFrame(outer, fg_color="transparent")
        url_frame.pack(fill="x", **pad)
        ctk.CTkLabel(url_frame, text="CAP 服务地址", font=("", 11)).pack(anchor="w")
        row = ctk.CTkFrame(url_frame, fg_color="transparent")
        row.pack(fill="x")
        self._url_entry = ctk.CTkEntry(row, width=280)
        self._url_entry.insert(0, self.app.cfg.get("server_url", "http://localhost:4004"))
        self._url_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(row, text="🔌 测试连接", width=110, command=self._test_conn).pack(side="left", padx=(0, 8))
        self._conn_status = ctk.CTkLabel(row, text="", font=("", 11))
        self._conn_status.pack(side="left")

        # ── Provider + Language ───────────────────────────────────────────────
        opts_frame = ctk.CTkFrame(outer, fg_color="transparent")
        opts_frame.pack(fill="x", **pad)
        ctk.CTkLabel(opts_frame, text="默认 Provider", font=("", 11)).grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts_frame, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=140).grid(row=1, column=0, padx=(0, 20))
        ctk.CTkLabel(opts_frame, text="默认语言", font=("", 11)).grid(row=0, column=1, sticky="w")
        self._lang_var = ctk.StringVar(value=self.app.cfg.get("language", "ja"))
        ctk.CTkOptionMenu(opts_frame, variable=self._lang_var,
                          values=["ja", "en", "zh"], width=140).grid(row=1, column=1)

        # ── Excel 列配置 ──────────────────────────────────────────────────────
        ctk.CTkLabel(outer, text="Excel 列配置", font=("", 13, "bold")).pack(
            anchor="w", padx=20, pady=(16, 4)
        )
        excel_section = ctk.CTkFrame(outer, fg_color="#1e293b", corner_radius=8)
        excel_section.pack(fill="x", padx=20, pady=(0, 8))

        excel_cfg = self.app.cfg.get("excel", EXCEL_DEFAULTS)

        # Global sheet/row settings
        global_frame = ctk.CTkFrame(excel_section, fg_color="transparent")
        global_frame.pack(fill="x", padx=12, pady=(10, 4))

        def _labeled_entry(parent, label, value, width=160) -> ctk.CTkEntry:
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(side="left", padx=(0, 16))
            ctk.CTkLabel(f, text=label, font=("", 10), text_color="#64748b").pack(anchor="w")
            e = ctk.CTkEntry(f, width=width)
            e.insert(0, str(value))
            e.pack()
            return e

        self._sheet_head_entry  = _labeled_entry(global_frame, "Sheet（抬头）", excel_cfg["sheet_head"])
        self._sheet_data_entry  = _labeled_entry(global_frame, "Sheet（数据）", excel_cfg["sheet_data"])
        self._header_row_entry  = _labeled_entry(global_frame, "抬头行", excel_cfg["header_row"], width=60)
        self._start_row_entry   = _labeled_entry(global_frame, "起始行", excel_cfg["start_row"],  width=60)

        det_frame = ctk.CTkFrame(excel_section, fg_color="transparent")
        det_frame.pack(fill="x", padx=12, pady=(4, 8))
        det = excel_cfg["detection"]
        self._det_col_entry     = _labeled_entry(det_frame, "检测列", det["col"],     width=52)
        self._det_row_entry     = _labeled_entry(det_frame, "检测行", det["row"],     width=60)
        self._det_keyword_entry = _labeled_entry(det_frame, "关键字", det["keyword"], width=80)

        # Direction tabs
        tab = ctk.CTkTabview(excel_section)
        tab.pack(fill="x", padx=12, pady=(0, 12))
        tab.add("普通方向")
        tab.add("SAP方向")

        for direction, tab_name, hkey, rkey, okey in [
            ("normal", "普通方向", "normal_header", "normal_row", "normal_output"),
            ("sap",    "SAP方向",  "sap_header",    "sap_row",    "sap_output"),
        ]:
            dir_cfg = excel_cfg["directions"][direction]
            scroll = ctk.CTkScrollableFrame(tab.tab(tab_name), height=340)
            scroll.pack(fill="both", expand=True)
            _build_col_grid(scroll, "抬头列",    dir_cfg["input_header_cols"], self._excel_entries[hkey])
            _build_col_grid(scroll, "输入数据列", dir_cfg["input_row_cols"],    self._excel_entries[rkey])
            _build_col_grid(scroll, "输出列",    dir_cfg["output_cols"],       self._excel_entries[okey])

        # Save button
        ctk.CTkButton(outer, text="💾 保存设置", width=120, command=self._save).pack(
            anchor="e", padx=20, pady=12
        )

    def _test_conn(self):
        url = self._url_entry.get().strip()
        self._conn_status.configure(text="测试中…", text_color="gray")

        def _check():
            from api.cap_client import CapClient
            ok = CapClient(url).ping()
            self.after(0, lambda: self._conn_status.configure(
                text="✓ 已连接" if ok else "✗ 无法连接",
                text_color="#22c55e" if ok else "#ef4444",
            ))

        threading.Thread(target=_check, daemon=True).start()

    def _save(self):
        self.app.cfg["server_url"] = self._url_entry.get().strip()
        self.app.cfg["provider"]   = self._provider_var.get()
        self.app.cfg["language"]   = self._lang_var.get()

        excel_cfg = self.app.cfg.setdefault("excel", {})
        excel_cfg["sheet_head"] = self._sheet_head_entry.get().strip()
        excel_cfg["sheet_data"] = self._sheet_data_entry.get().strip()
        try:
            excel_cfg["header_row"] = int(self._header_row_entry.get())
            excel_cfg["start_row"]  = int(self._start_row_entry.get())
        except ValueError:
            pass
        excel_cfg.setdefault("detection", {})
        excel_cfg["detection"]["col"]     = self._det_col_entry.get().strip().upper()
        excel_cfg["detection"]["keyword"] = self._det_keyword_entry.get().strip()
        try:
            excel_cfg["detection"]["row"] = int(self._det_row_entry.get())
        except ValueError:
            pass

        for direction, hkey, rkey, okey in [
            ("normal", "normal_header", "normal_row", "normal_output"),
            ("sap",    "sap_header",    "sap_row",    "sap_output"),
        ]:
            dir_cfg = excel_cfg.setdefault("directions", {}).setdefault(direction, {})
            dir_cfg["input_header_cols"] = {
                k: e.get().strip().upper() for k, e in self._excel_entries[hkey].items()
            }
            dir_cfg["input_row_cols"] = {
                k: e.get().strip().upper() for k, e in self._excel_entries[rkey].items()
            }
            dir_cfg["output_cols"] = {
                k: e.get().strip().upper() for k, e in self._excel_entries[okey].items()
            }

        config.save(self.app.cfg)
        self.app.update_status(False)
        self.app._refresh_status()
```

- [ ] **Step 2: Run full test suite**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest -v
```

Expected: all tests PASS (settings_frame has no automated tests, but app_smoke should pass).

- [ ] **Step 3: Commit**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && git add gui/frames/settings_frame.py && git commit -m "feat: add Excel column config section to settings frame"
```

---

## Task 6: Update config.example.json

**Files:**
- Modify: `config.example.json`

- [ ] **Step 1: Update config.example.json to include excel section**

Replace `config.example.json`:

```json
{
  "server_url": "http://localhost:4004",
  "provider": "claude",
  "language": "ja",
  "last_input_dir": "",
  "excel": {
    "sheet_head": "対象IF",
    "sheet_data": "IFマッピング定義",
    "header_row": 6,
    "start_row": 5,
    "detection": {
      "col": "F",
      "row": 6,
      "keyword": "SAP"
    },
    "directions": {
      "normal": {
        "input_header_cols": { "module": "D", "if_name": "C", "if_desc": "E" },
        "input_row_cols": {
          "field_name": "C", "is_append": "D", "key_flag": "E", "obligatory": "F",
          "data_type": "I", "table_id": "G", "field_id": "H",
          "length_total": "J", "length_dec": "K",
          "field_text": "L", "sample_value": "N", "remark": "M", "verify": "AF"
        },
        "output_cols": {
          "field_name": "S", "is_append": "T", "key_flag": "U", "obligatory": "V",
          "table_id": "W", "field_id": "X", "data_type": "Y",
          "length_total": "Z", "length_dec": "AA",
          "notes": "AB", "sample_value": "AC", "match_score": "AE",
          "match_source": "AF", "verify": "AG"
        }
      },
      "sap": {
        "input_header_cols": { "module": "D", "if_name": "C", "if_desc": "E" },
        "input_row_cols": {
          "field_name": "S", "is_append": "T", "key_flag": "U", "obligatory": "V",
          "table_id": "W", "field_id": "X", "data_type": "Y",
          "length_total": "Z", "length_dec": "AA",
          "field_text": "AB", "sample_value": "AC", "remark": "AE", "verify": "AF"
        },
        "output_cols": {
          "field_name": "C", "is_append": "D", "key_flag": "E", "obligatory": "F",
          "table_id": "G", "field_id": "H", "data_type": "I",
          "length_total": "J", "length_dec": "K",
          "notes": "M", "sample_value": "N", "match_score": "AE",
          "match_source": "AF", "verify": "AG"
        }
      }
    }
  }
}
```

- [ ] **Step 2: Run full test suite one final time**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && python -m pytest -v
```

Expected: all tests PASS.

- [ ] **Step 3: Final commit**

```bash
cd "d:/Users/PC/Projects/if_mapping_gui" && git add config.example.json && git commit -m "docs: update config.example.json with full excel column config"
```
