# Excel Config & Direction Design

**Date:** 2026-05-21  
**Project:** if_mapping_gui  

## Problem

The Excel reader/writer in `if_mapping_gui` hardcodes sheet names, row numbers, and column letters. It reads only 4 fields (`fieldName`, `fieldText`, `sampleValue`, `remark`) and supports only one column layout. The CAP pipeline prompts expect a richer field set, and real-world Excel files come in two layouts depending on interface direction (normal vs SAP).

## Goals

1. Read the full field set that the CAP prompts expect (aligning with `if_gen_tool`'s `InterfaceField`)
2. Support two column layouts (normal direction and SAP direction) with automatic detection
3. Make all sheet names, row numbers, and column letters configurable via the settings GUI

## Architecture

### Data Flow

```
User selects file(s)
  → read_fields(path, excel_cfg)
      → read detection cell (excel_cfg.detection.col + row, on sheet_head)
      → "SAP" in cell value → direction = "sap", else "normal"
      → use directions[direction].input_header_cols to read module/if_name/if_desc
      → use directions[direction].input_row_cols to read each data row
      → return (fields, workbook, direction)
  → client.match([f.to_dict() for f in fields], ...)
  → write_results(path, workbook, results,
                  excel_cfg.directions[direction].output_cols)
```

### Files Changed

| File | Change |
|------|--------|
| `config.py` | Add `EXCEL_DEFAULTS`; deep-merge excel section in `load()` |
| `excel/reader.py` | Accept `excel_cfg`; detect direction; return `(fields, wb, direction)`; expand `InterfaceFieldInput` |
| `excel/writer.py` | Accept `output_cols` dict; remove hardcoded `OUTPUT_COLS` |
| `gui/frames/settings_frame.py` | Wrap in scrollable frame; add Excel config section with detection settings + CTkTabview for two directions |
| `gui/frames/match_frame.py` | Pass `excel_cfg` to reader; store `direction` per file; pass `output_cols` to writer |

## Config Structure

Added to `config.json` under key `"excel"`:

```json
{
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
```

Default values mirror `if_gen_tool`'s `ConfigurationManager.get_column_mappings()` (normal) and `get_column_mappings_sap()` (sap).

## Component Design

### `config.py`

- `EXCEL_DEFAULTS`: nested dict with the full default structure above
- `_deep_merge(base, override)`: recursive merge so partial user configs are filled with defaults
- `load()`: after JSON load, call `_deep_merge(EXCEL_DEFAULTS, data.get("excel", {}))` and set `data["excel"]`

### `excel/reader.py`

`InterfaceFieldInput` fields (all `str = ""`except `rowIndex: int`):

```
rowIndex, module, ifName, ifDesc,
fieldName, fieldText, sampleValue, remark,      # existing
tableId, fieldId, keyFlag, obligatory,           # new
dataType, lengthTotal, lengthDec, isAppend, verify  # new
```

`to_dict()` serialises all fields with camelCase keys matching the CAP `InterfaceFieldInput` CDS type surface.

`read_fields(path, excel_cfg) -> tuple[list[InterfaceFieldInput], Workbook, str]`:
- Validate `sheet_head` and `sheet_data` exist, raise `ExcelReadError` with sheet names if not
- Read detection cell; compare `excel_cfg["detection"]["keyword"]` case-insensitively
- Use `directions[direction]` mappings for all column reads
- Skip rows where `field_name` cell is None / empty / `"e"`
- `read_kb_fields` (KB upload helper) is unchanged

### `excel/writer.py`

`write_results(input_path, workbook, results, output_cols) -> Path`:
- `result_map`: `{r["rowIndex"]: r}`
- For each `(row_idx, result)` write `output_cols` keys → column letters
- CAP result keys are camelCase; writer maps them to snake_case output_cols keys via a small internal `_RESULT_KEY_MAP` dict
- Save as `processed_{ts}_{filename}`

`_RESULT_KEY_MAP` (CAP camelCase → output_cols snake_case):

```python
{
  "fieldText":   "field_name",
  "tableId":     "table_id",
  "fieldId":     "field_id",
  "dataType":    "data_type",
  "notes":       "notes",
  "sampleValue": "sample_value",
  "matchScore":  "match_score",
  "matchSource": "match_source",
}
```

### `gui/frames/settings_frame.py`

- Outer frame becomes `CTkScrollableFrame`
- Existing CAP / Provider / Language rows unchanged
- New "Excel 列配置" section:
  - Row 1: `sheet_head` entry, `sheet_data` entry, `header_row` entry, `start_row` entry
  - Row 2: Detection `col`, `row`, `keyword` entries
  - `CTkTabview` with tabs "普通方向" and "SAP方向"
  - Each tab: `CTkScrollableFrame` containing three labelled groups:
    - "抬头列": 3 field×column pairs in a 3-up grid
    - "输入数据列": 13 field×column pairs
    - "输出列": 15 field×column pairs
  - Each entry is `CTkEntry(width=52)` — just a column letter
- `_save()` reads all entries back into `self.app.cfg["excel"]` before calling `config.save()`

### `gui/frames/match_frame.py`

`match_worker` gains `excel_cfg: dict` parameter:
- Calls `read_fields(path, excel_cfg)` → unpacks `(fields, workbook, direction)`
- Stores `direction` in `all_directions: dict[str, str]` keyed by file path
- Passes `all_directions` in the `"done"` queue message

`_export()`:
- Retrieves `direction = self._directions.get(file_str, "normal")`
- Gets `output_cols = self.app.cfg["excel"]["directions"][direction]["output_cols"]`
- Calls `write_results(path, workbook, file_results, output_cols)`

## Error Handling

- Missing sheet → `ExcelReadError` with sheet names list (existing pattern, unchanged)
- Detection cell empty / no keyword match → default to `"normal"` (silent fallback, logged at debug)
- Unknown output_cols key in result → skip silently (writer already does this)

## Testing

Existing tests in `tests/test_config.py` — extend to cover deep_merge behaviour.  
No new test files required for this change; manual verification against sample Excel files covers the direction detection path.
