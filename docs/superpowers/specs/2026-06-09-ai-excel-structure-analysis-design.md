# Design: AI-Driven Excel Structure Analysis

**Date:** 2026-06-09  
**Scope:** if_mapping_gui (GUI) + if_mapping_cap (CAP backend)  
**Status:** Approved

---

## 1. Problem Statement

`excel/reader.py::read_fields()` currently depends on `EXCEL_DEFAULTS` in `config.py` — a hard-coded mapping of Excel column letters to field names (e.g., `"field_name": "C"`). This breaks whenever a new Excel template uses a different column layout or sheet naming convention.

**Goal:** Replace the fixed `excel_cfg` with an AI-analyzed `ExcelStructure` that is inferred dynamically from the file's actual content before each batch run. No code changes required when a new template variant appears.

---

## 2. Out of Scope

- `read_kb_fields()` (knowledge-base upload) — unchanged in this iteration
- Caching of AI analysis results
- User confirmation/edit of detected structure
- Changes to `match` action or any existing write-back logic

---

## 3. Architecture

```
GUI (match_frame.py) — per-file loop
  │
  ├─Phase 1─► cap_client.analyze_excel_structure(sheetPreviews, provider, language)
  │                ↓  [new CAP action]
  │             excel-structure-analyzer.ts
  │                └─► aiCore.callWithTools(prompt, toolSchema, provider, model)
  │                        ↓ returns ExcelStructure via tool call
  │
  └─Phase 2─► read_fields(path, excel_cfg=ExcelStructure)   ← unchanged function body
                   ↓
              cap_client.match(fields, ...)                  ← unchanged
```

`EXCEL_DEFAULTS` is **not removed** — it remains as the fallback value shown in the Settings UI and for any non-AI code paths.

---

## 4. Interface Contract

### 4.1 New CDS Types (`db/schema.cds`)

```cds
type SheetPreview {
  sheetName   : String;
  previewText : String;   // formatted as "[A]val [B]val ..." per row, rows separated by \n
}

type HeaderCols {
  module  : String;
  if_name : String;
  if_desc : String;
}

type RowCols {
  field_name    : String;
  key_flag      : String;
  obligatory    : String;
  is_append     : String;
  data_type     : String;
  table_id      : String;
  field_id      : String;
  length_total  : String;
  length_dec    : String;
  field_text    : String;
  sample_value  : String;
  remark        : String;
  verify        : String;
}

type ExcelStructure {
  sheet_head  : String;
  sheet_data  : String;
  header_row  : Integer;
  start_row   : Integer;
  direction   : String;   // 'normal' | 'sap' — tells GUI which directions key to populate
  header_cols : HeaderCols;
  row_cols    : RowCols;
}
```

### 4.2 New CDS Action (`srv/if-mapping-service.cds`)

```cds
action analyzeExcelStructure(
  sheetPreviews : array of SheetPreview,
  provider      : String(20),
  language      : String(5)
) returns ExcelStructure;
```

### 4.3 New GUI Client Method (`api/cap_client.py`)

```python
def analyze_excel_structure(
    self,
    sheet_previews: list[dict],   # [{"sheetName": str, "previewText": str}, ...]
    provider: str,
    language: str,
) -> dict:                        # ExcelStructure dict, shape matches EXCEL_DEFAULTS
```

Raises `CapConnectionError` on HTTP error (same pattern as `match()`).

---

## 5. GUI Changes (`if_mapping_gui`)

### 5.1 `excel/reader.py` — new function

```python
PREVIEW_ROWS = 15

def build_sheet_previews(path: Path) -> list[dict]:
    """Return [{sheetName, previewText}] for all sheets in the workbook.

    previewText format: one row per line, cells as "[ColLetter]value".
    Only the first PREVIEW_ROWS rows are included.
    """
```

- Uses `openpyxl.load_workbook(path, data_only=True, read_only=True)` to avoid full load
- Skips sheets where all preview rows are empty
- Column letters derived from `openpyxl.utils.get_column_letter()`

### 5.2 `gui/frames/match_frame.py` — `match_worker()` changes

Before the existing `read_fields()` call, insert:

```python
# Phase 1: AI structure analysis
sheet_previews = build_sheet_previews(path)
excel_cfg = self.client.analyze_excel_structure(
    sheet_previews, cfg["provider"], cfg["language"]
)
# Phase 2: field extraction (unchanged)
fields, wb, direction = read_fields(path, excel_cfg)
```

- The returned `ExcelStructure` is assembled into an `excel_cfg` dict matching `EXCEL_DEFAULTS` shape before passing to `read_fields()`:
  ```python
  direction = excel_struct["direction"]  # 'normal' or 'sap'
  excel_cfg = {
      "sheet_head": excel_struct["sheet_head"],
      "sheet_data": excel_struct["sheet_data"],
      "header_row": excel_struct["header_row"],
      "start_row":  excel_struct["start_row"],
      "detection":  cfg["excel"]["detection"],   # reuse from config; AI doesn't need to re-detect
      "directions": {
          direction: {
              "input_header_cols": excel_struct["header_cols"],
              "input_row_cols":    excel_struct["row_cols"],
          }
      },
  }
  ```
  This means `read_fields()` is **zero-change** — it receives the exact same dict shape it always expected.
- If `analyze_excel_structure` raises `CapConnectionError`, the file is marked as failed and processing continues to the next file (same isolation pattern as other errors in the loop)
- Progress log updated to show "Analyzing structure..." step before "Reading fields..."

### 5.3 No changes to

- `config.py` / `EXCEL_DEFAULTS`
- `excel/writer.py`
- `api/cap_client.py` existing methods
- All other frames

---

## 6. CAP Backend Changes (`if_mapping_cap`)

### 6.1 `db/schema.cds`

Add `SheetPreview`, `HeaderCols`, `RowCols`, `ExcelStructure` type definitions (see §4.1).

### 6.2 `srv/if-mapping-service.cds`

Add `analyzeExcelStructure` action declaration (see §4.2).

### 6.3 `srv/if-mapping-service.ts`

Register handler in `init()`:

```ts
this.on('analyzeExcelStructure', async (req) => {
  const { sheetPreviews, provider, language } = req.data;
  return analyzeExcelStructure(sheetPreviews, provider, language, this.aiCore);
});
```

### 6.4 `srv/analysis/excel-structure-analyzer.ts` (new file)

Responsibilities:
- Build a system prompt that describes the IF mapping document structure
- Define the tool schema as a JSON Schema matching `ExcelStructure`
- Call `aiCore.callWithTools(messages, [toolSchema], provider, model)`
- Cast `result.toolInput` to `ExcelStructure` and return it

**Prompt guidance to AI:**
- The document has two sheets: one header sheet (contains IF metadata: module, IF name, IF description) and one data sheet (contains field rows)
- SAP direction is detected by checking whether a specific cell contains the keyword "SAP"
- Column letters must be valid Excel column identifiers (A–ZZ)
- `start_row` is the first row containing actual field data (not the header row)

**Tool schema** (JSON Schema for `ExcelStructure`) is defined inline in the analyzer file, not loaded from DB, because it is structural not tuneable.

### 6.5 Error handling

- If `callWithTools` returns no `toolInput`, throw a typed error that maps to HTTP 422 via CAP's `req.error(422, ...)`
- Token usage logged via existing `TokenLog` mechanism (pass `correlationId` through if available)

---

## 7. Data Flow — End to End

```
1. User clicks "Run Match" in match_frame
2. For each .xlsx file in input dir:
   a. build_sheet_previews(path)          → list[SheetPreview]
   b. cap_client.analyze_excel_structure() → ExcelStructure dict
   c. read_fields(path, excel_cfg)         → list[InterfaceFieldInput]
   d. cap_client.match(fields, ...)        → list[MatchedFieldResult]
   e. writer.write_results(wb, results)
3. Summary log written
```

---

## 8. Error Handling & Resilience

| Failure point | Behavior |
|---|---|
| `analyze_excel_structure` HTTP error | File skipped, error logged, next file continues |
| AI returns no tool call | CAP returns 422; GUI treats as file failure |
| AI returns wrong column letter | `read_fields` raises `ExcelReadError`; file skipped |
| Sheet name mismatch | `read_fields` raises `ExcelReadError`; file skipped |

Consistent with existing single-file failure isolation pattern in `match_worker()`.

---

## 9. Testing

- **Unit test** (`tests/test_reader.py`): test `build_sheet_previews()` against a sample xlsx — verify format, row limit, empty sheet skipping
- **Unit test** (`tests/test_excel_structure_analyzer.ts`): mock `aiCore.callWithTools`, verify prompt contains sheet preview text, verify returned structure matches expected shape
- **Integration test**: use an actual xlsx fixture with a non-standard column layout; verify full Phase 1 → Phase 2 → match pipeline produces correct field list
