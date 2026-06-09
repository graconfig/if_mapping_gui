# AI Excel Structure Analysis Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the hard-coded `EXCEL_DEFAULTS` column mapping in `if_mapping_gui` with AI-driven structure detection via a new `analyzeExcelStructure` CAP action, so that new Excel template layouts are handled automatically without code changes.

**Architecture:** GUI reads the first 15 rows of every sheet as formatted text and sends them to a new CAP action `analyzeExcelStructure`. The backend calls `aiCore.callWithTools()` to return an `ExcelStructure` JSON. The GUI assembles this into the same dict shape that `read_fields()` already expects — the core parsing function has zero changes.

**Tech Stack:**
- Backend: TypeScript, `@sap/cds`, `aicore-client.ts` (`callWithTools`)
- Frontend: Python 3.11+, `openpyxl`, `customtkinter`, `requests`
- Test tools: `pytest`, `pytest-mock` (GUI); Vitest or Jest (CAP, if present)

---

## File Map

### `if_mapping_cap` (backend — do first)

| Action | File |
|---|---|
| Create | `srv/analysis/excel-structure-analyzer.ts` |
| Modify | `db/schema.cds` — add 4 new types |
| Modify | `srv/if-mapping-service.cds` — add action declaration |
| Modify | `srv/if-mapping-service.ts` — register handler |

### `if_mapping_gui` (GUI — do after backend compiles)

| Action | File |
|---|---|
| Modify | `excel/reader.py` — add `build_sheet_previews()` |
| Modify | `api/cap_client.py` — add `analyze_excel_structure()` |
| Modify | `gui/frames/match_frame.py` — insert Phase 1 in `match_worker()` |
| Create | `tests/test_reader.py` — unit tests for `build_sheet_previews()` |

---

## Task 1: Add CDS Types to `db/schema.cds`

**Files:**
- Modify: `d:/Users/PC/Projects/if_mapping_cap/db/schema.cds`

- [ ] **Step 1: Append new types after the existing `UploadResult` type (line 152)**

Open `d:/Users/PC/Projects/if_mapping_cap/db/schema.cds` and append the following block at the end of the file (after the closing `}` of `UploadResult`):

```cds
type SheetPreview {
  sheetName   : String(100);
  previewText : String(5000);
}

type HeaderCols {
  module  : String(5);
  if_name : String(5);
  if_desc : String(5);
}

type RowCols {
  field_name   : String(5);
  key_flag     : String(5);
  obligatory   : String(5);
  is_append    : String(5);
  data_type    : String(5);
  table_id     : String(5);
  field_id     : String(5);
  length_total : String(5);
  length_dec   : String(5);
  field_text   : String(5);
  sample_value : String(5);
  remark       : String(5);
  verify       : String(5);
}

type ExcelStructure {
  sheet_head   : String(100);
  sheet_data   : String(100);
  header_row   : Integer;
  start_row    : Integer;
  direction    : String(10);
  header_cols  : HeaderCols;
  row_cols     : RowCols;
}
```

- [ ] **Step 2: Verify the CDS file is valid**

```bash
cd d:/Users/PC/Projects/if_mapping_cap
npx cds compile db/schema.cds --to json > /dev/null
```

Expected: no errors printed to stderr.

- [ ] **Step 3: Commit**

```bash
cd d:/Users/PC/Projects/if_mapping_cap
git add db/schema.cds
git commit -m "feat: add SheetPreview, ExcelStructure CDS types"
```

---

## Task 2: Declare `analyzeExcelStructure` Action in Service CDS

**Files:**
- Modify: `d:/Users/PC/Projects/if_mapping_cap/srv/if-mapping-service.cds`

- [ ] **Step 1: Add import and action declaration**

In `srv/if-mapping-service.cds`, update the `using` block to import the new types, then add the action inside the service:

```cds
using {
  external,
  PromptTemplates as db_PromptTemplates,
  TokenLogs       as db_TokenLogs,
  InterfaceFieldInput,
  MatchedFieldResult,
  CustomFieldUploadInput,
  UploadResult,
  SheetPreview,
  ExcelStructure
} from '../db/schema';

service IfMappingService @(path: '/if-mapping') {

  // Core matching action
  action match(
    fields   : array of InterfaceFieldInput,
    provider : String(20),
    language : String(5)
  ) returns array of MatchedFieldResult;

  // Knowledge base upload
  action uploadCustomFields(
    records : array of CustomFieldUploadInput,
    mode    : String(10)
  ) returns UploadResult;

  // Excel structure analysis
  action analyzeExcelStructure(
    sheetPreviews : array of SheetPreview,
    provider      : String(20),
    language      : String(5)
  ) returns ExcelStructure;

  // Read-only lookups for debugging / admin
  @readonly @cds.persistence.skip entity CdsViews   as projection on external.CdsViews;
  @readonly @cds.persistence.skip entity ViewFields as projection on external.ViewFields;

  // Prompt template management (full CRUD)
  entity PromptTemplates as projection on db_PromptTemplates;
  action reloadPrompts() returns { success : Boolean };

  // Token usage read-only
  @readonly entity TokenLogs as projection on db_TokenLogs;
}
```

- [ ] **Step 2: Verify the CDS compiles**

```bash
cd d:/Users/PC/Projects/if_mapping_cap
npx cds compile srv/if-mapping-service.cds --to json > /dev/null
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add srv/if-mapping-service.cds
git commit -m "feat: declare analyzeExcelStructure action in service CDS"
```

---

## Task 3: Implement `excel-structure-analyzer.ts`

**Files:**
- Create: `d:/Users/PC/Projects/if_mapping_cap/srv/analysis/excel-structure-analyzer.ts`

- [ ] **Step 1: Create the directory and file**

```bash
mkdir -p d:/Users/PC/Projects/if_mapping_cap/srv/analysis
```

Create `srv/analysis/excel-structure-analyzer.ts` with the following content:

```typescript
import type { AiCoreClient } from '../ai/aicore-client.js';
import { buildRequestConfig } from '../utils/config.js';
import { log } from '../utils/logger.js';
import { AppError } from '../utils/errors.js';

export interface ExcelStructure {
  sheet_head:   string;
  sheet_data:   string;
  header_row:   number;
  start_row:    number;
  direction:    'normal' | 'sap';
  header_cols:  {
    module:  string;
    if_name: string;
    if_desc: string;
  };
  row_cols: {
    field_name:   string;
    key_flag:     string;
    obligatory:   string;
    is_append:    string;
    data_type:    string;
    table_id:     string;
    field_id:     string;
    length_total: string;
    length_dec:   string;
    field_text:   string;
    sample_value: string;
    remark:       string;
    verify:       string;
  };
}

export interface SheetPreview {
  sheetName:   string;
  previewText: string;
}

const TOOL_SCHEMA = {
  name: 'report_excel_structure',
  description: 'Report the detected structure of an IF mapping definition Excel file.',
  inputSchema: {
    type: 'object',
    properties: {
      sheet_head: {
        type: 'string',
        description: 'Name of the header sheet (contains module, IF name, IF description metadata).',
      },
      sheet_data: {
        type: 'string',
        description: 'Name of the data sheet (contains field rows).',
      },
      header_row: {
        type: 'integer',
        description: 'Row number (1-based) in the header sheet where module/IF name/IF description are located.',
      },
      start_row: {
        type: 'integer',
        description: 'Row number (1-based) in the data sheet where actual field data begins (first data row, not header row).',
      },
      direction: {
        type: 'string',
        enum: ['normal', 'sap'],
        description: '"sap" if the data sheet has SAP fields on the right side (columns S onward); "normal" otherwise.',
      },
      header_cols: {
        type: 'object',
        description: 'Column letters for metadata cells in the header sheet.',
        properties: {
          module:  { type: 'string', description: 'Column letter for module name.' },
          if_name: { type: 'string', description: 'Column letter for IF name.' },
          if_desc: { type: 'string', description: 'Column letter for IF description.' },
        },
        required: ['module', 'if_name', 'if_desc'],
      },
      row_cols: {
        type: 'object',
        description: 'Column letters for each field in a data row.',
        properties: {
          field_name:   { type: 'string', description: 'Column letter for field name (primary key column; rows where this is empty or "e" are skipped).' },
          key_flag:     { type: 'string', description: 'Column letter for key flag.' },
          obligatory:   { type: 'string', description: 'Column letter for obligatory flag.' },
          is_append:    { type: 'string', description: 'Column letter for append flag.' },
          data_type:    { type: 'string', description: 'Column letter for data type.' },
          table_id:     { type: 'string', description: 'Column letter for table ID.' },
          field_id:     { type: 'string', description: 'Column letter for field ID.' },
          length_total: { type: 'string', description: 'Column letter for total length.' },
          length_dec:   { type: 'string', description: 'Column letter for decimal length.' },
          field_text:   { type: 'string', description: 'Column letter for field description text.' },
          sample_value: { type: 'string', description: 'Column letter for sample value.' },
          remark:       { type: 'string', description: 'Column letter for remarks.' },
          verify:       { type: 'string', description: 'Column letter for verify flag.' },
        },
        required: [
          'field_name', 'key_flag', 'obligatory', 'is_append', 'data_type',
          'table_id', 'field_id', 'length_total', 'length_dec',
          'field_text', 'sample_value', 'remark', 'verify',
        ],
      },
    },
    required: [
      'sheet_head', 'sheet_data', 'header_row', 'start_row',
      'direction', 'header_cols', 'row_cols',
    ],
  },
};

function buildPrompt(sheets: SheetPreview[]): string {
  const previewBlock = sheets
    .map(s => `=== Sheet: "${s.sheetName}" ===\n${s.previewText}`)
    .join('\n\n');

  return `You are analyzing an IF mapping definition Excel file used in SAP integration projects.

The file contains two sheets:
1. A **header sheet** (対象IF or similar) — contains document metadata: module name, IF name (interface name), IF description. These are typically in fixed cells, not in a table row format.
2. A **data sheet** (IFマッピング定義 or similar) — contains a table of interface fields, one field per row. There is a header row labeling columns, followed by data rows.

Direction detection:
- If the data sheet has SAP-side fields starting from column S or later (columns like S, T, U, V, W, X, Y, Z, AA...), set direction = "sap".
- Otherwise set direction = "normal".

Your task: identify the sheet names, the row numbers for metadata and data, and the exact column letters for each field.

Cell format in the previews below: [ColLetter]CellValue — for example [C]fieldName means column C contains "fieldName".

${previewBlock}

Use the report_excel_structure tool to return your findings.`;
}

export async function analyzeExcelStructure(
  sheets:   SheetPreview[],
  provider: string,
  language: string,
  aiCore:   AiCoreClient,
): Promise<ExcelStructure> {
  const config = buildRequestConfig(provider, language);
  const prompt = buildPrompt(sheets);

  log.info('analyzeExcelStructure: calling AI', { provider, sheetCount: sheets.length });

  const result = await aiCore.callWithTools(
    [{ role: 'user', content: prompt }],
    [TOOL_SCHEMA],
    config.provider,
    config.llmModel,
  );

  if (!result.toolInput) {
    throw new AppError('AI did not return a tool call for excel structure analysis', 422);
  }

  return result.toolInput as ExcelStructure;
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd d:/Users/PC/Projects/if_mapping_cap
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add srv/analysis/excel-structure-analyzer.ts
git commit -m "feat: implement excel-structure-analyzer AI analysis"
```

---

## Task 4: Register Handler in `if-mapping-service.ts`

**Files:**
- Modify: `d:/Users/PC/Projects/if_mapping_cap/srv/if-mapping-service.ts`

- [ ] **Step 1: Add import at top of file**

After the existing imports (around line 11), add:

```typescript
import { analyzeExcelStructure } from './analysis/excel-structure-analyzer.js';
```

- [ ] **Step 2: Register the handler inside `init()`, after the `reloadPrompts` handler (line 93)**

```typescript
    this.on('analyzeExcelStructure', async (req) => {
      const { sheetPreviews, provider, language } = req.data as {
        sheetPreviews: import('./analysis/excel-structure-analyzer.js').SheetPreview[];
        provider?:     string;
        language?:     string;
      };
      try {
        return await analyzeExcelStructure(
          sheetPreviews,
          provider ?? 'claude',
          language ?? 'ja',
          aiCore,
        );
      } catch (err) {
        log.error('analyzeExcelStructure failed', { error: String(err) });
        if (err instanceof AppError) {
          return req.error(err.statusCode, err.message);
        }
        return req.error(500, 'Excel structure analysis failed');
      }
    });
```

- [ ] **Step 3: Verify TypeScript compiles**

```bash
cd d:/Users/PC/Projects/if_mapping_cap
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 4: Smoke test — start the server and call the new action**

```bash
cd d:/Users/PC/Projects/if_mapping_cap
npm run start &
sleep 5
curl -s -X POST http://localhost:4004/if-mapping/analyzeExcelStructure \
  -H "Content-Type: application/json" \
  -d '{"sheetPreviews":[{"sheetName":"対象IF","previewText":"[A]test [B]data"}],"provider":"claude","language":"ja"}' | head -c 500
```

Expected: JSON response containing `sheet_head`, `sheet_data`, `row_cols` etc. (actual values depend on AI output).

- [ ] **Step 5: Commit**

```bash
git add srv/if-mapping-service.ts
git commit -m "feat: register analyzeExcelStructure handler"
```

---

## Task 5: Add `build_sheet_previews()` to `excel/reader.py`

**Files:**
- Modify: `d:/Users/PC/Projects/if_mapping_gui/excel/reader.py`
- Create: `d:/Users/PC/Projects/if_mapping_gui/tests/__init__.py` (empty)
- Create: `d:/Users/PC/Projects/if_mapping_gui/tests/test_reader.py`

- [ ] **Step 1: Write the failing test first**

Create `tests/__init__.py` (empty file) and `tests/test_reader.py`:

```python
# tests/test_reader.py
import io
from pathlib import Path
import openpyxl
import pytest

from excel.reader import build_sheet_previews, PREVIEW_ROWS


def _make_xlsx(sheets: dict[str, list[list]]) -> Path:
    """Write an in-memory xlsx to a temp file and return its path."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r, row in enumerate(rows, 1):
            for c, val in enumerate(row, 1):
                ws.cell(r, c, val)
    import tempfile, os
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb.save(path)
    return Path(path)


def test_preview_contains_sheet_name():
    path = _make_xlsx({"Sheet1": [["hello", "world"]]})
    previews = build_sheet_previews(path)
    assert len(previews) == 1
    assert previews[0]["sheetName"] == "Sheet1"
    path.unlink()


def test_preview_text_format():
    """Each cell must appear as [ColLetter]value."""
    path = _make_xlsx({"Data": [["alpha", "beta"]]})
    previews = build_sheet_previews(path)
    text = previews[0]["previewText"]
    assert "[A]alpha" in text
    assert "[B]beta" in text
    path.unlink()


def test_preview_row_limit():
    """Only first PREVIEW_ROWS rows should be included."""
    rows = [[f"r{i}"] for i in range(PREVIEW_ROWS + 5)]
    path = _make_xlsx({"Big": rows})
    previews = build_sheet_previews(path)
    line_count = previews[0]["previewText"].count("\n") + 1
    assert line_count <= PREVIEW_ROWS
    path.unlink()


def test_empty_sheet_skipped():
    """Sheets with no non-empty cells should be excluded."""
    path = _make_xlsx({"Empty": [[None, None]], "Real": [["x"]]})
    previews = build_sheet_previews(path)
    names = [p["sheetName"] for p in previews]
    assert "Empty" not in names
    assert "Real" in names
    path.unlink()


def test_multiple_sheets():
    path = _make_xlsx({
        "対象IF":       [["module", "MyMod"]],
        "IFマッピング定義": [["fieldName", "tableId"]],
    })
    previews = build_sheet_previews(path)
    assert len(previews) == 2
    path.unlink()
```

- [ ] **Step 2: Run the tests to confirm they fail**

```bash
cd d:/Users/PC/Projects/if_mapping_gui
python -m pytest tests/test_reader.py -v 2>&1 | head -30
```

Expected: `ImportError` or `AttributeError` because `build_sheet_previews` doesn't exist yet.

- [ ] **Step 3: Implement `build_sheet_previews()` in `excel/reader.py`**

Add the following constant and function to `excel/reader.py`, after the existing imports and before the `KB_FIELDS` line:

```python
from openpyxl.utils import get_column_letter

PREVIEW_ROWS = 15


def build_sheet_previews(path: Path) -> list[dict]:
    """Return [{sheetName, previewText}] for all non-empty sheets.

    previewText: first PREVIEW_ROWS rows, each cell as [ColLetter]value,
    cells separated by spaces, rows separated by newlines.
    Empty sheets (all preview rows are None) are skipped.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        raise ExcelReadError(f"{path.name}: cannot open for preview — {e}") from e

    previews: list[dict] = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        lines: list[str] = []
        row_count = 0
        for row in ws.iter_rows(max_row=PREVIEW_ROWS, values_only=False):
            row_count += 1
            cells = []
            for cell in row:
                if cell.value is not None:
                    col_letter = get_column_letter(cell.column)
                    cells.append(f"[{col_letter}]{cell.value}")
            if cells:
                lines.append(" ".join(cells))
        if lines:
            previews.append({
                "sheetName":   sheet_name,
                "previewText": "\n".join(lines),
            })
    wb.close()
    return previews
```

- [ ] **Step 4: Run the tests to confirm they pass**

```bash
cd d:/Users/PC/Projects/if_mapping_gui
python -m pytest tests/test_reader.py -v
```

Expected: all 5 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add excel/reader.py tests/__init__.py tests/test_reader.py
git commit -m "feat: add build_sheet_previews() with tests"
```

---

## Task 6: Add `analyze_excel_structure()` to `api/cap_client.py`

**Files:**
- Modify: `d:/Users/PC/Projects/if_mapping_gui/api/cap_client.py`

- [ ] **Step 1: Add the method after `upload_custom_fields` (around line 91)**

```python
def analyze_excel_structure(
    self,
    sheet_previews: list[dict],
    provider: str,
    language: str,
) -> dict:
    """Call analyzeExcelStructure action and return ExcelStructure dict."""
    try:
        r = requests.post(
            f"{self.base}/analyzeExcelStructure",
            json={
                "sheetPreviews": sheet_previews,
                "provider":      provider,
                "language":      language,
            },
            headers=self._headers(),
            timeout=self.timeout,
        )
        r.raise_for_status()
        return r.json()
    except requests.HTTPError as e:
        raise CapConnectionError(str(e)) from e
```

- [ ] **Step 2: Verify the method is importable**

```bash
cd d:/Users/PC/Projects/if_mapping_gui
python -c "from api.cap_client import CapClient; c = CapClient('http://localhost:4004'); print('ok')"
```

Expected: `ok`

- [ ] **Step 3: Commit**

```bash
git add api/cap_client.py
git commit -m "feat: add analyze_excel_structure() to CapClient"
```

---

## Task 7: Integrate Phase 1 into `match_worker()`

**Files:**
- Modify: `d:/Users/PC/Projects/if_mapping_gui/gui/frames/match_frame.py`

- [ ] **Step 1: Update the import at the top of `match_frame.py`**

Add `build_sheet_previews` to the existing reader import line:

```python
from excel.reader import read_fields, build_sheet_previews, ExcelReadError
```

- [ ] **Step 2: Replace the `read_fields` call in `match_worker()` (around line 141)**

Find the block inside the `for idx, file_path in enumerate(files):` loop:

```python
        try:
            fields, workbook, direction = read_fields(file_path, excel_cfg)
```

Replace it with:

```python
        try:
            # Phase 1: AI structure detection
            log_f(f"Analyzing structure...")
            try:
                sheet_previews = build_sheet_previews(file_path)
                excel_cfg_file = client.analyze_excel_structure(
                    sheet_previews, provider, language
                )
                # Assemble into the shape read_fields() expects
                detected_direction = excel_cfg_file.get("direction", "normal")
                excel_cfg_resolved = {
                    "sheet_head": excel_cfg_file.get("sheet_head", excel_cfg["sheet_head"]),
                    "sheet_data": excel_cfg_file.get("sheet_data", excel_cfg["sheet_data"]),
                    "header_row": excel_cfg_file.get("header_row", excel_cfg["header_row"]),
                    "start_row":  excel_cfg_file.get("start_row",  excel_cfg["start_row"]),
                    "detection":  excel_cfg["detection"],
                    "directions": {
                        detected_direction: {
                            "input_header_cols": excel_cfg_file.get("header_cols", {}),
                            "input_row_cols":    excel_cfg_file.get("row_cols", {}),
                        }
                    },
                }
            except Exception as ai_err:
                log_f(f"AI structure analysis failed ({ai_err}), falling back to config", "warn")
                excel_cfg_resolved = excel_cfg

            # Phase 2: field extraction
            fields, workbook, direction = read_fields(file_path, excel_cfg_resolved)
```

Note: the try/except block for `ExcelReadError` that was already there (line 146-148) remains unchanged — it wraps both Phase 1 and Phase 2.

- [ ] **Step 3: Also pass `excel_cfg_resolved` direction to `_export` — no change needed**

The `_export` method uses `self._directions` (populated from `read_fields` return value) and `self.app.cfg["excel"]` for `output_cols`. `output_cols` comes from `EXCEL_DEFAULTS` via config and is **not** changed by AI analysis — only the input column mapping changes. Verify by reading line 458:

```python
dir_cfg = excel_cfg["directions"][direction]
output_cols = dir_cfg["output_cols"]
```

`excel_cfg` here is `self.app.cfg.get("excel", config.EXCEL_DEFAULTS)` — unchanged. This is correct: AI analysis only affects *reading*, not *writing*.

- [ ] **Step 4: Manual smoke test**

1. Start the CAP backend: `cd d:/Users/PC/Projects/if_mapping_cap && npm run start`
2. Start the GUI: `cd d:/Users/PC/Projects/if_mapping_gui && python gui_main.py`
3. Place a test Excel file in the `input/` directory
4. Click "Run Match" — observe log shows "Analyzing structure..." before "Reading fields..."
5. Confirm output Excel is written to `output/` directory

- [ ] **Step 5: Commit**

```bash
cd d:/Users/PC/Projects/if_mapping_gui
git add gui/frames/match_frame.py
git commit -m "feat: insert AI Phase 1 structure analysis into match_worker"
```

---

## Self-Review Checklist

- [x] **Spec § 3 (Architecture)**: Phase 1 → Phase 2 flow implemented in Tasks 5-7. `read_fields()` unchanged. ✓
- [x] **Spec § 4.1 (CDS Types)**: `SheetPreview`, `HeaderCols`, `RowCols`, `ExcelStructure` added in Task 1. ✓
- [x] **Spec § 4.2 (Action)**: `analyzeExcelStructure` declared in Task 2. ✓
- [x] **Spec § 4.3 (GUI client)**: `analyze_excel_structure()` added in Task 6. ✓
- [x] **Spec § 5.1 (`build_sheet_previews`)**: Implemented with `PREVIEW_ROWS=15`, skip empty sheets, `read_only=True` in Task 5. ✓
- [x] **Spec § 5.2 (match_worker)**: Phase 1 inserted, AI fallback to config on error in Task 7. ✓
- [x] **Spec § 6.4 (analyzer.ts)**: `TOOL_SCHEMA` inline (not from DB), `callWithTools` call pattern matches step3-field-matching.ts. ✓
- [x] **Spec § 8 (Error handling)**: AI failure → warn log + fallback to `excel_cfg` config in Task 7 step 2. `callWithTools` no `toolInput` → AppError 422 in Task 3. ✓
- [x] **Spec § 9 (Tests)**: 5 unit tests for `build_sheet_previews` in Task 5. ✓
- [x] **Type consistency**: `ExcelStructure.header_cols` / `ExcelStructure.row_cols` in CDS (Task 1), TypeScript interface (Task 3), and GUI assembly (Task 7) all use same field names. ✓
- [x] **`output_cols` not changed**: `_export` still reads from `self.app.cfg["excel"]` (EXCEL_DEFAULTS). AI only affects input parsing. ✓
