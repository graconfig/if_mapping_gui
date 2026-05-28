# IF Mapping GUI

Local desktop GUI for the `if_mapping_cap` CAP service. Lets business users and consultants run SAP field matching, upload knowledge bases, manage prompt templates, and view token usage logs — without needing a terminal.

Built with Python 3.10+ and CustomTkinter. Communicates with the CAP backend via OData HTTP. Reads/writes Excel files with openpyxl.

---

## Screenshots

| 字段匹配 (Field Matching) | ⚙ 设置 (Settings) |
|--------------------------|-------------------|
| File picker → progress bar → log → export | CAP URL · provider · language |

---

## Prerequisites

- Python 3.10 or later
- A running instance of `if_mapping_cap` (local or remote)
- Windows (the GUI uses Tkinter, which requires a display)

---

## Setup

### 1. Clone and install dependencies

```bash
cd d:/Users/PC/Projects/if_mapping_gui
pip install -r requirements.txt
```

### 2. Configure the server URL

Copy `config.example.json` to `config.json` and edit as needed:

```bash
cp config.example.json config.json
```

```json
{
  "server_url": "http://localhost:4004",
  "provider": "claude",
  "language": "ja",
  "last_input_dir": ""
}
```

| Key | Default | Description |
|-----|---------|-------------|
| `server_url` | `http://localhost:4004` | Base URL of the CAP service |
| `provider` | `claude` | Default AI provider: `claude` \| `openai` \| `gemini` |
| `language` | `ja` | Default language: `ja` \| `en` \| `zh` |
| `last_input_dir` | `""` | Last-used directory for the file picker (auto-saved) |

Config is auto-saved on every change. Settings are also editable in-app via the ⚙ 设置 frame.

---

## Running

```bash
cd d:/Users/PC/Projects/if_mapping_gui
python gui_main.py
```

The window opens at 920×640 in dark mode. The status bar at the bottom shows the connection state to the CAP service.

---

## Features

### ▶ 字段匹配 — Field Matching

1. Click the drop zone to select one or more `.xlsx` / `.xls` input files
2. Choose the AI provider and language
3. Click **▶ 开始匹配** to start — progress bar and log update in real time
4. Click **■ 停止** to interrupt (partial results are still exportable)
5. Click **📥 导出结果 Excel** to save results

Each input file produces a separate output file named `{original}_matched_{YYYYMMDD}_{HHMMSS}.xlsx`.

**Input Excel columns** (case-insensitive, alias-aware):

| Column | Aliases | Required |
|--------|---------|----------|
| `sourceField` | `field_name`, `フィールド名` | Yes |
| `sourceDesc` | `description`, `説明`, `描述` | Yes |
| `sourceTable` | `table_name`, `テーブル名` | No |

**Output Excel columns** (original columns preserved, result columns appended):

| Column | Values |
|--------|--------|
| `targetField` | Matched CDS field name |
| `targetEntity` | Matched CDS view name |
| `matchType` | `exact` / `vector` / `ai` / `unmatched` |
| `confidence` | Float 0–1 |
| `aiReason` | AI explanation (empty for exact/vector) |

---

### ⬆ 上传知识库 — Knowledge Base Upload

Upload a custom-field knowledge base Excel file to the CAP service.

- **Upsert** — insert new records, update existing ones
- **Overwrite** — replace all records

The log shows inserted / updated / deleted counts on completion.

---

### 📝 Prompt 管理 — Prompt Template Management

View and edit the AI prompt templates stored in the CAP service.

- Filter by language: JA / EN / ZH
- Click a prompt in the list to open it in the editor
- **💾 保存** — save changes to the CAP database
- **取消** — discard unsaved edits
- **🔄 重载** — reload the server-side prompt cache (`POST /if-mapping/reloadPrompts`)

---

### 📊 Token 日志 — Token Usage Log

View AI token consumption. Shows total input tokens, total output tokens, and call count as summary chips, plus a table with per-call detail.

Click **🔄 刷新** to refresh from the CAP service.

---

### ⚙ 设置 — Settings

Configure the connection and defaults:

- **CAP 服务地址** — CAP service URL with a test-connection button
- **默认 Provider** — `claude` / `openai` / `gemini`
- **默认语言** — `ja` / `en` / `zh`
- **💾 保存设置** — persists to `config.json`

---

## Project Structure

```
if_mapping_gui/
├── gui_main.py              # Entry point
├── config.py                # Load/save config.json
├── config.example.json      # Config template
├── requirements.txt
├── api/
│   └── cap_client.py        # HTTP client for all CAP OData endpoints
├── excel/
│   ├── reader.py            # Parse input Excel → InterfaceFieldInput list
│   └── writer.py            # Write match results → output Excel
├── gui/
│   ├── app.py               # Main window, sidebar, status bar, frame switching
│   └── frames/
│       ├── __init__.py      # BaseFrame base class
│       ├── match_frame.py   # Field matching workflow + match_worker thread
│       ├── upload_frame.py  # Knowledge base upload + upload_worker thread
│       ├── prompts_frame.py # Prompt list + editor
│       ├── logs_frame.py    # Token usage log table
│       └── settings_frame.py# Connection settings
└── tests/
    ├── test_config.py
    ├── test_cap_client.py
    ├── test_excel_reader.py
    ├── test_excel_writer.py
    ├── test_match_worker.py
    └── test_app_smoke.py
```

---

## Threading Model

All long-running operations (matching, upload) run in `threading.Thread(daemon=True)`. The worker thread communicates with the GUI exclusively via a `queue.Queue`. The main thread polls the queue every 100 ms using `after(100, poll)`.

| Message type | Payload | GUI action |
|-------------|---------|------------|
| `log` | `text`, `level` | Append line to log area |
| `progress` | `pct` (0–1) | Update progress bar |
| `done` | `results` | Show stats, enable Export button |
| `error` | `msg` | Show error in log, re-enable Start button |

---

## Testing

```bash
cd d:/Users/PC/Projects/if_mapping_gui

# Unit tests (no display required)
python -m pytest tests/test_config.py tests/test_cap_client.py tests/test_excel_reader.py tests/test_excel_writer.py tests/test_match_worker.py -v

# All tests including GUI smoke tests (Windows only)
python -m pytest -v
```

---

## CAP API Endpoints Used

| Frame | Endpoint |
|-------|----------|
| Field matching | `POST /if-mapping/match` |
| Knowledge base upload | `POST /if-mapping/uploadCustomFields` |
| Prompt list | `GET /if-mapping/PromptTemplates` |
| Prompt save | `PATCH /if-mapping/PromptTemplates('{id}')` |
| Prompt reload | `POST /if-mapping/reloadPrompts` |
| Token logs | `GET /if-mapping/TokenLogs` |

All calls go through `api/cap_client.py` with a 30-second timeout. Network errors raise `CapConnectionError`.
