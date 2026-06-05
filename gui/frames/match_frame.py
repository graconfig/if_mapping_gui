# gui/frames/match_frame.py
import json
import queue
import threading
import uuid
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

import openpyxl
import customtkinter as ctk

import config
import i18n
from api.cap_client import CapClient
from excel.reader import read_fields, ExcelReadError
from excel.writer import write_results
from gui.frames import BaseFrame

LOG_COLORS = {"info": "#22c55e", "warn": "#f59e0b", "error": "#ef4444", "step": "#7b8cde"}

_EXCEL_GLOB = ("*.xlsx", "*.xls")


def _scan_excel(folder: Path) -> list[Path]:
    files = []
    for pat in _EXCEL_GLOB:
        files.extend(folder.glob(pat))
    return sorted(files)


def _sse_worker(client: CapClient, correlation_id: str, q: queue.Queue, stop: threading.Event) -> None:
    """Read CAP log stream (SSE) and forward entries to the main queue."""
    try:
        r = client.open_log_stream(correlation_id)
        r.raise_for_status()
        buf = b""
        for chunk in r.iter_content(chunk_size=None):
            if stop.is_set():
                break
            buf += chunk
            while b"\n\n" in buf:
                raw, buf = buf.split(b"\n\n", 1)
                for line in raw.splitlines():
                    if not line.startswith(b"data: "):
                        continue
                    try:
                        entry = json.loads(line[6:])
                    except json.JSONDecodeError:
                        continue
                    level = entry.get("level", "info")
                    msg   = entry.get("message", "")
                    ts    = entry.get("timestamp", "")[:19].replace("T", " ")
                    ctx   = entry.get("context") or {}
                    extra = "  " + "  ".join(
                        f"{k}={v}" for k, v in ctx.items()
                        if k != "correlationId" and v is not None
                    )
                    q.put({
                        "type":  "log",
                        "text":  f"[{ts}] {level.upper():<5} {msg}{extra if extra.strip() else ''}",
                        "level": level,
                    })
    except Exception:
        pass


def match_worker(
    input_dir: str,
    output_dir: str,
    provider: str,
    language: str,
    server_url: str,
    timeout: int,
    excel_cfg: dict,
    q: queue.Queue,
    stop: threading.Event,
    xsuaa: dict | None = None,
) -> None:
    """Background worker: scans input_dir for Excel files, runs matching, writes to output_dir."""

    def log(text, level="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        q.put({"type": "log", "text": f"[{ts}]  {level.upper():<5} {text}", "level": level})

    client = CapClient(server_url, timeout=timeout, xsuaa=xsuaa)
    if not client.ping():
        log(i18n.t("match.conn_fail_log", url=server_url), "error")
        q.put({"type": "error", "msg": "CAP service unreachable"})
        return

    in_path = Path(input_dir)
    out_path = Path(output_dir)
    files = _scan_excel(in_path)

    if not files:
        log(i18n.t("match.no_excel_in_dir"), "error")
        q.put({"type": "error", "msg": "no Excel files"})
        return

    log(i18n.t("match.scan_log", dir=in_path.name, count=str(len(files))))

    # Start SSE log stream for this run
    correlation_id = uuid.uuid4().hex[:12]
    sse_stop = threading.Event()
    sse_thread = threading.Thread(
        target=_sse_worker,
        args=(client, correlation_id, q, sse_stop),
        daemon=True,
    )
    sse_thread.start()

    all_results: list[dict] = []
    all_wb: dict[str, openpyxl.Workbook] = {}
    all_row_indices: dict[str, set[int]] = {}
    all_directions: dict[str, str] = {}
    file_strs = [str(f) for f in files]

    try:
        for idx, file_path in enumerate(files):
            if stop.is_set():
                log(i18n.t("match.user_stop_log"), "warn")
                break

            try:
                fields, workbook, direction = read_fields(file_path, excel_cfg)
                all_wb[str(file_path)] = workbook
                all_row_indices[str(file_path)] = {f.rowIndex for f in fields}
                all_directions[str(file_path)] = direction
                log(i18n.t("match.parse_log", name=file_path.name, count=str(len(fields)), direction=direction))
            except ExcelReadError as e:
                log(str(e), "error")
                continue

            try:
                results = client.match(
                    [f.to_dict() for f in fields],
                    provider=provider,
                    language=language,
                    correlation_id=correlation_id,
                )
                all_results.extend(results)
                custom = sum(1 for r in results if r.get("matchSource") in ("exact", "vector"))
                ai     = sum(1 for r in results if r.get("matchSource") == "ai")
                log(i18n.t("match.match_done_log", total=str(len(results)), custom=str(custom), ai=str(ai)), "step")
            except Exception as e:
                log(i18n.t("match.match_fail_log", name=file_path.name, error=str(e)), "error")

            q.put({"type": "progress", "pct": (idx + 1) / len(files)})

        q.put({
            "type": "done",
            "results": all_results,
            "wb": all_wb,
            "row_indices": all_row_indices,
            "directions": all_directions,
            "files": file_strs,
            "output_dir": str(out_path),
        })
    finally:
        sse_stop.set()


class MatchFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._input_dir: str = ""
        self._output_dir: str = ""
        self._results: list[dict] = []
        self._wb: dict[str, openpyxl.Workbook] = {}
        self._row_indices: dict[str, set[int]] = {}
        self._directions: dict[str, str] = {}
        self._last_output_dir: str = ""
        self._queue: queue.Queue = queue.Queue()
        self._stop_event = threading.Event()
        self._build()

    def _build(self):
        pad = {"padx": 16, "pady": 6}

        # ── Directory pickers ─────────────────────────────────────────────────
        dir_frame = ctk.CTkFrame(self, fg_color="#1e293b", corner_radius=6)
        dir_frame.pack(fill="x", padx=16, pady=(10, 4))

        # Input dir row
        in_row = ctk.CTkFrame(dir_frame, fg_color="transparent")
        in_row.pack(fill="x", padx=10, pady=(8, 4))
        self._input_dir_label = ctk.CTkLabel(in_row, text=i18n.t("match.input_dir_label"),
                                             font=("", 11), width=80, anchor="w")
        self._input_dir_label.pack(side="left")
        self._input_dir_btn = ctk.CTkButton(in_row, text="📂", width=32, height=26,
                                            command=self._pick_input_dir)
        self._input_dir_btn.pack(side="left", padx=(0, 6))
        self._input_dir_display = ctk.CTkLabel(in_row, text=self._get_default_input_display(),
                                               font=("Consolas", 10), text_color="#94a3b8", anchor="w")
        self._input_dir_display.pack(side="left", fill="x", expand=True)

        # Output dir row
        out_row = ctk.CTkFrame(dir_frame, fg_color="transparent")
        out_row.pack(fill="x", padx=10, pady=(0, 8))
        self._output_dir_label = ctk.CTkLabel(out_row, text=i18n.t("match.output_dir_label"),
                                              font=("", 11), width=80, anchor="w")
        self._output_dir_label.pack(side="left")
        self._output_dir_btn = ctk.CTkButton(out_row, text="📂", width=32, height=26,
                                             command=self._pick_output_dir)
        self._output_dir_btn.pack(side="left", padx=(0, 6))
        self._output_dir_display = ctk.CTkLabel(out_row, text=self._get_default_output_display(),
                                                font=("Consolas", 10), text_color="#94a3b8", anchor="w")
        self._output_dir_display.pack(side="left", fill="x", expand=True)

        # Set defaults from config
        self._apply_default_dirs()

        # ── Options bar ───────────────────────────────────────────────────────
        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", padx=16, pady=(0, 6))
        self._provider_label = ctk.CTkLabel(opts, text=i18n.t("match.provider_label"),
                                            font=("", 10), text_color="gray")
        self._provider_label.grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=120).grid(row=1, column=0, padx=(0, 10))
        self._lang_label = ctk.CTkLabel(opts, text=i18n.t("match.lang_label"),
                                        font=("", 10), text_color="gray")
        self._lang_label.grid(row=0, column=1, sticky="w")
        self._lang_var = ctk.StringVar(value=self.app.cfg.get("language", "ja"))
        ctk.CTkOptionMenu(opts, variable=self._lang_var,
                          values=["ja", "en", "zh"], width=100).grid(row=1, column=1, padx=(0, 10))
        self._start_btn = ctk.CTkButton(opts, text=i18n.t("match.start_btn"), width=100, command=self._start)
        self._start_btn.grid(row=1, column=2, padx=(10, 6))
        self._stop_btn = ctk.CTkButton(opts, text=i18n.t("match.stop_btn"), width=80,
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
        self._export_btn = ctk.CTkButton(result_bar, text=i18n.t("match.export_btn"),
                                          width=140, command=self._export, state="disabled")
        self._export_btn.pack(side="right")

    # ── Default dir helpers ───────────────────────────────────────────────────

    def _get_default_input_dir(self) -> str:
        saved = self.app.cfg.get("last_input_dir", "")
        if saved and Path(saved).is_dir():
            return saved
        d = config.PROJECT_ROOT / "input"
        return str(d) if d.is_dir() else str(config.PROJECT_ROOT)

    def _get_default_output_dir(self) -> str:
        saved = self.app.cfg.get("last_output_dir", "")
        if saved and Path(saved).is_dir():
            return saved
        d = config.PROJECT_ROOT / "output"
        return str(d) if d.is_dir() else str(config.PROJECT_ROOT)

    def _get_default_input_display(self) -> str:
        d = self._get_default_input_dir()
        return Path(d).name if d else i18n.t("match.no_dir_selected")

    def _get_default_output_display(self) -> str:
        d = self._get_default_output_dir()
        return Path(d).name if d else i18n.t("match.no_dir_selected")

    def _apply_default_dirs(self) -> None:
        self._input_dir = self._get_default_input_dir()
        self._output_dir = self._get_default_output_dir()
        self._update_dir_display()

    def _update_dir_display(self) -> None:
        if self._input_dir:
            self._input_dir_display.configure(text=str(Path(self._input_dir)))
        if self._output_dir:
            self._output_dir_display.configure(text=str(Path(self._output_dir)))

    # ── Dir pickers ───────────────────────────────────────────────────────────

    def _pick_input_dir(self):
        initial = self._input_dir or self.app.cfg.get("last_input_dir") or str(config.PROJECT_ROOT)
        d = filedialog.askdirectory(
            title=i18n.t("match.pick_input_dir"),
            initialdir=initial,
        )
        if d:
            self._input_dir = d
            self.app.cfg["last_input_dir"] = d
            config.save(self.app.cfg)
            self._input_dir_display.configure(text=str(Path(d)))

    def _pick_output_dir(self):
        initial = self._output_dir or self.app.cfg.get("last_output_dir") or str(config.PROJECT_ROOT)
        d = filedialog.askdirectory(
            title=i18n.t("match.pick_output_dir"),
            initialdir=initial,
        )
        if d:
            self._output_dir = d
            self.app.cfg["last_output_dir"] = d
            config.save(self.app.cfg)
            self._output_dir_display.configure(text=str(Path(d)))

    # ── Controls ──────────────────────────────────────────────────────────────

    def _log_append(self, text: str, level: str = "info"):
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _start(self):
        if not self._input_dir or not Path(self._input_dir).is_dir():
            self._log_append(i18n.t("match.no_file_error"), "error")
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
        thr = threading.Thread(
            target=match_worker,
            args=(self._input_dir, self._output_dir,
                  self._provider_var.get(), self._lang_var.get(),
                  self.app.cfg.get("server_url", "http://localhost:4004"),
                  self.app.cfg.get("timeout", 600),
                  excel_cfg, self._queue, self._stop_event,
                  self.app.cfg.get("xsuaa")),
            daemon=True,
        )
        thr.start()
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
                    self._last_output_dir = msg.get("output_dir", self._output_dir)
                    self._on_done()
                    return
                elif msg["type"] == "error":
                    self._on_error()
                    return
        except Exception:
            pass
        self.after(100, self._poll)

    def _on_done(self):
        custom = sum(1 for r in self._results if r.get("matchSource") in ("exact", "vector"))
        ai     = sum(1 for r in self._results if r.get("matchSource") == "ai")
        self._result_label.configure(
            text=i18n.t("match.result_summary", total=str(len(self._results)), custom=str(custom), ai=str(ai))
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
        out_dir = Path(self._last_output_dir) if self._last_output_dir else Path(self._output_dir)
        for file_str in list(self._wb.keys()):
            workbook = self._wb.get(file_str)
            if workbook is None:
                continue
            row_idxs = self._row_indices.get(file_str, set())
            file_results = [r for r in self._results if r.get("rowIndex") in row_idxs]
            direction = self._directions.get(file_str, "normal")
            dir_cfg = excel_cfg["directions"][direction]
            output_cols = dir_cfg["output_cols"]
            input_row_cols = dir_cfg.get("input_row_cols")
            sheet_data = excel_cfg.get("sheet_data", "IFマッピング定義")
            out = write_results(Path(file_str), workbook, file_results, output_cols,
                                sheet_data, input_row_cols, out_dir)
            self._log_append(i18n.t("match.export_done_log", name=out.name), "info")

    def retranslate(self) -> None:
        self._input_dir_label.configure(text=i18n.t("match.input_dir_label"))
        self._output_dir_label.configure(text=i18n.t("match.output_dir_label"))
        self._provider_label.configure(text=i18n.t("match.provider_label"))
        self._lang_label.configure(text=i18n.t("match.lang_label"))
        self._start_btn.configure(text=i18n.t("match.start_btn"))
        self._stop_btn.configure(text=i18n.t("match.stop_btn"))
        self._export_btn.configure(text=i18n.t("match.export_btn"))
