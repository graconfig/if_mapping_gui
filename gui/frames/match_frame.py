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
    files: list[str],
    provider: str,
    language: str,
    server_url: str,
    timeout: int,
    excel_cfg: dict,
    q: queue.Queue,
    stop: threading.Event,
) -> None:
    """Background worker: runs matching pipeline and puts messages on q."""

    def log(text, level="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        q.put({"type": "log", "text": f"[{ts}]  {level.upper():<5} {text}", "level": level})

    client = CapClient(server_url, timeout=timeout)
    if not client.ping():
        log(i18n.t("match.conn_fail_log", url=server_url), "error")
        q.put({"type": "error", "msg": "CAP service unreachable"})
        return

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

    try:
        for idx, file_str in enumerate(files):
            if stop.is_set():
                log(i18n.t("match.user_stop_log"), "warn")
                break

            path = Path(file_str)
            try:
                fields, workbook, direction = read_fields(path, excel_cfg)
                all_wb[file_str] = workbook
                all_row_indices[file_str] = {f.rowIndex for f in fields}
                all_directions[file_str] = direction
                log(i18n.t("match.parse_log", name=path.name, count=str(len(fields)), direction=direction))
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
                log(i18n.t("match.match_fail_log", name=path.name, error=str(e)), "error")

            q.put({"type": "progress", "pct": (idx + 1) / len(files)})

        q.put({
            "type": "done",
            "results": all_results,
            "wb": all_wb,
            "row_indices": all_row_indices,
            "directions": all_directions,
            "files": files,
        })
    finally:
        sse_stop.set()


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
            self, text=i18n.t("match.pick_files"),
            height=72, fg_color="#1e293b", hover_color="#334155",
            text_color="#64748b", command=self._pick_files,
        )
        self._drop_zone.pack(fill="x", **pad)

        self._file_list_frame = ctk.CTkScrollableFrame(self, height=80, fg_color="#1e293b")
        self._file_list_frame.pack(fill="x", padx=16, pady=(0, 6))

        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", padx=16, pady=(0, 6))
        self._provider_label = ctk.CTkLabel(opts, text=i18n.t("match.provider_label"), font=("", 10), text_color="gray")
        self._provider_label.grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=120).grid(row=1, column=0, padx=(0, 10))
        self._lang_label = ctk.CTkLabel(opts, text=i18n.t("match.lang_label"), font=("", 10), text_color="gray")
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

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title=i18n.t("match.pick_title"),
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
            args=(list(self._files), self._provider_var.get(), self._lang_var.get(),
                  self.app.cfg.get("server_url", "http://localhost:4004"),
                  self.app.cfg.get("timeout", 600),
                  excel_cfg, self._queue, self._stop_event),
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
        for file_str in self._files:
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
            out = write_results(Path(file_str), workbook, file_results, output_cols, sheet_data, input_row_cols)
            self._log_append(i18n.t("match.export_done_log", name=out.name), "info")

    def retranslate(self) -> None:
        self._drop_zone.configure(text=i18n.t("match.pick_files"))
        self._provider_label.configure(text=i18n.t("match.provider_label"))
        self._lang_label.configure(text=i18n.t("match.lang_label"))
        self._start_btn.configure(text=i18n.t("match.start_btn"))
        self._stop_btn.configure(text=i18n.t("match.stop_btn"))
        self._export_btn.configure(text=i18n.t("match.export_btn"))
