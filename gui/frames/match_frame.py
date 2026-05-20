# gui/frames/match_frame.py
import queue
import threading
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

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
    all_raw: dict[str, list[dict]] = {}

    for i, file_str in enumerate(files):
        if stop.is_set():
            log("用户停止 — 处理中断", "warn")
            break

        path = Path(file_str)
        try:
            fields, raw_rows = read_fields(path)
            all_raw[file_str] = raw_rows
            log(f"解析 {path.name} — {len(fields)} 条字段")
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
            exact = sum(1 for r in results if r.get("matchType") == "exact")
            vector = sum(1 for r in results if r.get("matchType") == "vector")
            ai = sum(1 for r in results if r.get("matchType") == "ai")
            log(f"匹配完成: {len(results)} 条 (精确 {exact} · 向量 {vector} · AI {ai})", "step")
        except Exception as e:
            log(f"{path.name} 匹配失败: {e}", "error")

        q.put({"type": "progress", "pct": (i + 1) / len(files)})

    q.put({"type": "done", "results": all_results, "raw": all_raw, "files": files})


class MatchFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._files: list[str] = []
        self._results: list[dict] = []
        self._raw: dict[str, list[dict]] = {}
        self._queue: queue.Queue = queue.Queue()
        self._stop_event = threading.Event()
        self._build()

    def _build(self):
        pad = {"padx": 16, "pady": 6}

        # Drop zone (click to select)
        self._drop_zone = ctk.CTkButton(
            self, text="📂  点击选择 Excel 文件（可多选）\n支持 .xlsx / .xls",
            height=72, fg_color="#1e293b", hover_color="#334155",
            text_color="#64748b", command=self._pick_files,
        )
        self._drop_zone.pack(fill="x", **pad)

        # File list
        self._file_list_frame = ctk.CTkScrollableFrame(self, height=80, fg_color="#1e293b")
        self._file_list_frame.pack(fill="x", padx=16, pady=(0, 6))

        # Options row
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

        # Progress bar
        self._progress = ctk.CTkProgressBar(self)
        self._progress.set(0)
        self._progress.pack(fill="x", padx=16, pady=(0, 6))

        # Log area
        self._log = ctk.CTkTextbox(self, height=160, font=("Consolas", 10), state="disabled")
        self._log.pack(fill="both", expand=True, padx=16, pady=(0, 6))

        # Result + export bar
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
        color = LOG_COLORS.get(level, "white")
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _start(self):
        if not self._files:
            self._log_append("[ERROR] 请先选择 Excel 文件", "error")
            return
        self._results.clear()
        self._raw.clear()
        self._stop_event.clear()
        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._export_btn.configure(state="disabled")
        self._progress.set(0)
        self._progress.configure(progress_color=("#7b8cde", "#7b8cde"))

        t = threading.Thread(
            target=match_worker,
            args=(list(self._files), self._provider_var.get(), self._lang_var.get(),
                  self.app.cfg.get("server_url", "http://localhost:4004"),
                  self._queue, self._stop_event),
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
                    self._raw = msg["raw"]
                    self._on_done()
                    return
                elif msg["type"] == "error":
                    self._on_error()
                    return
        except Exception:
            pass
        self.after(100, self._poll)

    def _on_done(self):
        exact = sum(1 for r in self._results if r.get("matchType") == "exact")
        vector = sum(1 for r in self._results if r.get("matchType") == "vector")
        ai = sum(1 for r in self._results if r.get("matchType") == "ai")
        self._result_label.configure(
            text=f"完成: {len(self._results)} 条 ｜ 精确 {exact} ｜ 向量 {vector} ｜ AI {ai}"
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
        for file_str in self._files:
            raw_rows = self._raw.get(file_str)
            if not raw_rows:
                continue
            file_results = [r for r in self._results if r.get("rowIndex") in {rr["rowIndex"] for rr in raw_rows}]
            out = write_results(Path(file_str), raw_rows, file_results)
            self._log_append(f"已导出: {out.name}", "info")
