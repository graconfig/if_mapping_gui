# gui/frames/upload_frame.py
import queue
import threading
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

import customtkinter as ctk

import config
from api.cap_client import CapClient, CapConnectionError
from excel.reader import read_fields, ExcelReadError
from gui.frames import BaseFrame


def upload_worker(
    file_str: str,
    mode: str,
    server_url: str,
    q: queue.Queue,
    stop: threading.Event,
) -> None:
    def log(text, level="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        q.put({"type": "log", "text": f"[{ts}]  {level.upper():<5} {text}", "level": level})

    client = CapClient(server_url)
    if not client.ping():
        log(f"无法连接到 {server_url}", "error")
        q.put({"type": "error"})
        return

    path = Path(file_str)
    try:
        fields, _ = read_fields(path)
        log(f"解析 {path.name} — {len(fields)} 条记录")
    except ExcelReadError as e:
        log(str(e), "error")
        q.put({"type": "error"})
        return

    try:
        records = [f.to_dict() for f in fields]
        result = client.upload_custom_fields(records, mode=mode)
        inserted = result.get("inserted", 0)
        updated = result.get("updated", 0)
        deleted = result.get("deleted", 0)
        log(f"上传完成: 插入 {inserted}, 更新 {updated}, 删除 {deleted}")
        q.put({"type": "done"})
    except CapConnectionError as e:
        log(f"上传失败: {e}", "error")
        q.put({"type": "error"})


class UploadFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._file: str = ""
        self._queue: queue.Queue = queue.Queue()
        self._stop_event = threading.Event()
        self._build()

    def _build(self):
        pad = {"padx": 16, "pady": 8}

        ctk.CTkLabel(self, text="上传知识库", font=("", 16, "bold")).pack(anchor="w", padx=20, pady=(20, 4))

        # Drop zone
        self._drop_btn = ctk.CTkButton(
            self, text="📂  点击选择知识库 Excel 文件", height=60,
            fg_color="#1e293b", hover_color="#334155", text_color="#64748b",
            command=self._pick_file,
        )
        self._drop_btn.pack(fill="x", **pad)

        self._file_label = ctk.CTkLabel(self, text="", font=("", 11), text_color="#94a3b8")
        self._file_label.pack(anchor="w", padx=16)

        # Mode + upload row
        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", **pad)
        ctk.CTkLabel(opts, text="上传模式", font=("", 10), text_color="gray").pack(side="left", padx=(0, 6))
        self._mode_var = ctk.StringVar(value="upsert")
        ctk.CTkOptionMenu(opts, variable=self._mode_var,
                          values=["upsert", "overwrite"], width=160).pack(side="left", padx=(0, 12))
        self._upload_btn = ctk.CTkButton(opts, text="⬆ 上传", width=90, command=self._start)
        self._upload_btn.pack(side="left")

        # Log
        self._log = ctk.CTkTextbox(self, height=160, font=("Consolas", 10), state="disabled")
        self._log.pack(fill="both", expand=True, padx=16, pady=(6, 16))

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="选择知识库 Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=self.app.cfg.get("last_input_dir") or None,
        )
        if path:
            self._file = path
            self._file_label.configure(text=f"📄 {Path(path).name}")

    def _log_append(self, text: str):
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _start(self):
        if not self._file:
            self._log_append("[ERROR] 请先选择文件")
            return
        self._upload_btn.configure(state="disabled")
        self._stop_event.clear()
        threading.Thread(
            target=upload_worker,
            args=(self._file, self._mode_var.get(),
                  self.app.cfg.get("server_url", "http://localhost:4004"),
                  self._queue, self._stop_event),
            daemon=True,
        ).start()
        self.after(100, self._poll)

    def _poll(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                if msg["type"] == "log":
                    self._log_append(msg["text"])
                elif msg["type"] in ("done", "error"):
                    self._upload_btn.configure(state="normal")
                    return
        except Exception:
            pass
        self.after(100, self._poll)
