# gui/frames/upload_frame.py
import queue
import threading
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

import customtkinter as ctk

import config
import i18n
from api.cap_client import CapClient, CapConnectionError
from excel.reader import read_kb_fields, ExcelReadError
from gui.frames import BaseFrame


def upload_worker(
    file_str: str,
    mode: str,
    server_url: str,
    q: queue.Queue,
    stop: threading.Event,
    kb_cfg: dict | None = None,
) -> None:
    def log(text, level="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        q.put({"type": "log", "text": f"[{ts}]  {level.upper():<5} {text}", "level": level})

    client = CapClient(server_url)
    if not client.ping():
        log(i18n.t("upload.conn_fail_log", url=server_url), "error")
        q.put({"type": "error"})
        return

    path = Path(file_str)
    try:
        records = read_kb_fields(path, kb_cfg=kb_cfg)
        log(i18n.t("upload.parse_log", name=path.name, count=str(len(records))))
    except ExcelReadError as e:
        log(str(e), "error")
        q.put({"type": "error"})
        return

    try:
        result = client.upload_custom_fields(records, mode=mode)
        inserted = result.get("inserted", 0)
        updated = result.get("updated", 0)
        deleted = result.get("deleted", 0)
        log(i18n.t("upload.done_log", inserted=str(inserted), updated=str(updated), deleted=str(deleted)))
        q.put({"type": "done"})
    except CapConnectionError as e:
        log(i18n.t("upload.fail_log", error=str(e)), "error")
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

        self._title_label = ctk.CTkLabel(self, text=i18n.t("upload.title"), font=("", 16, "bold"))
        self._title_label.pack(anchor="w", padx=20, pady=(20, 4))

        # Drop zone
        self._drop_btn = ctk.CTkButton(
            self, text=i18n.t("upload.pick_btn"), height=60,
            fg_color="#1e293b", hover_color="#334155", text_color="#64748b",
            command=self._pick_file,
        )
        self._drop_btn.pack(fill="x", **pad)

        self._file_label = ctk.CTkLabel(self, text="", font=("", 11), text_color="#94a3b8")
        self._file_label.pack(anchor="w", padx=16)

        # Mode + upload row
        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", **pad)
        self._mode_label = ctk.CTkLabel(opts, text=i18n.t("upload.mode_label"), font=("", 10), text_color="gray")
        self._mode_label.pack(side="left", padx=(0, 6))
        self._mode_var = ctk.StringVar(value="upsert")
        ctk.CTkOptionMenu(opts, variable=self._mode_var,
                          values=["upsert", "overwrite"], width=160).pack(side="left", padx=(0, 12))
        self._upload_btn = ctk.CTkButton(opts, text=i18n.t("upload.upload_btn"), width=90, command=self._start)
        self._upload_btn.pack(side="left")

        # Log
        self._log = ctk.CTkTextbox(self, height=160, font=("Consolas", 10), state="disabled")
        self._log.pack(fill="both", expand=True, padx=16, pady=(6, 16))

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title=i18n.t("upload.pick_title"),
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
            self._log_append(i18n.t("upload.no_file_error"))
            return
        self._upload_btn.configure(state="disabled")
        self._stop_event.clear()
        threading.Thread(
            target=upload_worker,
            args=(self._file, self._mode_var.get(),
                  self.app.cfg.get("server_url", "http://localhost:4004"),
                  self._queue, self._stop_event,
                  self.app.cfg.get("kb_upload")),
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

    def retranslate(self) -> None:
        self._title_label.configure(text=i18n.t("upload.title"))
        self._drop_btn.configure(text=i18n.t("upload.pick_btn"))
        self._mode_label.configure(text=i18n.t("upload.mode_label"))
        self._upload_btn.configure(text=i18n.t("upload.upload_btn"))
