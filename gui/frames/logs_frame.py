# gui/frames/logs_frame.py
import threading
import customtkinter as ctk
from gui.frames import BaseFrame


class LogsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(header, text="Token 日志", font=("", 16, "bold")).pack(side="left")
        ctk.CTkButton(header, text="🔄 刷新", width=80, command=self._load).pack(side="right")

        # Summary chips
        chips = ctk.CTkFrame(self, fg_color="transparent")
        chips.pack(fill="x", padx=16, pady=(0, 12))
        self._input_label = self._chip(chips, "0", "总 Input Tokens")
        self._input_label.pack(side="left", padx=(0, 10))
        self._output_label = self._chip(chips, "0", "总 Output Tokens")
        self._output_label.pack(side="left", padx=(0, 10))
        self._calls_label = self._chip(chips, "0", "调用次数")
        self._calls_label.pack(side="left")

        # Table header
        header_row = ctk.CTkFrame(self, fg_color="#1e293b", height=28)
        header_row.pack(fill="x", padx=16)
        for col, w in [("时间", 80), ("Provider", 70), ("Step", 120), ("Input", 70), ("Output", 70)]:
            ctk.CTkLabel(header_row, text=col, font=("", 10, "bold"), width=w).pack(side="left", padx=4)

        # Scrollable table body
        self._table = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._table.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self._load()

    def _chip(self, parent, value: str, label: str) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color="#1e293b", corner_radius=6)
        val_lbl = ctk.CTkLabel(frame, text=value, font=("", 18, "bold"), text_color="#7b8cde")
        val_lbl.pack(padx=14, pady=(8, 0))
        ctk.CTkLabel(frame, text=label, font=("", 10), text_color="gray").pack(padx=14, pady=(0, 8))
        frame._value_label = val_lbl
        return frame

    def _load(self):
        def _fetch():
            try:
                logs = self.app.get_client().get_token_logs()
                self.after(0, lambda: self._populate(logs))
            except Exception:
                pass
        threading.Thread(target=_fetch, daemon=True).start()

    def _populate(self, logs: list[dict]):
        for w in self._table.winfo_children():
            w.destroy()

        total_in = sum(r.get("inputTokens", 0) for r in logs)
        total_out = sum(r.get("outputTokens", 0) for r in logs)
        self._input_label._value_label.configure(text=f"{total_in:,}")
        self._output_label._value_label.configure(text=f"{total_out:,}")
        self._calls_label._value_label.configure(text=str(len(logs)))

        for entry in logs:
            row = ctk.CTkFrame(self._table, fg_color="transparent")
            row.pack(fill="x", pady=1)
            ts = str(entry.get("createdAt", ""))[:19].replace("T", " ")
            for val, w in [
                (ts, 80),
                (entry.get("provider", ""), 70),
                (entry.get("step", ""), 120),
                (str(entry.get("inputTokens", "")), 70),
                (str(entry.get("outputTokens", "")), 70),
            ]:
                ctk.CTkLabel(row, text=val, font=("Consolas", 10), width=w,
                             text_color="#94a3b8").pack(side="left", padx=4)
