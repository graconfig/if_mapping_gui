# gui/frames/logs_frame.py
import threading
import customtkinter as ctk
import i18n
from gui.frames import BaseFrame


class LogsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        self._title_label = ctk.CTkLabel(header, text=i18n.t("logs.title"), font=("", 16, "bold"))
        self._title_label.pack(side="left")
        self._refresh_btn = ctk.CTkButton(header, text=i18n.t("logs.refresh_btn"), width=80, command=self._load)
        self._refresh_btn.pack(side="right")

        # Summary chips
        chips = ctk.CTkFrame(self, fg_color="transparent")
        chips.pack(fill="x", padx=16, pady=(0, 12))
        self._input_label = self._chip(chips, "0", i18n.t("logs.total_input"))
        self._input_label.pack(side="left", padx=(0, 10))
        self._output_label = self._chip(chips, "0", i18n.t("logs.total_output"))
        self._output_label.pack(side="left", padx=(0, 10))
        self._calls_label = self._chip(chips, "0", i18n.t("logs.call_count"))
        self._calls_label.pack(side="left")

        # Table header
        self._header_row = ctk.CTkFrame(self, fg_color="#1e293b", height=28)
        self._header_row.pack(fill="x", padx=16)
        self._col_labels: list[tuple] = []
        for col_key, w in [
            ("logs.col_time", 80), ("logs.col_provider", 70), ("logs.col_step", 120),
            ("logs.col_input", 70), ("logs.col_output", 70)
        ]:
            lbl = ctk.CTkLabel(self._header_row, text=i18n.t(col_key), font=("", 10, "bold"), width=w)
            lbl.pack(side="left", padx=4)
            self._col_labels.append((lbl, col_key))

        # Scrollable table body
        self._table = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._table.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self._load()

    def _chip(self, parent, value: str, label: str) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color="#1e293b", corner_radius=6)
        val_lbl = ctk.CTkLabel(frame, text=value, font=("", 18, "bold"), text_color="#7b8cde")
        val_lbl.pack(padx=14, pady=(8, 0))
        text_lbl = ctk.CTkLabel(frame, text=label, font=("", 10), text_color="gray")
        text_lbl.pack(padx=14, pady=(0, 8))
        frame._value_label = val_lbl
        frame._text_label = text_lbl
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

    def retranslate(self) -> None:
        self._title_label.configure(text=i18n.t("logs.title"))
        self._refresh_btn.configure(text=i18n.t("logs.refresh_btn"))
        self._input_label._text_label.configure(text=i18n.t("logs.total_input"))
        self._output_label._text_label.configure(text=i18n.t("logs.total_output"))
        self._calls_label._text_label.configure(text=i18n.t("logs.call_count"))
        for lbl, col_key in self._col_labels:
            lbl.configure(text=i18n.t(col_key))
