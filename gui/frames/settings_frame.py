# gui/frames/settings_frame.py
import threading
import customtkinter as ctk
import config
import i18n
from config import EXCEL_DEFAULTS
from gui.frames import BaseFrame


class SettingsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._build()

    def _build(self):
        outer = ctk.CTkScrollableFrame(self)
        outer.pack(fill="both", expand=True)
        pad = {"padx": 20, "pady": 8}

        self._title_label = ctk.CTkLabel(outer, text=i18n.t("settings.title"), font=("", 16, "bold"))
        self._title_label.pack(anchor="w", padx=20, pady=(20, 4))

        # ── CAP URL ──────────────────────────────────────────────────────────
        url_frame = ctk.CTkFrame(outer, fg_color="transparent")
        url_frame.pack(fill="x", **pad)
        self._cap_url_label = ctk.CTkLabel(url_frame, text=i18n.t("settings.cap_url_label"), font=("", 11))
        self._cap_url_label.pack(anchor="w")
        row = ctk.CTkFrame(url_frame, fg_color="transparent")
        row.pack(fill="x")
        self._url_entry = ctk.CTkEntry(row, width=240)
        self._url_entry.insert(0, self.app.cfg.get("server_url", "http://localhost:4004"))
        self._url_entry.pack(side="left", padx=(0, 8))
        self._timeout_label = ctk.CTkLabel(row, text=i18n.t("settings.timeout_label"), font=("", 10), text_color="#64748b")
        self._timeout_label.pack(side="left")
        self._timeout_entry = ctk.CTkEntry(row, width=60)
        self._timeout_entry.insert(0, str(self.app.cfg.get("timeout", 600)))
        self._timeout_entry.pack(side="left", padx=(4, 8))
        self._test_conn_btn = ctk.CTkButton(row, text=i18n.t("settings.test_conn_btn"), width=110, command=self._test_conn)
        self._test_conn_btn.pack(side="left", padx=(0, 8))
        self._conn_status = ctk.CTkLabel(row, text="", font=("", 11))
        self._conn_status.pack(side="left")

        # ── Provider + Language + UI Language ─────────────────────────────────
        opts_frame = ctk.CTkFrame(outer, fg_color="transparent")
        opts_frame.pack(fill="x", **pad)
        self._provider_label = ctk.CTkLabel(opts_frame, text=i18n.t("settings.provider_label"), font=("", 11))
        self._provider_label.grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts_frame, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=140).grid(row=1, column=0, padx=(0, 20))
        self._lang_label = ctk.CTkLabel(opts_frame, text=i18n.t("settings.lang_label"), font=("", 11))
        self._lang_label.grid(row=0, column=1, sticky="w")
        self._lang_var = ctk.StringVar(value=self.app.cfg.get("language", "ja"))
        ctk.CTkOptionMenu(opts_frame, variable=self._lang_var,
                          values=["ja", "en", "zh"], width=140).grid(row=1, column=1, padx=(0, 20))
        self._ui_lang_label = ctk.CTkLabel(opts_frame, text=i18n.t("settings.ui_lang_label"), font=("", 11))
        self._ui_lang_label.grid(row=0, column=2, sticky="w")
        self._ui_lang_var = ctk.StringVar(value=self.app.cfg.get("ui_language", "zh"))
        ctk.CTkOptionMenu(opts_frame, variable=self._ui_lang_var,
                          values=["zh", "ja"], width=100).grid(row=1, column=2)

        # ── Excel 兜底配置 ────────────────────────────────────────────────────
        self._excel_label = ctk.CTkLabel(outer, text=i18n.t("settings.excel_section"), font=("", 13, "bold"))
        self._excel_label.pack(anchor="w", padx=20, pady=(16, 4))
        excel_section = ctk.CTkFrame(outer, fg_color="#1e293b", corner_radius=8)
        excel_section.pack(fill="x", padx=20, pady=(0, 8))

        excel_cfg = self.app.cfg.get("excel", EXCEL_DEFAULTS)

        def _labeled_entry(parent, label, value, width=160) -> ctk.CTkEntry:
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(side="left", padx=(0, 16))
            ctk.CTkLabel(f, text=label, font=("", 10), text_color="#64748b").pack(anchor="w")
            e = ctk.CTkEntry(f, width=width)
            e.insert(0, str(value))
            e.pack()
            return e

        # sheet names, row numbers, detection (fallback when AI fails)
        global_frame = ctk.CTkFrame(excel_section, fg_color="transparent")
        global_frame.pack(fill="x", padx=12, pady=(10, 4))
        self._sheet_head_entry  = _labeled_entry(global_frame, i18n.t("settings.sheet_head"), excel_cfg["sheet_head"])
        self._sheet_data_entry  = _labeled_entry(global_frame, i18n.t("settings.sheet_data"), excel_cfg["sheet_data"])
        self._header_row_entry  = _labeled_entry(global_frame, i18n.t("settings.header_row"), excel_cfg["header_row"], width=60)
        self._start_row_entry   = _labeled_entry(global_frame, i18n.t("settings.start_row"),  excel_cfg["start_row"],  width=60)

        det_frame = ctk.CTkFrame(excel_section, fg_color="transparent")
        det_frame.pack(fill="x", padx=12, pady=(4, 12))
        det = excel_cfg["detection"]
        self._det_col_entry     = _labeled_entry(det_frame, i18n.t("settings.det_col"),     det["col"],     width=52)
        self._det_row_entry     = _labeled_entry(det_frame, i18n.t("settings.det_row"),     det["row"],     width=60)
        self._det_keyword_entry = _labeled_entry(det_frame, i18n.t("settings.det_keyword"), det["keyword"], width=80)
        self._preview_rows_entry = _labeled_entry(
            det_frame, i18n.t("settings.ai_preview_rows"),
            self.app.cfg.get("ai_preview_rows", 15), width=60,
        )

        self._save_settings_btn = ctk.CTkButton(outer, text=i18n.t("settings.save_btn"), width=120, command=self._save)
        self._save_settings_btn.pack(anchor="e", padx=20, pady=12)

    def _test_conn(self):
        url = self._url_entry.get().strip()
        self._conn_status.configure(text=i18n.t("settings.testing"), text_color="gray")

        def _check():
            from api.cap_client import CapClient
            ok = CapClient(url, xsuaa=self.app.cfg.get("xsuaa")).ping()
            self.after(0, lambda: self._conn_status.configure(
                text=i18n.t("settings.connected") if ok else i18n.t("settings.disconnected"),
                text_color="#22c55e" if ok else "#ef4444",
            ))

        threading.Thread(target=_check, daemon=True).start()

    def _save(self):
        self.app.cfg["server_url"] = self._url_entry.get().strip()
        self.app.cfg["provider"]   = self._provider_var.get()
        self.app.cfg["language"]   = self._lang_var.get()
        new_ui_lang = self._ui_lang_var.get()
        self.app.cfg["ui_language"] = new_ui_lang
        try:
            self.app.cfg["timeout"] = int(self._timeout_entry.get())
        except ValueError:
            pass
        try:
            self.app.cfg["ai_preview_rows"] = int(self._preview_rows_entry.get())
        except ValueError:
            pass

        excel_cfg = self.app.cfg.setdefault("excel", {})
        excel_cfg["sheet_head"] = self._sheet_head_entry.get().strip()
        excel_cfg["sheet_data"] = self._sheet_data_entry.get().strip()
        try:
            excel_cfg["header_row"] = int(self._header_row_entry.get())
            excel_cfg["start_row"]  = int(self._start_row_entry.get())
        except ValueError:
            pass
        excel_cfg.setdefault("detection", {})
        excel_cfg["detection"]["col"]     = self._det_col_entry.get().strip().upper()
        excel_cfg["detection"]["keyword"] = self._det_keyword_entry.get().strip()
        try:
            excel_cfg["detection"]["row"] = int(self._det_row_entry.get())
        except ValueError:
            pass

        config.save(self.app.cfg)

        if new_ui_lang != i18n.current():
            i18n.load(new_ui_lang)
            self.app.retranslate()

        self.app.update_status(False)
        self.app._refresh_status()

    def retranslate(self) -> None:
        self._title_label.configure(text=i18n.t("settings.title"))
        self._cap_url_label.configure(text=i18n.t("settings.cap_url_label"))
        self._timeout_label.configure(text=i18n.t("settings.timeout_label"))
        self._test_conn_btn.configure(text=i18n.t("settings.test_conn_btn"))
        self._provider_label.configure(text=i18n.t("settings.provider_label"))
        self._lang_label.configure(text=i18n.t("settings.lang_label"))
        self._ui_lang_label.configure(text=i18n.t("settings.ui_lang_label"))
        self._excel_label.configure(text=i18n.t("settings.excel_section"))
        self._save_settings_btn.configure(text=i18n.t("settings.save_btn"))
