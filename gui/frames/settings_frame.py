# gui/frames/settings_frame.py
import threading
import customtkinter as ctk
import config
from config import EXCEL_DEFAULTS
from gui.frames import BaseFrame

_FIELD_LABELS: dict[str, str] = {
    "module": "module", "if_name": "if_name", "if_desc": "if_desc",
    "field_name": "field_name", "is_append": "is_append", "key_flag": "key_flag",
    "obligatory": "obligatory", "data_type": "data_type", "table_id": "table_id",
    "field_id": "field_id", "length_total": "length_total", "length_dec": "length_dec",
    "field_text": "field_text", "sample_value": "sample_value", "remark": "remark",
    "verify": "verify",
    "notes": "notes", "match_score": "match_score", "match_source": "match_source",
}


def _build_col_grid(parent: ctk.CTkFrame, label: str, fields: dict[str, str],
                    entries: dict[str, ctk.CTkEntry]) -> None:
    """Build a labelled group of field→column letter entries in a 3-up grid."""
    ctk.CTkLabel(parent, text=label, font=("", 11, "bold"), text_color="#94a3b8").pack(
        anchor="w", padx=4, pady=(8, 2)
    )
    grid = ctk.CTkFrame(parent, fg_color="transparent")
    grid.pack(fill="x", padx=4)

    items = list(fields.items())
    for idx, (key, default_val) in enumerate(items):
        col = idx % 3
        row = idx // 3
        cell = ctk.CTkFrame(grid, fg_color="transparent")
        cell.grid(row=row, column=col, padx=(0, 12), pady=2, sticky="w")
        ctk.CTkLabel(cell, text=_FIELD_LABELS.get(key, key), font=("", 10),
                     text_color="#64748b", width=90, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(cell, width=52)
        entry.insert(0, default_val)
        entry.pack(side="left")
        entries[key] = entry


class SettingsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._excel_entries: dict[str, dict[str, ctk.CTkEntry]] = {
            "detection": {},
            "global": {},
            "normal_header": {}, "normal_row": {}, "normal_output": {},
            "sap_header": {},    "sap_row": {},    "sap_output": {},
        }
        self._build()

    def _build(self):
        outer = ctk.CTkScrollableFrame(self)
        outer.pack(fill="both", expand=True)
        pad = {"padx": 20, "pady": 8}

        ctk.CTkLabel(outer, text="设置", font=("", 16, "bold")).pack(anchor="w", padx=20, pady=(20, 4))

        # ── CAP URL ──────────────────────────────────────────────────────────
        url_frame = ctk.CTkFrame(outer, fg_color="transparent")
        url_frame.pack(fill="x", **pad)
        ctk.CTkLabel(url_frame, text="CAP 服务地址", font=("", 11)).pack(anchor="w")
        row = ctk.CTkFrame(url_frame, fg_color="transparent")
        row.pack(fill="x")
        self._url_entry = ctk.CTkEntry(row, width=280)
        self._url_entry.insert(0, self.app.cfg.get("server_url", "http://localhost:4004"))
        self._url_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(row, text="🔌 测试连接", width=110, command=self._test_conn).pack(side="left", padx=(0, 8))
        self._conn_status = ctk.CTkLabel(row, text="", font=("", 11))
        self._conn_status.pack(side="left")

        # ── Provider + Language ───────────────────────────────────────────────
        opts_frame = ctk.CTkFrame(outer, fg_color="transparent")
        opts_frame.pack(fill="x", **pad)
        ctk.CTkLabel(opts_frame, text="默认 Provider", font=("", 11)).grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts_frame, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=140).grid(row=1, column=0, padx=(0, 20))
        ctk.CTkLabel(opts_frame, text="默认语言", font=("", 11)).grid(row=0, column=1, sticky="w")
        self._lang_var = ctk.StringVar(value=self.app.cfg.get("language", "ja"))
        ctk.CTkOptionMenu(opts_frame, variable=self._lang_var,
                          values=["ja", "en", "zh"], width=140).grid(row=1, column=1)

        # ── Excel 列配置 ──────────────────────────────────────────────────────
        ctk.CTkLabel(outer, text="Excel 列配置", font=("", 13, "bold")).pack(
            anchor="w", padx=20, pady=(16, 4)
        )
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

        global_frame = ctk.CTkFrame(excel_section, fg_color="transparent")
        global_frame.pack(fill="x", padx=12, pady=(10, 4))
        self._sheet_head_entry  = _labeled_entry(global_frame, "Sheet（抬头）", excel_cfg["sheet_head"])
        self._sheet_data_entry  = _labeled_entry(global_frame, "Sheet（数据）", excel_cfg["sheet_data"])
        self._header_row_entry  = _labeled_entry(global_frame, "抬头行", excel_cfg["header_row"], width=60)
        self._start_row_entry   = _labeled_entry(global_frame, "起始行", excel_cfg["start_row"],  width=60)

        det_frame = ctk.CTkFrame(excel_section, fg_color="transparent")
        det_frame.pack(fill="x", padx=12, pady=(4, 8))
        det = excel_cfg["detection"]
        self._det_col_entry     = _labeled_entry(det_frame, "检测列", det["col"],     width=52)
        self._det_row_entry     = _labeled_entry(det_frame, "检测行", det["row"],     width=60)
        self._det_keyword_entry = _labeled_entry(det_frame, "关键字", det["keyword"], width=80)

        tab = ctk.CTkTabview(excel_section)
        tab.pack(fill="x", padx=12, pady=(0, 12))
        tab.add("普通方向")
        tab.add("SAP方向")

        for direction, tab_name, hkey, rkey, okey in [
            ("normal", "普通方向", "normal_header", "normal_row", "normal_output"),
            ("sap",    "SAP方向",  "sap_header",    "sap_row",    "sap_output"),
        ]:
            dir_cfg = excel_cfg["directions"][direction]
            scroll = ctk.CTkScrollableFrame(tab.tab(tab_name), height=340)
            scroll.pack(fill="both", expand=True)
            _build_col_grid(scroll, "抬头列",    dir_cfg["input_header_cols"], self._excel_entries[hkey])
            _build_col_grid(scroll, "输入数据列", dir_cfg["input_row_cols"],    self._excel_entries[rkey])
            _build_col_grid(scroll, "输出列",    dir_cfg["output_cols"],       self._excel_entries[okey])

        ctk.CTkButton(outer, text="💾 保存设置", width=120, command=self._save).pack(
            anchor="e", padx=20, pady=12
        )

    def _test_conn(self):
        url = self._url_entry.get().strip()
        self._conn_status.configure(text="测试中…", text_color="gray")

        def _check():
            from api.cap_client import CapClient
            ok = CapClient(url).ping()
            self.after(0, lambda: self._conn_status.configure(
                text="✓ 已连接" if ok else "✗ 无法连接",
                text_color="#22c55e" if ok else "#ef4444",
            ))

        threading.Thread(target=_check, daemon=True).start()

    def _save(self):
        self.app.cfg["server_url"] = self._url_entry.get().strip()
        self.app.cfg["provider"]   = self._provider_var.get()
        self.app.cfg["language"]   = self._lang_var.get()

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

        for direction, hkey, rkey, okey in [
            ("normal", "normal_header", "normal_row", "normal_output"),
            ("sap",    "sap_header",    "sap_row",    "sap_output"),
        ]:
            dir_cfg = excel_cfg.setdefault("directions", {}).setdefault(direction, {})
            dir_cfg["input_header_cols"] = {
                k: e.get().strip().upper() for k, e in self._excel_entries[hkey].items()
            }
            dir_cfg["input_row_cols"] = {
                k: e.get().strip().upper() for k, e in self._excel_entries[rkey].items()
            }
            dir_cfg["output_cols"] = {
                k: e.get().strip().upper() for k, e in self._excel_entries[okey].items()
            }

        config.save(self.app.cfg)
        self.app.update_status(False)
        self.app._refresh_status()
