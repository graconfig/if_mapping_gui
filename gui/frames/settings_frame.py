# gui/frames/settings_frame.py
import threading
import customtkinter as ctk
import config
from gui.frames import BaseFrame


class SettingsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._build()

    def _build(self):
        pad = {"padx": 20, "pady": 8}
        ctk.CTkLabel(self, text="设置", font=("", 16, "bold")).pack(anchor="w", padx=20, pady=(20, 4))

        # CAP URL row
        url_frame = ctk.CTkFrame(self, fg_color="transparent")
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

        # Provider + Language row
        opts_frame = ctk.CTkFrame(self, fg_color="transparent")
        opts_frame.pack(fill="x", **pad)
        ctk.CTkLabel(opts_frame, text="默认 Provider", font=("", 11)).grid(row=0, column=0, sticky="w")
        self._provider_var = ctk.StringVar(value=self.app.cfg.get("provider", "claude"))
        ctk.CTkOptionMenu(opts_frame, variable=self._provider_var,
                          values=["claude", "openai", "gemini"], width=140).grid(row=1, column=0, padx=(0, 20))
        ctk.CTkLabel(opts_frame, text="默认语言", font=("", 11)).grid(row=0, column=1, sticky="w")
        self._lang_var = ctk.StringVar(value=self.app.cfg.get("language", "ja"))
        ctk.CTkOptionMenu(opts_frame, variable=self._lang_var,
                          values=["ja", "en", "zh"], width=140).grid(row=1, column=1)

        # Save button
        ctk.CTkButton(self, text="💾 保存设置", width=120, command=self._save).pack(anchor="e", padx=20, pady=12)

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
        self.app.cfg["provider"] = self._provider_var.get()
        self.app.cfg["language"] = self._lang_var.get()
        config.save(self.app.cfg)
        self.app.update_status(False)
        self.app._refresh_status()
