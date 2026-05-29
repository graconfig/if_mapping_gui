# gui/app.py
import customtkinter as ctk
import config
import i18n
from gui.frames.match_frame import MatchFrame
from gui.frames.upload_frame import UploadFrame
from gui.frames.prompts_frame import PromptsFrame
from gui.frames.logs_frame import LogsFrame
from gui.frames.settings_frame import SettingsFrame
from api.cap_client import CapClient

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_NAV_KEYS = [
    ("match",   "nav.match"),
    ("upload",  "nav.upload"),
    ("prompts", "nav.prompts"),
    ("logs",    "nav.logs"),
]


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cfg = config.load()
        i18n.load(self.cfg.get("ui_language", "zh"))

        self.title(i18n.t("app.title"))
        self.geometry("920x640")

        self._build_layout()
        self.show_frame("match")
        self.after(500, self._refresh_status)

    def _build_layout(self):
        # Status bar must be packed FIRST so tkinter reserves bottom space before
        # the side="left" widgets consume all remaining area.
        self._statusbar = ctk.CTkFrame(self, height=28, corner_radius=0)
        self._statusbar.pack(side="bottom", fill="x")
        self._statusbar.pack_propagate(False)
        self._conn_label = ctk.CTkLabel(self._statusbar, text=i18n.t("status.checking"), font=("", 11))
        self._conn_label.pack(side="left", padx=12)
        ctk.CTkLabel(self._statusbar, text="v0.1.0", font=("", 10)).pack(side="right", padx=12)

        # Sidebar
        self._sidebar = ctk.CTkFrame(self, width=164, corner_radius=0)
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)

        self._app_title_label = ctk.CTkLabel(
            self._sidebar, text=i18n.t("app.title"), font=("", 13, "bold")
        )
        self._app_title_label.pack(pady=(18, 14), padx=10)

        self._nav_btns: dict[str, ctk.CTkButton] = {}
        self._nav_keys = _NAV_KEYS
        for key, label_key in _NAV_KEYS:
            btn = ctk.CTkButton(
                self._sidebar, text=i18n.t(label_key), anchor="w", height=34,
                command=lambda k=key: self.show_frame(k),
            )
            btn.pack(fill="x", padx=8, pady=2)
            self._nav_btns[key] = btn

        settings_btn = ctk.CTkButton(
            self._sidebar, text=i18n.t("nav.settings"), anchor="w", height=34,
            command=lambda: self.show_frame("settings"),
        )
        settings_btn.pack(fill="x", padx=8, pady=2, side="bottom")
        self._nav_btns["settings"] = settings_btn

        # Main area
        self._main = ctk.CTkFrame(self, corner_radius=0)
        self._main.pack(side="left", fill="both", expand=True)

        self._frames: dict[str, ctk.CTkFrame] = {
            "match":    MatchFrame(self._main, self),
            "upload":   UploadFrame(self._main, self),
            "prompts":  PromptsFrame(self._main, self),
            "logs":     LogsFrame(self._main, self),
            "settings": SettingsFrame(self._main, self),
        }
        for frame in self._frames.values():
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)

    def show_frame(self, name: str) -> None:
        self._frames[name].tkraise()
        for k, btn in self._nav_btns.items():
            btn.configure(fg_color=("#0f3460", "#0f3460") if k == name else "transparent")

    def retranslate(self) -> None:
        self.title(i18n.t("app.title"))
        self._app_title_label.configure(text=i18n.t("app.title"))
        for key, label_key in self._nav_keys:
            self._nav_btns[key].configure(text=i18n.t(label_key))
        self._nav_btns["settings"].configure(text=i18n.t("nav.settings"))
        for frame in self._frames.values():
            frame.retranslate()

    def update_status(self, connected: bool) -> None:
        url = self.cfg.get("server_url", "")
        if connected:
            self._conn_label.configure(text=f"●  {url}", text_color="#22c55e")
        else:
            self._conn_label.configure(text=i18n.t("status.disconnected"), text_color="#ef4444")

    def get_client(self) -> CapClient:
        return CapClient(
            self.cfg.get("server_url", "http://localhost:4004"),
            xsuaa=self.cfg.get("xsuaa"),
        )

    def _refresh_status(self) -> None:
        import threading
        def _check():
            ok = self.get_client().ping()
            self.after(0, lambda: self.update_status(ok))
        threading.Thread(target=_check, daemon=True).start()
        self.after(30_000, self._refresh_status)
