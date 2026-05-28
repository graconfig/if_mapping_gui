# gui/frames/prompts_frame.py
import threading
import customtkinter as ctk
import i18n
from gui.frames import BaseFrame


class PromptsFrame(BaseFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._prompts: list[dict] = []
        self._selected_idx: int | None = None
        self._lang_filter = "ja"
        self._build()

    def _build(self):
        self._title_label = ctk.CTkLabel(self, text=i18n.t("prompts.title"), font=("", 16, "bold"))
        self._title_label.pack(anchor="w", padx=20, pady=(20, 8))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        main.columnconfigure(1, weight=1)
        main.rowconfigure(1, weight=1)

        # Language filter chips
        filter_row = ctk.CTkFrame(main, fg_color="transparent")
        filter_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self._lang_btns: dict[str, ctk.CTkButton] = {}
        for lang in ("ja", "en", "zh"):
            btn = ctk.CTkButton(filter_row, text=lang.upper(), width=44, height=26,
                                command=lambda l=lang: self._set_lang(l))
            btn.pack(side="left", padx=2)
            self._lang_btns[lang] = btn
        self._set_lang("ja")

        # Prompt list (left)
        self._list_frame = ctk.CTkScrollableFrame(main, width=210)
        self._list_frame.grid(row=1, column=0, sticky="ns", padx=(0, 10))

        # Edit panel (right)
        right = ctk.CTkFrame(main, fg_color="transparent")
        right.grid(row=0, column=1, rowspan=2, sticky="nsew")
        right.rowconfigure(1, weight=1)

        self._edit_title = ctk.CTkLabel(right, text="", font=("", 12, "bold"))
        self._edit_title.grid(row=0, column=0, sticky="w", pady=(0, 6))
        self._editor = ctk.CTkTextbox(right, font=("Consolas", 11))
        self._editor.grid(row=1, column=0, sticky="nsew")
        right.columnconfigure(0, weight=1)

        btn_row = ctk.CTkFrame(right, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="e", pady=(8, 0))
        self._cancel_btn = ctk.CTkButton(btn_row, text=i18n.t("prompts.cancel_btn"), width=70,
                                         fg_color="#1e293b", command=self._cancel)
        self._cancel_btn.pack(side="left", padx=4)
        self._save_btn = ctk.CTkButton(btn_row, text=i18n.t("prompts.save_btn"), width=80, command=self._save)
        self._save_btn.pack(side="left", padx=4)
        self._reload_btn = ctk.CTkButton(btn_row, text=i18n.t("prompts.reload_btn"), width=80,
                                         fg_color="#1e293b", command=self._reload_server)
        self._reload_btn.pack(side="left", padx=4)

        # Load prompts on display
        self.bind("<Visibility>", lambda e: self._load_prompts())
        self._load_prompts()

    def _set_lang(self, lang: str):
        self._lang_filter = lang
        for l, btn in self._lang_btns.items():
            btn.configure(fg_color="#0f3460" if l == lang else "#1e293b")
        self._load_prompts()

    def _load_prompts(self):
        def _fetch():
            try:
                prompts = self.app.get_client().get_prompts(language=self._lang_filter)
                self.after(0, lambda: self._populate_list(prompts))
            except Exception:
                pass
        threading.Thread(target=_fetch, daemon=True).start()

    def _populate_list(self, prompts: list[dict]):
        self._prompts = prompts
        for w in self._list_frame.winfo_children():
            w.destroy()
        for i, p in enumerate(prompts):
            label = f"{p.get('step', '')} / {p.get('promptType', '')}"
            btn = ctk.CTkButton(self._list_frame, text=label, anchor="w", height=30,
                                font=("", 11), command=lambda idx=i: self._select(idx))
            btn.pack(fill="x", pady=1)
        if prompts and self._selected_idx is None:
            self._select(0)

    def _select(self, idx: int):
        self._selected_idx = idx
        p = self._prompts[idx]
        self._edit_title.configure(
            text=f"{p.get('step')} / {p.get('promptType')} / {p.get('language', '').upper()}"
        )
        self._editor.delete("1.0", "end")
        self._editor.insert("1.0", p.get("content", ""))

    def _cancel(self):
        if self._selected_idx is not None:
            self._select(self._selected_idx)

    def _save(self):
        if self._selected_idx is None:
            return
        p = self._prompts[self._selected_idx]
        content = self._editor.get("1.0", "end-1c")
        def _patch():
            try:
                self.app.get_client().patch_prompt(p["ID"], content)
                self._prompts[self._selected_idx]["content"] = content
            except Exception as e:
                print(f"patch_prompt error: {e}")
        threading.Thread(target=_patch, daemon=True).start()

    def _reload_server(self):
        def _reload():
            try:
                self.app.get_client().reload_prompts()
            except Exception as e:
                print(f"reload_prompts error: {e}")
        threading.Thread(target=_reload, daemon=True).start()

    def retranslate(self) -> None:
        self._title_label.configure(text=i18n.t("prompts.title"))
        self._cancel_btn.configure(text=i18n.t("prompts.cancel_btn"))
        self._save_btn.configure(text=i18n.t("prompts.save_btn"))
        self._reload_btn.configure(text=i18n.t("prompts.reload_btn"))
