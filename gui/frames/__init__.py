import customtkinter as ctk

class BaseFrame(ctk.CTkFrame):
    """Base class for all content frames."""
    def __init__(self, master, app, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app

    def retranslate(self) -> None:
        """Override in subclasses to refresh all static widget text."""
        pass
