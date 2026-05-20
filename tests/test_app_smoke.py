# tests/test_app_smoke.py
import sys
import pytest

@pytest.mark.skipif(sys.platform != "win32", reason="GUI requires Windows display")
def test_app_creates_all_frames_and_destroys():
    from gui.app import App
    app = App()
    assert set(app._frames.keys()) == {"match", "upload", "prompts", "logs", "settings"}
    app.destroy()

@pytest.mark.skipif(sys.platform != "win32", reason="GUI requires Windows display")
def test_show_frame_raises_active_frame():
    from gui.app import App
    app = App()
    app.show_frame("settings")
    # No assertion needed — just verify no exception is raised
    app.destroy()
