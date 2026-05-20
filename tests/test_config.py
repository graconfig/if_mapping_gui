# tests/test_config.py
import json
from pathlib import Path
import pytest
import config

@pytest.fixture(autouse=True)
def isolated_config(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "CONFIG_PATH", tmp_path / "config.json")

def test_load_returns_defaults_when_no_file():
    cfg = config.load()
    assert cfg["server_url"] == "http://localhost:4004"
    assert cfg["provider"] == "claude"
    assert cfg["language"] == "ja"
    assert cfg["last_input_dir"] == ""

def test_load_merges_saved_values_with_defaults():
    config.CONFIG_PATH.write_text(json.dumps({"server_url": "http://myserver:4004"}))
    cfg = config.load()
    assert cfg["server_url"] == "http://myserver:4004"
    assert cfg["provider"] == "claude"  # default still present

def test_save_and_reload_roundtrip():
    cfg = config.load()
    cfg["language"] = "en"
    cfg["last_input_dir"] = "C:/files"
    config.save(cfg)
    reloaded = config.load()
    assert reloaded["language"] == "en"
    assert reloaded["last_input_dir"] == "C:/files"

def test_save_creates_file_if_missing():
    assert not config.CONFIG_PATH.exists()
    config.save(config.load())
    assert config.CONFIG_PATH.exists()
