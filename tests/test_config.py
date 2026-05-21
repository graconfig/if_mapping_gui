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

def test_load_returns_excel_defaults_when_no_file():
    cfg = config.load()
    assert "excel" in cfg
    assert cfg["excel"]["sheet_head"] == "対象IF"
    assert cfg["excel"]["sheet_data"] == "IFマッピング定義"
    assert cfg["excel"]["header_row"] == 6
    assert cfg["excel"]["start_row"] == 5
    assert cfg["excel"]["detection"]["col"] == "F"
    assert cfg["excel"]["detection"]["row"] == 6
    assert cfg["excel"]["detection"]["keyword"] == "SAP"
    assert cfg["excel"]["directions"]["normal"]["input_row_cols"]["field_name"] == "C"
    assert cfg["excel"]["directions"]["sap"]["input_row_cols"]["field_name"] == "S"

def test_load_deep_merges_partial_excel_config():
    config.CONFIG_PATH.write_text(json.dumps({
        "excel": {"sheet_head": "MySheet", "directions": {"normal": {"input_row_cols": {"field_name": "B"}}}}
    }))
    cfg = config.load()
    assert cfg["excel"]["sheet_head"] == "MySheet"           # overridden
    assert cfg["excel"]["sheet_data"] == "IFマッピング定義"   # default preserved
    assert cfg["excel"]["directions"]["normal"]["input_row_cols"]["field_name"] == "B"  # overridden
    assert cfg["excel"]["directions"]["normal"]["input_row_cols"]["field_text"] == "L"  # default preserved
    assert "sap" in cfg["excel"]["directions"]                # other direction preserved
