# tests/test_match_worker.py
import queue
import threading
from pathlib import Path
from unittest.mock import patch, MagicMock
import openpyxl
import pytest

from gui.frames.match_frame import match_worker

def _make_xlsx(tmp_path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["sourceField", "sourceDesc"])
    for r in rows:
        ws.append(r)
    p = tmp_path / "input.xlsx"
    wb.save(p)
    return p

def _drain(q):
    msgs = []
    while not q.empty():
        msgs.append(q.get_nowait())
    return msgs

def test_worker_emits_done_with_results(tmp_path):
    xlsx = _make_xlsx(tmp_path, [["MATNR", "Material"]])
    q = queue.Queue()
    stop = threading.Event()
    mock_results = [{"rowIndex": 2, "targetField": "Material", "matchType": "exact",
                     "targetEntity": "CDS_View", "confidence": 1.0, "aiReason": ""}]
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        MockClient.return_value.match.return_value = mock_results
        match_worker([str(xlsx)], "claude", "ja", "http://localhost:4004", q, stop)
    msgs = _drain(q)
    types = [m["type"] for m in msgs]
    assert "done" in types
    done_msg = next(m for m in msgs if m["type"] == "done")
    assert done_msg["results"] == mock_results

def test_worker_emits_error_when_cap_unreachable(tmp_path):
    xlsx = _make_xlsx(tmp_path, [["MATNR", "Material"]])
    q = queue.Queue()
    stop = threading.Event()
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = False
        match_worker([str(xlsx)], "claude", "ja", "http://localhost:4004", q, stop)
    msgs = _drain(q)
    types = [m["type"] for m in msgs]
    assert "error" in types

def test_worker_skips_bad_excel_and_continues(tmp_path):
    bad = tmp_path / "bad.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["wrong_col"])
    wb.save(bad)
    good = _make_xlsx(tmp_path, [["MATNR", "Material"]])
    q = queue.Queue()
    stop = threading.Event()
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        MockClient.return_value.match.return_value = []
        match_worker([str(bad), str(good)], "claude", "ja", "http://localhost:4004", q, stop)
    msgs = _drain(q)
    log_texts = [m["text"] for m in msgs if m["type"] == "log"]
    assert any("sourceField" in t or "ERROR" in t for t in log_texts)
    assert any(m["type"] == "done" for m in msgs)

def test_worker_respects_stop_event(tmp_path):
    for i in range(3):
        (tmp_path / f"f{i}").mkdir(exist_ok=True)
        wb = openpyxl.Workbook()
        wb.active.append(["sourceField", "sourceDesc"])
        wb.active.append(["MATNR", "Material"])
        wb.save(tmp_path / f"f{i}" / "input.xlsx")
    files = [str(tmp_path / f"f{i}" / "input.xlsx") for i in range(3)]
    q = queue.Queue()
    stop = threading.Event()
    stop.set()  # stopped before start
    with patch("gui.frames.match_frame.CapClient") as MockClient:
        MockClient.return_value.ping.return_value = True
        match_worker(files, "claude", "ja", "http://localhost:4004", q, stop)
    msgs = _drain(q)
    log_texts = [m["text"] for m in msgs if m["type"] == "log"]
    assert any("停止" in t for t in log_texts)
