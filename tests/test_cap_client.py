# tests/test_cap_client.py
import requests
import pytest
from unittest.mock import patch, MagicMock
from api.cap_client import CapClient, CapConnectionError

BASE = "http://localhost:4004"

def _mock_response(json_data=None, status=200):
    m = MagicMock()
    m.status_code = status
    m.json.return_value = json_data or {}
    m.raise_for_status = MagicMock()
    return m

def test_ping_true_when_service_responds():
    with patch("requests.get", return_value=_mock_response(status=200)):
        assert CapClient(BASE).ping() is True

def test_ping_false_on_connection_error():
    with patch("requests.get", side_effect=requests.ConnectionError()):
        assert CapClient(BASE).ping() is False

def test_ping_false_on_timeout():
    with patch("requests.get", side_effect=requests.Timeout()):
        assert CapClient(BASE).ping() is False

def test_match_posts_correct_payload_and_returns_results():
    results = [{"sourceField": "MATNR", "matchType": "exact", "targetField": "Material",
                "targetEntity": "C_PurchaseOrderItemTP", "confidence": 1.0, "aiReason": ""}]
    with patch("requests.post", return_value=_mock_response({"value": results})) as mock_post:
        client = CapClient(BASE)
        out = client.match(
            [{"sourceField": "MATNR", "sourceDesc": "Material No.", "sourceTable": "", "rowIndex": 2}],
            provider="claude", language="ja"
        )
    assert out == results
    mock_post.assert_called_once_with(
        f"{BASE}/if-mapping/match",
        json={
            "fields": [{"sourceField": "MATNR", "sourceDesc": "Material No.", "sourceTable": "", "rowIndex": 2}],
            "provider": "claude",
            "language": "ja",
        },
        timeout=30,
    )

def test_match_raises_cap_connection_error_on_http_error():
    mock_resp = _mock_response(status=500)
    mock_resp.raise_for_status.side_effect = requests.HTTPError("500")
    with patch("requests.post", return_value=mock_resp):
        with pytest.raises(CapConnectionError):
            CapClient(BASE).match([], "claude", "ja")

def test_upload_custom_fields_posts_and_returns_counts():
    counts = {"inserted": 10, "updated": 2, "deleted": 0}
    with patch("requests.post", return_value=_mock_response(counts)) as mock_post:
        out = CapClient(BASE).upload_custom_fields([{"sourceField": "X"}], mode="upsert")
    assert out == counts
    mock_post.assert_called_once_with(
        f"{BASE}/if-mapping/uploadCustomFields",
        json={"records": [{"sourceField": "X"}], "mode": "upsert"},
        timeout=30,
    )

def test_get_prompts_returns_list():
    prompts = [{"id": "1", "step": "field_matching", "language": "ja", "promptType": "user", "content": "..."}]
    with patch("requests.get", return_value=_mock_response({"value": prompts})):
        out = CapClient(BASE).get_prompts(language="ja")
    assert out == prompts

def test_patch_prompt_sends_content():
    updated = {"id": "1", "content": "new text"}
    with patch("requests.patch", return_value=_mock_response(updated)) as mock_patch:
        out = CapClient(BASE).patch_prompt("1", "new text")
    assert out == updated
    mock_patch.assert_called_once_with(
        f"{BASE}/if-mapping/PromptTemplates('1')",
        json={"content": "new text"},
        timeout=30,
    )

def test_reload_prompts_posts_to_endpoint():
    with patch("requests.post", return_value=_mock_response()) as mock_post:
        CapClient(BASE).reload_prompts()
    mock_post.assert_called_once_with(f"{BASE}/if-mapping/reloadPrompts", timeout=30)

def test_get_token_logs_returns_list():
    logs = [{"timestamp": "14:32:04", "provider": "claude", "step": "field_matching",
             "inputTokens": 523, "outputTokens": 48}]
    with patch("requests.get", return_value=_mock_response({"value": logs})):
        out = CapClient(BASE).get_token_logs()
    assert out == logs
