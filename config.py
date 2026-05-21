# config.py
import json
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.json"

DEFAULTS: dict = {
    "server_url": "http://localhost:4004",
    "provider": "claude",
    "language": "ja",
    "last_input_dir": "",
}

EXCEL_DEFAULTS: dict = {
    "sheet_head": "対象IF",
    "sheet_data": "IFマッピング定義",
    "header_row": 6,
    "start_row": 5,
    "detection": {"col": "F", "row": 6, "keyword": "SAP"},
    "directions": {
        "normal": {
            "input_header_cols": {"module": "D", "if_name": "C", "if_desc": "E"},
            "input_row_cols": {
                "field_name": "C", "is_append": "D", "key_flag": "E", "obligatory": "F",
                "data_type": "I", "table_id": "G", "field_id": "H",
                "length_total": "J", "length_dec": "K",
                "field_text": "L", "sample_value": "N", "remark": "M", "verify": "AF",
            },
            "output_cols": {
                "field_name": "S", "is_append": "T", "key_flag": "U", "obligatory": "V",
                "table_id": "W", "field_id": "X", "data_type": "Y",
                "length_total": "Z", "length_dec": "AA",
                "notes": "AB", "sample_value": "AC", "match_score": "AE",
                "match_source": "AF", "verify": "AG",
            },
        },
        "sap": {
            "input_header_cols": {"module": "D", "if_name": "C", "if_desc": "E"},
            "input_row_cols": {
                "field_name": "S", "is_append": "T", "key_flag": "U", "obligatory": "V",
                "table_id": "W", "field_id": "X", "data_type": "Y",
                "length_total": "Z", "length_dec": "AA",
                "field_text": "AB", "sample_value": "AC", "remark": "AE", "verify": "AF",
            },
            "output_cols": {
                "field_name": "C", "is_append": "D", "key_flag": "E", "obligatory": "F",
                "table_id": "G", "field_id": "H", "data_type": "I",
                "length_total": "J", "length_dec": "K",
                "notes": "M", "sample_value": "N", "match_score": "AE",
                "match_source": "AF", "verify": "AG",
            },
        },
    },
}


def _deep_merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def load() -> dict:
    if not CONFIG_PATH.exists():
        return {**DEFAULTS, "excel": EXCEL_DEFAULTS}
    with CONFIG_PATH.open(encoding="utf-8") as f:
        data = json.load(f)
    excel_override = data.pop("excel", {})
    merged = {**DEFAULTS, **data}
    merged["excel"] = _deep_merge(EXCEL_DEFAULTS, excel_override)
    return merged


def save(cfg: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
