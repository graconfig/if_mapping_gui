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

def load() -> dict:
    if not CONFIG_PATH.exists():
        return dict(DEFAULTS)
    with CONFIG_PATH.open(encoding="utf-8") as f:
        data = json.load(f)
    return {**DEFAULTS, **data}

def save(cfg: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
