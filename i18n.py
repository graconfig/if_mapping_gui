import json
from pathlib import Path

_FALLBACK = "zh"
_cache: dict[str, dict[str, str]] = {}
_lang = _FALLBACK


def load(lang: str) -> None:
    global _lang
    _lang = lang
    if lang not in _cache:
        path = Path(__file__).parent / "i18n" / f"{lang}.json"
        _cache[lang] = json.loads(path.read_text(encoding="utf-8"))


def t(key: str, **kwargs: str) -> str:
    strings = _cache.get(_lang) or _cache.get(_FALLBACK) or {}
    text = strings.get(key, key)
    return text.format(**kwargs) if kwargs else text


def current() -> str:
    return _lang
