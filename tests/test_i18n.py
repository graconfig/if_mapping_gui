import i18n


def test_zh_translation():
    i18n.load("zh")
    assert i18n.t("app.title") == "IF Mapping"
    assert i18n.t("nav.match") == "▶  字段匹配"


def test_ja_translation():
    i18n.load("ja")
    assert i18n.t("app.title") == "IF Mapping"
    assert i18n.t("nav.match") == "▶  フィールドマッチング"


def test_interpolation_zh():
    i18n.load("zh")
    result = i18n.t("match.parse_log", name="test.xlsx", count="10", direction="normal")
    assert result == "解析 test.xlsx — 10 条字段 [normal]"


def test_interpolation_ja():
    i18n.load("ja")
    result = i18n.t("match.parse_log", name="test.xlsx", count="10", direction="normal")
    assert result == "test.xlsx を解析 — 10 件 [normal]"


def test_missing_key_returns_key():
    i18n.load("zh")
    assert i18n.t("nonexistent.key") == "nonexistent.key"


def test_current_returns_active_lang():
    i18n.load("ja")
    assert i18n.current() == "ja"
    i18n.load("zh")
    assert i18n.current() == "zh"


def test_all_keys_present_in_both_langs():
    i18n.load("zh")
    zh_keys = set(i18n._cache["zh"].keys())
    i18n.load("ja")
    ja_keys = set(i18n._cache["ja"].keys())
    missing_in_ja = zh_keys - ja_keys
    assert not missing_in_ja, f"Keys missing in ja.json: {missing_in_ja}"
