import json

import config


def test_load_missing_file_returns_defaults(tmp_path):
    assert config.load_config(tmp_path / "none.json") == config.DEFAULTS


def test_load_corrupt_file_returns_defaults(tmp_path):
    path = tmp_path / "c.json"
    path.write_text("{không phải json", encoding="utf-8")
    assert config.load_config(path) == config.DEFAULTS


def test_save_then_load_round_trip_and_ignores_unknown_keys(tmp_path):
    path = tmp_path / "c.json"
    cfg = dict(config.DEFAULTS, start_cell="B18", file_path="E:/x.xlsx", extra="bỏ")
    config.save_config(cfg, path)
    assert "extra" not in json.loads(path.read_text(encoding="utf-8"))

    loaded = config.load_config(path)
    assert loaded["start_cell"] == "B18" and loaded["file_path"] == "E:/x.xlsx"


def test_find_chrome_prefers_custom_path(tmp_path):
    fake = tmp_path / "chrome.exe"
    fake.write_text("x")
    assert config.find_chrome(str(fake)) == str(fake)


def test_find_chrome_falls_back_to_candidates(tmp_path, monkeypatch):
    fake = tmp_path / "chrome.exe"
    fake.write_text("x")
    monkeypatch.setattr(config, "chrome_candidates", lambda: iter(["Z:/khong/co.exe", str(fake)]))
    assert config.find_chrome("") == str(fake)
    monkeypatch.setattr(config, "chrome_candidates", lambda: iter([]))
    assert config.find_chrome("") is None
