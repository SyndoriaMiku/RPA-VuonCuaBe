import json

import config


def test_new_settings_have_defaults():
    assert config.DEFAULTS["notify_sound"] is True
    assert config.DEFAULTS["notify_flash"] is True
    assert config.DEFAULTS["status_interval"] == 5
    assert config.DEFAULTS["log_retention_days"] == 7
    assert config.DEFAULTS["branch_overrides"] == {}


def test_load_old_config_fills_new_settings(tmp_path):
    path = tmp_path / "c.json"
    path.write_text(json.dumps({"file_path": "E:/x.xlsx"}), encoding="utf-8")
    loaded = config.load_config(path)
    assert loaded["notify_sound"] is True
    assert loaded["status_interval"] == 5
    assert loaded["branch_overrides"] == {}


def test_loaded_branch_overrides_do_not_share_the_default_dict(tmp_path):
    loaded = config.load_config(tmp_path / "none.json")
    loaded["branch_overrides"]["x"] = "1"
    assert config.DEFAULTS["branch_overrides"] == {}


def test_save_round_trips_branch_overrides(tmp_path):
    path = tmp_path / "c.json"
    cfg = dict(config.DEFAULTS, branch_overrides={"cs08 22 vu trong phung": "152162"})
    config.save_config(cfg, path)
    assert config.load_config(path)["branch_overrides"] == {"cs08 22 vu trong phung": "152162"}


def valid_config(tmp_path=None):
    return dict(
        config.DEFAULTS,
        chrome_path="",
        debug_port=9999,
        user_data_dir=r"C:\ChromeDebug",
        wait_seconds="5.0",
        note_text="Cân hàng lệch",
        status_interval=5,
    )


def test_validate_config_accepts_valid_settings(tmp_path):
    assert config.validate_config(valid_config()) == {}


def test_validate_config_accepts_numeric_strings():
    cfg = valid_config()
    cfg["debug_port"] = "9222"
    cfg["status_interval"] = "10"
    assert config.validate_config(cfg) == {}


def test_validate_config_rejects_missing_chrome_file(tmp_path):
    cfg = valid_config()
    cfg["chrome_path"] = str(tmp_path / "khong-co.exe")
    assert "chrome_path" in config.validate_config(cfg)


def test_validate_config_accepts_existing_chrome_file(tmp_path):
    fake = tmp_path / "chrome.exe"
    fake.write_text("x")
    cfg = valid_config()
    cfg["chrome_path"] = str(fake)
    assert config.validate_config(cfg) == {}


def test_validate_config_checks_port():
    for value in ("abc", "80", "70000", None):
        cfg = valid_config()
        cfg["debug_port"] = value
        assert "debug_port" in config.validate_config(cfg)


def test_validate_config_checks_user_data_dir():
    cfg = valid_config()
    cfg["user_data_dir"] = "  "
    assert "user_data_dir" in config.validate_config(cfg)


def test_validate_config_checks_wait_seconds():
    for value in ("0", "61", "lâu"):
        cfg = valid_config()
        cfg["wait_seconds"] = value
        assert "wait_seconds" in config.validate_config(cfg)


def test_validate_config_checks_note_length():
    cfg = valid_config()
    cfg["note_text"] = "a" * 2001
    assert "note_text" in config.validate_config(cfg)
    cfg["note_text"] = "a" * 2000
    assert "note_text" not in config.validate_config(cfg)


def test_validate_config_checks_status_interval():
    for value in ("2", "61", "nhanh"):
        cfg = valid_config()
        cfg["status_interval"] = value
        assert "status_interval" in config.validate_config(cfg)


def test_validate_config_checks_log_retention_days():
    for value in ("0", "366", "lâu", None):
        cfg = valid_config()
        cfg["log_retention_days"] = value
        assert "log_retention_days" in config.validate_config(cfg)


def test_validate_config_accepts_log_retention_as_number_or_string():
    cfg = valid_config()
    cfg["log_retention_days"] = "14"
    assert "log_retention_days" not in config.validate_config(cfg)
    cfg["log_retention_days"] = 1
    assert "log_retention_days" not in config.validate_config(cfg)
