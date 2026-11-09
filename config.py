import copy
import json
import os
import shutil
import sys
from pathlib import Path

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"

DEFAULTS = {
    "file_path": "",
    "sheet_name": "Cân Chỉnh Hàng Hóa",
    "start_cell": "B4",
    "wait_seconds": "5.0",
    "chrome_path": "",
    "debug_port": 9999,
    "user_data_dir": r"C:\ChromeDebug",
    "note_text": "Cân hàng lệch",
    "notify_sound": True,
    "notify_flash": True,
    "status_interval": 5,
    "log_retention_days": 7,
    "branch_overrides": {},
}


def load_config(path=None):
    path = path or CONFIG_PATH
    config = copy.deepcopy(DEFAULTS)
    try:
        with open(path, "r", encoding="utf-8") as f:
            saved = json.load(f)
    except (OSError, ValueError):
        return config

    if isinstance(saved, dict):
        # Chỉ nhận các khóa đã biết để file cấu hình hỏng/thừa không làm hỏng Tool
        config.update({key: saved[key] for key in DEFAULTS if key in saved})
    return config


def save_config(config, path=None):
    path = path or CONFIG_PATH
    data = {key: config.get(key, default) for key, default in DEFAULTS.items()}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _as_int(value):
    if isinstance(value, bool):
        return None
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _as_float(value):
    if isinstance(value, bool):
        return None
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def validate_config(cfg):
    """Kiểm tra các thiết lập trong cửa sổ Cài đặt. Trả về {khóa: lỗi}; rỗng là hợp lệ."""
    errors = {}

    chrome_path = str(cfg.get("chrome_path") or "").strip()
    if chrome_path and not os.path.isfile(chrome_path):
        errors["chrome_path"] = "Không tìm thấy chrome.exe ở đường dẫn này."

    port = _as_int(cfg.get("debug_port"))
    if port is None or not 1024 <= port <= 65535:
        errors["debug_port"] = "Cổng debug phải là số nguyên từ 1024 đến 65535."

    if not str(cfg.get("user_data_dir") or "").strip():
        errors["user_data_dir"] = "Thư mục profile không được để trống."

    wait = _as_float(cfg.get("wait_seconds"))
    if wait is None or not 1 <= wait <= 60:
        errors["wait_seconds"] = "Thời gian chờ phải từ 1 đến 60 giây."

    if len(str(cfg.get("note_text") or "")) > 2000:
        errors["note_text"] = "Ghi chú tối đa 2000 ký tự."

    interval = _as_int(cfg.get("status_interval"))
    if interval is None or not 3 <= interval <= 60:
        errors["status_interval"] = "Tần suất cập nhật phải từ 3 đến 60 giây."

    retention = _as_int(cfg.get("log_retention_days"))
    if retention is None or not 1 <= retention <= 365:
        errors["log_retention_days"] = "Số ngày giữ nhật ký phải từ 1 đến 365."

    return errors


def chrome_candidates():
    roots = [
        os.environ.get("PROGRAMFILES"),
        os.environ.get("PROGRAMFILES(X86)"),
        os.environ.get("LOCALAPPDATA"),
    ]
    for root in roots:
        if root:
            yield os.path.join(root, "Google", "Chrome", "Application", "chrome.exe")

    found = shutil.which("chrome")
    if found:
        yield found


def find_chrome(custom_path=""):
    """Ưu tiên đường dẫn người dùng đã chọn, sau đó dò các vị trí cài đặt phổ biến."""
    if custom_path and os.path.isfile(custom_path):
        return custom_path

    for candidate in chrome_candidates():
        if os.path.isfile(candidate):
            return candidate
    return None
