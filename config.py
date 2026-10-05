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
}


def load_config(path=CONFIG_PATH):
    config = dict(DEFAULTS)
    try:
        with open(path, "r", encoding="utf-8") as f:
            saved = json.load(f)
    except (OSError, ValueError):
        return config

    if isinstance(saved, dict):
        # Chỉ nhận các khóa đã biết để file cấu hình hỏng/thừa không làm hỏng Tool
        config.update({key: saved[key] for key in DEFAULTS if key in saved})
    return config


def save_config(config, path=CONFIG_PATH):
    data = {key: config.get(key, default) for key, default in DEFAULTS.items()}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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
