"""Ghi nhật ký của Tool ra file và tự xóa log cũ.

Mỗi lần mở Tool tạo một file trong thư mục `logs` cạnh AutoApp.exe, để khi cần báo lỗi chỉ
việc gửi kèm file của lần chạy đó. Các file cũ hơn số ngày giữ sẽ bị xóa lúc mở Tool.
"""
import re
import threading
from datetime import datetime, timedelta
from pathlib import Path

LOG_DIR_NAME = "logs"
FILE_PREFIX = "AutoRobot_"
FILE_SUFFIX = ".log"
_NAME_RE = re.compile(r"^AutoRobot_(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2})(\d{2})\.log$")


def log_dir(app_dir):
    return Path(app_dir) / LOG_DIR_NAME


def _start_time(path):
    """Thời điểm bắt đầu ghi của file log, hoặc None nếu không phải file do Tool tạo."""
    match = _NAME_RE.match(path.name)
    if not match:
        return None
    try:
        return datetime(*(int(part) for part in match.groups()))
    except ValueError:
        return None


def purge_old(directory, retention_days, now=None):
    """Xóa các file log do Tool tạo và cũ hơn `retention_days` ngày."""
    directory = Path(directory)
    if not directory.is_dir():
        return
    cutoff = (now or datetime.now()) - timedelta(days=retention_days)
    for path in directory.iterdir():
        started = _start_time(path)
        if started is None or started >= cutoff:
            continue  # không phải log của Tool, hoặc còn trong hạn giữ
        try:
            path.unlink()
        except OSError:
            pass  # file đang mở/không xóa được: để lần sau


class SessionLog:
    """File log của một lần mở Tool. Ghi lỗi không được làm Tool dừng."""

    def __init__(self, path=None):
        self.path = path
        self._lock = threading.Lock()
        self._handle = None
        if path is not None:
            try:
                self._handle = open(path, "a", encoding="utf-8")
            except OSError:
                self._handle = None

    def write(self, message):
        if self._handle is None:
            return
        line = f"[{datetime.now():%H:%M:%S}] {message}\n"
        with self._lock:
            try:
                self._handle.write(line)
                self._handle.flush()  # ghi ngay để log còn khi Tool bị tắt đột ngột
            except OSError:
                pass

    def close(self):
        with self._lock:
            if self._handle is not None:
                try:
                    self._handle.close()
                except OSError:
                    pass
                self._handle = None


def open_session(app_dir, retention_days, now=None):
    """Tạo file log cho lần mở Tool này và dọn log cũ. Không ghi được thì trả về log rỗng."""
    directory = log_dir(app_dir)
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except OSError:
        return SessionLog(None)
    purge_old(directory, retention_days, now)
    path = directory / f"{FILE_PREFIX}{(now or datetime.now()):%Y-%m-%d_%H%M%S}{FILE_SUFFIX}"
    return SessionLog(path)
