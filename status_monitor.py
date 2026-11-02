import threading

import kiotviet

STATUS_COLORS = {
    "off": "#6B7785",
    "warn": "#B45309",
    "ok": "#2E7D32",
    "busy": "#2E7D32",
}


def describe_status(status, running=False):
    """Chuyển kết quả kiotviet.read_status thành chữ + màu cho thanh trạng thái."""
    if running:
        return {"text": "Đang chạy…", "level": "busy", "can_open_chrome": False}

    if status is None:
        return {"text": "Chưa kết nối Chrome Debug", "level": "off", "can_open_chrome": True}

    url = str(status.get("url") or "").lower()
    if "kiotviet.vn" not in url:
        return {"text": "Chưa mở KiotViet", "level": "warn", "can_open_chrome": False}

    page = status.get("page")
    if page == kiotviet.PAGE_LIST:
        page_text = "Danh sách kiểm kho – sẽ tự tạo phiếu mới"
        level = "ok"
    elif page == kiotviet.PAGE_EDIT:
        page_text = f"Màn hình phiếu · {int(status.get('rows') or 0)} mã"
        level = "ok"
        if status.get("draft"):
            page_text += "  |  ⚠ Có bản nháp chưa lưu"
    else:
        return {"text": "Chưa ở màn hình Kiểm kho", "level": "warn", "can_open_chrome": False}

    branch = status.get("branch")
    text = f"Cơ sở: {branch[1]}  |  {page_text}" if branch else page_text
    return {"text": text, "level": level, "can_open_chrome": False}


class StatusMonitor:
    """Luồng nền đọc trạng thái KiotViet định kỳ, không giành focus và không chồng lệnh với lúc chạy."""

    def __init__(self, driver_session, on_status, interval=5.0, available=None):
        self.session = driver_session
        self.on_status = on_status
        self.interval = max(1.0, float(interval))
        self.available = available or (lambda: True)
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._thread = None

    def start(self):
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, name="status-monitor", daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        self._wake.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
            self._thread = None

    def refresh(self):
        self._wake.set()

    def set_interval(self, seconds):
        self.interval = max(1.0, float(seconds))
        self.refresh()

    def _run(self):
        while not self._stop.is_set():
            self.check_once()
            self._wake.wait(self.interval)
            self._wake.clear()

    def check_once(self):
        try:
            if not self.available():
                self.on_status(None)
                return

            with self.session.exclusive(blocking=False) as driver:
                if driver is None:
                    return  # đang chạy/đang đổi cơ sở: giữ nguyên trạng thái trên giao diện
                try:
                    status = kiotviet.read_status(driver)
                except Exception:
                    self.session.invalidate()
                    self.on_status(None)
                    return
                self.on_status(status)
        except Exception:
            self.on_status(None)
