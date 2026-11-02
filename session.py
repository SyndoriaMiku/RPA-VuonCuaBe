"""Phiên Selenium dùng chung cho lúc chạy, đổi cơ sở và thanh trạng thái.

Chỉ một driver được giữ lại và dùng lần lượt dưới một lock, để các luồng nền không gửi
lệnh chồng lên nhau lên cùng một tab Chrome.
"""
import socket
import threading
from contextlib import contextmanager


class DriverSession:
    def __init__(self, connect):
        self._connect = connect
        self._lock = threading.RLock()
        self._driver = None

    def get(self):
        """Driver dùng chung, hoặc None nếu không kết nối được."""
        with self._lock:
            return self._ensure_driver()

    @contextmanager
    def exclusive(self, blocking=True):
        """Giữ lock trong suốt thời gian dùng driver; trả None nếu đang bận và blocking=False."""
        if not self._lock.acquire(blocking=blocking):
            yield None
            return
        try:
            yield self._ensure_driver()
        finally:
            self._lock.release()

    def invalidate(self):
        """Bỏ driver hiện tại (Chrome đóng/mở lại) và tạo lại ở lần dùng sau."""
        with self._lock:
            driver, self._driver = self._driver, None
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

    def close(self):
        self.invalidate()

    def _ensure_driver(self):
        if self._driver is None:
            try:
                self._driver = self._connect()
            except Exception:
                self._driver = None
        return self._driver


def debug_port_available(port, timeout=0.3):
    """Chrome Debug có đang mở cổng này không (kiểm tra rẻ, không khởi tạo chromedriver)."""
    try:
        with socket.create_connection(("127.0.0.1", int(port)), timeout=timeout):
            return True
    except (OSError, ValueError):
        return False
