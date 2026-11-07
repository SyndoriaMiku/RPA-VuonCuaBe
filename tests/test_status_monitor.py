from contextlib import contextmanager

import kiotviet
import status_monitor


def status(url, page, branch=None, rows=0, draft=False):
    return {"url": url, "page": page, "branch": branch, "rows": rows, "draft": draft, "branches": []}


def test_not_connected_is_marked_disconnected():
    view = status_monitor.describe_status(None)
    assert view == {"text": "Chưa kết nối Chrome Debug", "connected": False, "can_open_chrome": True}


def test_connected_even_when_kiotviet_is_not_open():
    view = status_monitor.describe_status(status("https://example.com/", kiotviet.PAGE_OTHER))
    assert view["connected"] is True
    assert view["text"] == "Chưa mở KiotViet"
    assert view["can_open_chrome"] is False


def test_connected_on_other_kiotviet_page():
    view = status_monitor.describe_status(
        status("https://vuoncuabe.kiotviet.vn/man/#/Products", kiotviet.PAGE_OTHER))
    assert view["connected"] is True
    assert view["text"] == "Chưa ở màn hình Kiểm kho"


def test_connected_shows_stocktake_list():
    view = status_monitor.describe_status(
        status("https://vuoncuabe.kiotviet.vn/man/#/StockTakes", kiotviet.PAGE_LIST,
               ("152162", "CS08 - 22 Vũ Trọng Phụng")))
    assert view["connected"] is True
    assert view["text"] == "Cơ sở: CS08 - 22 Vũ Trọng Phụng  |  Danh sách kiểm kho – sẽ tự tạo phiếu mới"


def test_connected_shows_stocktake_rows_and_unsaved_draft():
    view = status_monitor.describe_status(
        status("https://vuoncuabe.kiotviet.vn/man/#/StockTakes/0", kiotviet.PAGE_EDIT,
               ("152162", "CS08 - 22 Vũ Trọng Phụng"), rows=12, draft=True))
    assert view["connected"] is True
    assert "Màn hình phiếu · 12 mã" in view["text"]
    assert "Có bản nháp chưa lưu" in view["text"]


def test_running_is_connected():
    view = status_monitor.describe_status(
        status("https://vuoncuabe.kiotviet.vn/man/#/StockTakes/0", kiotviet.PAGE_EDIT), running=True)
    assert view["connected"] is True
    assert view["text"] == "Đang chạy…"
    assert view["can_open_chrome"] is False


class FakeSession:
    def __init__(self, driver, busy=False):
        self.driver = driver
        self.busy = busy
        self.invalidated = False

    @contextmanager
    def exclusive(self, blocking=True):
        if self.busy and not blocking:
            yield None
            return
        yield self.driver

    def invalidate(self):
        self.invalidated = True


def make_monitor(driver, busy=False, available=True):
    seen = []
    monitor = status_monitor.StatusMonitor(
        FakeSession(driver, busy=busy), seen.append, interval=5, available=lambda: available)
    return monitor, seen


def test_monitor_emits_read_status(monkeypatch):
    driver = object()
    monkeypatch.setattr(kiotviet, "read_status", lambda d: status("https://vuoncuabe.kiotviet.vn/man/#/StockTakes", kiotviet.PAGE_LIST))
    monitor, seen = make_monitor(driver)
    monitor.check_once()
    assert seen[0]["page"] == kiotviet.PAGE_LIST


def test_monitor_says_off_when_chrome_port_is_closed():
    monitor, seen = make_monitor(object(), available=False)
    monitor.check_once()
    assert seen == [None]


def test_monitor_does_nothing_while_session_is_busy():
    monitor, seen = make_monitor(object(), busy=True)
    monitor.check_once()
    assert seen == []


def test_monitor_invalidates_dead_driver(monkeypatch):
    def broken(driver):
        raise RuntimeError("chrome đã đóng")

    monkeypatch.setattr(kiotviet, "read_status", broken)
    monitor, seen = make_monitor(object())
    monitor.check_once()
    assert monitor.session.invalidated
    assert seen == [None]


def test_monitor_stops():
    monitor, _ = make_monitor(object(), available=False)
    monitor.start()
    monitor.stop()
    assert monitor._thread is None


def test_stop_without_wait_returns_even_if_a_check_is_stuck():
    import threading as _threading
    import time as _time

    stuck = _threading.Event()

    class StuckSession:
        @contextmanager
        def exclusive(self, blocking=True):
            stuck.set()
            _time.sleep(5)  # lệnh Selenium bị kẹt
            yield None

    monitor = status_monitor.StatusMonitor(StuckSession(), lambda status: None, interval=60)
    monitor.start()
    stuck.wait(2)
    started = _time.monotonic()
    monitor.stop(wait=False)
    assert _time.monotonic() - started < 0.5
