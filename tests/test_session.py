import socket
import threading

import session


class FakeDriver:
    def __init__(self, name):
        self.name = name
        self.quit_called = False

    def quit(self):
        self.quit_called = True


def test_get_creates_one_driver_and_reuses_it():
    created = []

    def connect():
        created.append(FakeDriver(len(created)))
        return created[-1]

    shared = session.DriverSession(connect)
    first = shared.get()
    second = shared.get()
    assert first is second and len(created) == 1


def test_get_returns_none_and_retries_after_connection_error():
    calls = []

    def flaky():
        calls.append(1)
        if len(calls) == 1:
            raise RuntimeError("chrome chưa mở")
        return FakeDriver("ok")

    shared = session.DriverSession(flaky)
    assert shared.get() is None
    assert shared.get().name == "ok"


def test_invalidate_quits_and_reconnects():
    created = []

    def connect():
        created.append(FakeDriver(len(created)))
        return created[-1]

    shared = session.DriverSession(connect)
    first = shared.get()
    shared.invalidate()
    assert first.quit_called
    second = shared.get()
    assert second is not first and len(created) == 2


def test_invalidate_is_safe_when_not_connected():
    session.DriverSession(lambda: None).invalidate()


def test_exclusive_nonblocking_returns_none_when_busy():
    shared = session.DriverSession(lambda: FakeDriver("ok"))
    result = {}
    started = threading.Event()

    def other_thread():
        with shared.exclusive(blocking=False) as driver:
            result["driver"] = driver
        started.set()

    with shared.exclusive() as driver:
        assert driver is not None
        thread = threading.Thread(target=other_thread)
        thread.start()
        thread.join(timeout=2)

    assert started.is_set()
    assert result["driver"] is None


def test_exclusive_releases_lock_after_error():
    shared = session.DriverSession(lambda: FakeDriver("ok"))
    try:
        with shared.exclusive():
            raise RuntimeError("lỗi trong lúc chạy")
    except RuntimeError:
        pass

    with shared.exclusive(blocking=False) as driver:
        assert driver is not None


def test_debug_port_available_detects_open_and_closed_port():
    listener = socket.socket()
    listener.bind(("127.0.0.1", 0))
    listener.listen(1)
    port = listener.getsockname()[1]
    try:
        assert session.debug_port_available(port, timeout=0.5)
    finally:
        listener.close()

    assert not session.debug_port_available(port, timeout=0.2)
