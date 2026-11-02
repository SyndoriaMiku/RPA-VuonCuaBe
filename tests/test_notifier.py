import notifier


class FakeWinsound:
    MB_OK = 0
    MB_ICONEXCLAMATION = 48

    def __init__(self, fail=False):
        self.calls = []
        self.fail = fail

    def MessageBeep(self, kind):
        self.calls.append(kind)
        if self.fail:
            raise RuntimeError("không phát được âm thanh")


def test_success_plays_ok_sound(monkeypatch):
    fake = FakeWinsound()
    monkeypatch.setattr(notifier, "winsound", fake)
    notifier.notify_done(True)
    assert fake.calls == [fake.MB_OK]


def test_failure_plays_warning_sound(monkeypatch):
    fake = FakeWinsound()
    monkeypatch.setattr(notifier, "winsound", fake)
    notifier.notify_done(False)
    assert fake.calls == [fake.MB_ICONEXCLAMATION]


def test_sound_can_be_disabled(monkeypatch):
    fake = FakeWinsound()
    monkeypatch.setattr(notifier, "winsound", fake)
    notifier.notify_done(True, sound=False)
    assert fake.calls == []


def test_flashes_only_when_enabled_and_handle_is_known(monkeypatch):
    flashed = []
    monkeypatch.setattr(notifier, "_flash_window", flashed.append)

    notifier.notify_done(True, flash=True, hwnd=None)
    assert flashed == []

    notifier.notify_done(True, flash=False, hwnd=123)
    assert flashed == []

    notifier.notify_done(True, flash=True, hwnd=123)
    assert flashed == [123]


def test_notifier_never_raises_when_platform_apis_fail(monkeypatch):
    monkeypatch.setattr(notifier, "winsound", FakeWinsound(fail=True))

    def broken_flash(hwnd):
        raise OSError("không nhấp nháy được")

    monkeypatch.setattr(notifier, "_flash_window", broken_flash)
    notifier.notify_done(False, flash=True, hwnd=1)  # không được ném lỗi


def test_no_sound_module_is_ignored(monkeypatch):
    monkeypatch.setattr(notifier, "winsound", None)
    notifier.notify_done(True)  # không được ném lỗi


def test_top_window_handle_uses_parent_window(monkeypatch, ):
    class FakeWidget:
        def winfo_id(self):
            return 77

    monkeypatch.setattr(notifier, "_parent_handle", lambda hwnd: hwnd + 100)
    assert notifier.top_window_handle(FakeWidget()) == 177


def test_top_window_handle_returns_none_on_error():
    class BrokenWidget:
        def winfo_id(self):
            raise RuntimeError("cửa sổ đã đóng")

    assert notifier.top_window_handle(BrokenWidget()) is None
