"""Báo khi Tool chạy xong bằng âm thanh hệ thống và nhấp nháy taskbar.

Mọi lỗi của API hệ điều hành đều bị bỏ qua: thông báo không bao giờ được làm hỏng lần chạy.
"""
import sys

try:
    import winsound
except ImportError:  # không phải Windows
    winsound = None

FLASHW_ALL = 0x00000003
FLASHW_TIMERNOFG = 0x0000000C


def notify_done(success, sound=True, flash=True, hwnd=None):
    """Phát âm "OK"/"cảnh báo" và nhấp nháy cửa sổ (nếu biết hwnd)."""
    if sound:
        try:
            _play_system_sound(success)
        except Exception:
            pass
    if flash and hwnd:
        try:
            _flash_window(hwnd)
        except Exception:
            pass


def _play_system_sound(success):
    if winsound is None:
        return
    try:
        winsound.MessageBeep(winsound.MB_OK if success else winsound.MB_ICONEXCLAMATION)
    except Exception:
        pass


def _flash_window(hwnd):
    try:
        import ctypes
        from ctypes import wintypes

        class FlashWindowInfo(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.UINT),
                ("hwnd", wintypes.HWND),
                ("dwFlags", wintypes.DWORD),
                ("uCount", wintypes.UINT),
                ("dwTimeout", wintypes.DWORD),
            ]

        user32 = ctypes.windll.user32
        if user32.GetForegroundWindow() == hwnd:
            return  # cửa sổ đang được xem thì không cần nhấp nháy
        info = FlashWindowInfo(ctypes.sizeof(FlashWindowInfo), hwnd,
                               FLASHW_ALL | FLASHW_TIMERNOFG, 0, 0)
        user32.FlashWindowEx(ctypes.byref(info))
    except Exception:
        pass


def top_window_handle(widget):
    """HWND của cửa sổ cấp cao nhất chứa widget Tk, hoặc None."""
    try:
        return _parent_handle(widget.winfo_id())
    except Exception:
        return None


def _parent_handle(hwnd):
    import ctypes

    return ctypes.windll.user32.GetParent(hwnd)
