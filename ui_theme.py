"""Bảng màu tối và các hàm dùng chung cho giao diện.

Tách riêng khỏi `gui.py` vì phần này thuần Python — không cần Tk — nên test được
bằng pytest, còn phần dựng widget vẫn nằm trong `gui.py`.
"""
FONT = "Segoe UI"
RADIUS = 10

# Hướng "Xanh đêm" (spec §7). Chấm màu kết nối nằm ở status_monitor.
PALETTE = {
    "bg": "#0D1117",
    "surface": "#161B22",
    "surface_alt": "#1C232D",
    "hover": "#222C39",
    "border": "#263040",
    "text": "#E6EDF3",
    "muted": "#8B98A5",
    "accent": "#4C8DFF",
    "accent_hover": "#6BA1FF",
    "on_accent": "#0B1220",
    "ok": "#43C77E",
    "warn": "#E3B341",
    "error": "#F0666C",
    "link": "#7CB6FF",
}


def _channels(color):
    """Tách '#RRGGBB' thành (r, g, b); ném ValueError nếu sai định dạng."""
    if not isinstance(color, str) or len(color) != 7 or not color.startswith("#"):
        raise ValueError(f"Màu không hợp lệ: {color!r} (cần dạng #RRGGBB)")
    try:
        return tuple(int(color[i:i + 2], 16) for i in (1, 3, 5))
    except ValueError:
        raise ValueError(f"Màu không hợp lệ: {color!r} (cần dạng #RRGGBB)") from None


def blend(color_a, color_b, t):
    """Trộn hai màu: t=0 ra color_a, t=1 ra color_b."""
    a = _channels(color_a)
    b = _channels(color_b)
    parts = [round(x + (y - x) * t) for x, y in zip(a, b)]
    return "#" + "".join(f"{value:02X}" for value in parts)


def contrast_ratio(color_a, color_b):
    """Tỉ lệ tương phản WCAG giữa hai màu."""
    def luminance(color):
        channels = []
        for value in _channels(color):
            c = value / 255
            channels.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
        return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]

    high, low = sorted((luminance(color_a), luminance(color_b)), reverse=True)
    return (high + 0.05) / (low + 0.05)


def tween(schedule, duration_ms, frame_ms=16, on_frame=None, on_done=None):
    """Gọi on_frame(t) với t chạy 0.0 → 1.0 rồi on_done().

    `schedule(delay_ms, callback)` do nơi gọi cấp: `root.after` khi chạy thật,
    hoặc hàm giả trong test. Không tự ngủ, nên không bao giờ chặn giao diện.
    """
    total, step = max(1, int(duration_ms)), max(1, int(frame_ms))

    def frame(index):
        t = min(1.0, index * step / total)
        if on_frame:
            on_frame(t)
        if t >= 1.0:
            if on_done:
                on_done()
        else:
            schedule(step, lambda: frame(index + 1))

    frame(0)


def enable_dark_title_bar(hwnd, setter):
    """Làm tối thanh tiêu đề Windows; trả True nếu thành công.

    `setter(attribute, value)` là hàm gọi DwmSetWindowAttribute. Thử thuộc tính
    mới (20) rồi thuộc tính cũ (19); máy không hỗ trợ thì bỏ qua êm, không lỗi.
    """
    for attribute in (20, 19):
        try:
            setter(attribute, 1)
            return True
        except Exception:
            continue
    return False
