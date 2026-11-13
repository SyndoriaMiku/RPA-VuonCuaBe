import pytest

import ui_theme


def test_palette_has_all_roles_and_valid_hex():
    required = {"bg", "surface", "surface_alt", "hover", "border", "text", "muted",
                "accent", "accent_hover", "on_accent", "ok", "warn", "error", "link"}
    assert required <= set(ui_theme.PALETTE)
    for value in ui_theme.PALETTE.values():
        assert len(value) == 7 and value.startswith("#")


def test_palette_values_match_spec():
    assert ui_theme.PALETTE["bg"] == "#0D1117"
    assert ui_theme.PALETTE["accent"] == "#4C8DFF"
    assert ui_theme.PALETTE["on_accent"] == "#0B1220"


def test_blend_returns_endpoints_and_quarter():
    assert ui_theme.blend("#000000", "#FFFFFF", 0.0) == "#000000"
    assert ui_theme.blend("#000000", "#FFFFFF", 1.0) == "#FFFFFF"
    assert ui_theme.blend("#000000", "#FFFFFF", 0.25) == "#404040"


def test_blend_rejects_bad_color():
    with pytest.raises(ValueError):
        ui_theme.blend("do", "#FFFFFF", 0.5)
    with pytest.raises(ValueError):
        ui_theme.blend("#FFF", "#FFFFFF", 0.5)


def test_contrast_ratio_of_key_pairs_is_accessible():
    assert ui_theme.contrast_ratio("#E6EDF3", "#0D1117") >= 4.5
    assert ui_theme.contrast_ratio("#E6EDF3", "#161B22") >= 4.5
    assert ui_theme.contrast_ratio("#0B1220", "#4C8DFF") >= 4.5
    assert ui_theme.contrast_ratio("#8B98A5", "#161B22") >= 4.5


def test_tween_runs_frames_to_one_then_done():
    scheduled, frames, done = [], [], []

    def schedule(delay_ms, callback):
        scheduled.append(callback)

    ui_theme.tween(schedule, duration_ms=160, frame_ms=16,
                   on_frame=frames.append, on_done=lambda: done.append(True))
    while scheduled:
        scheduled.pop(0)()

    assert frames[0] == 0.0
    assert frames[-1] == 1.0
    assert frames == sorted(frames)
    assert len(frames) == 11
    assert done == [True]


def test_dark_title_bar_swallows_failure():
    def boom(attribute, value):
        raise OSError("unsupported")

    assert ui_theme.enable_dark_title_bar(1234, boom) is False

    seen = []

    def ok(attribute, value):
        seen.append(attribute)
        return True

    assert ui_theme.enable_dark_title_bar(1234, ok) is True
    assert seen == [20]
