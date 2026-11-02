import version


def test_default_version_is_1_4_1():
    assert version.__version__ == "1.4.1"


def test_window_title_without_build_hash(monkeypatch):
    monkeypatch.setattr(version, "BUILD", "dev")
    assert version.window_title() == "Auto Robot v1.4.1"


def test_window_title_shows_build_hash_when_built(monkeypatch):
    monkeypatch.setattr(version, "BUILD", "24366f1 2026-11-02")
    assert version.window_title() == "Auto Robot v1.4.1 (24366f1 2026-11-02)"


def test_window_title_shows_file_title_without_build_hash(monkeypatch):
    monkeypatch.setattr(version, "BUILD", "24366f1 2026-11-02")
    assert version.window_title("CS08 - 22 Vũ Trọng Phụng") == "Auto Robot v1.4.1 - CS08 - 22 Vũ Trọng Phụng"


def test_build_text_always_shows_version_and_build(monkeypatch):
    monkeypatch.setattr(version, "BUILD", "dev")
    assert version.build_text() == "v1.4.1 (dev)"
    monkeypatch.setattr(version, "BUILD", "24366f1 2026-11-02")
    assert version.build_text() == "v1.4.1 (24366f1 2026-11-02)"
