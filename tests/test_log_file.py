import re
from datetime import datetime

import log_file


def test_open_session_creates_file_named_by_start_time(tmp_path):
    session = log_file.open_session(tmp_path, retention_days=7, now=datetime(2026, 11, 9, 14, 30, 5))
    session.write("Xin chào")
    session.close()
    path = tmp_path / "logs" / "AutoRobot_2026-11-09_143005.log"
    assert re.fullmatch(r"\[\d{2}:\d{2}:\d{2}\] Xin chào\n", path.read_text(encoding="utf-8"))


def test_session_appends_lines_in_order(tmp_path):
    session = log_file.open_session(tmp_path, retention_days=7, now=datetime(2026, 11, 9, 8, 0, 0))
    session.write("dòng 1")
    session.write("dòng 2")
    session.close()
    lines = (tmp_path / "logs" / "AutoRobot_2026-11-09_080000.log").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 2
    assert lines[0].endswith("dòng 1")
    assert lines[1].endswith("dòng 2")


def test_purge_deletes_logs_older_than_retention(tmp_path):
    directory = log_file.log_dir(tmp_path)
    directory.mkdir()
    old = directory / "AutoRobot_2026-11-01_080000.log"
    keep = directory / "AutoRobot_2026-11-08_080000.log"
    for path in (old, keep):
        path.write_text("x", encoding="utf-8")

    log_file.purge_old(directory, retention_days=7, now=datetime(2026, 11, 9, 8, 0, 0))

    assert not old.exists()
    assert keep.exists()


def test_purge_keeps_files_that_are_not_our_logs(tmp_path):
    directory = log_file.log_dir(tmp_path)
    directory.mkdir()
    other = directory / "ghi-chu.txt"
    renamed = directory / "nhat-ky-cu.log"
    for path in (other, renamed):
        path.write_text("giữ lại", encoding="utf-8")

    log_file.purge_old(directory, retention_days=1, now=datetime(2030, 1, 1))

    assert other.exists()
    assert renamed.exists()


def test_purge_ignores_missing_directory(tmp_path):
    log_file.purge_old(tmp_path / "khong-co", retention_days=7, now=datetime(2026, 11, 9))


def test_open_session_removes_old_logs_first(tmp_path):
    old = log_file.open_session(tmp_path, retention_days=7, now=datetime(2026, 10, 1, 9, 0, 0))
    old.write("cũ")
    old.close()

    session = log_file.open_session(tmp_path, retention_days=7, now=datetime(2026, 11, 9, 9, 0, 0))
    session.close()

    assert not (tmp_path / "logs" / "AutoRobot_2026-10-01_090000.log").exists()
    assert (tmp_path / "logs" / "AutoRobot_2026-11-09_090000.log").exists()


def test_session_write_survives_unwritable_directory(tmp_path, monkeypatch):
    def fail_mkdir(self, *args, **kwargs):
        raise OSError("không ghi được")

    monkeypatch.setattr(log_file.Path, "mkdir", fail_mkdir)

    session = log_file.open_session(tmp_path, retention_days=7)
    session.write("không được ghi")  # không được ném lỗi
    session.close()
