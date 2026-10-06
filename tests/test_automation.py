import threading

import automation
import kiotviet


def run(items, process, stop_event=None):
    logs, progress = [], []
    stats = automation.run_items(
        None, items, 1, stop_event or threading.Event(),
        lambda msg, level: logs.append((level, msg)),
        lambda done, total: progress.append((done, total)),
        process=process,
    )
    return stats, logs, progress


def item(code, qty, row):
    return {"code": code, "qty": qty, "row": row}


def test_counts_each_outcome_and_continues():
    def process(driver, code, qty, timeout):
        if code == "SKIP":
            raise kiotviet.SkipItem("không có")
        if code == "BAD":
            raise kiotviet.VerifyError("lệch")
        if code == "BOOM":
            raise RuntimeError("hỏng\nchi tiết thêm")

    items = [item("OK", 1, 4), item("SKIP", 1, 5), item("BAD", 1, 6), item("BOOM", 1, 7), item("OK2", 1, 8)]
    stats, logs, progress = run(items, process)

    assert (stats.ok, stats.skipped, stats.error, stats.stopped) == (2, 1, 2, False)
    assert progress[-1] == (5, 5)
    assert [level for level, _ in logs] == ["warn", "error", "error"]
    assert "dòng 7" in logs[-1][1] and "chi tiết thêm" not in logs[-1][1]


def test_invalid_quantity_is_skipped_without_touching_web():
    called = []
    stats, logs, _ = run([item("A", None, 4)], lambda *args: called.append(args))
    assert called == [] and stats.skipped == 1
    assert "không phải là số" in logs[0][1]


def test_stop_event_halts_before_next_item():
    stop = threading.Event()

    def process(driver, code, qty, timeout):
        stop.set()

    stats, _, _ = run([item("A", 1, 4), item("B", 1, 5)], process, stop)
    assert stats.ok == 1 and stats.stopped is True


def test_label_uses_web_stt_when_item_is_on_stocktake():
    def process(driver, code, qty, timeout):
        if code == "NEG":
            raise kiotviet.SkipItem("tổng âm", web_stt="117")
        raise kiotviet.SkipItem("không tìm thấy")

    _, logs, _ = run([item("NEG", -5, 4), item("MISS", 1, 9)], process)
    assert logs[0][1] == "BỎ QUA: Mã NEG (STT 117 trên KiotViet) - tổng âm"
    assert logs[1][1] == "BỎ QUA: Mã MISS (dòng 9 Excel) - không tìm thấy"
