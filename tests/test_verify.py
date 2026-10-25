import verify


def item(code, qty, row):
    return {"code": code, "qty": qty, "row": row}


def web(stt, code, adjustment):
    return {"stt": stt, "code": code, "adjustment": adjustment}


def test_all_matching():
    report = verify.compare_adjustments(
        [item("A", 1, 4), item("B", -2, 5), item("C", 0, 6)],
        [web("3", "C", 0), web("2", "B", -2.0), web("1", "a", 1)],
    )
    assert report.ok and report.matched == 3


def test_only_filled_items_are_checked():
    # NEG bị bỏ qua lúc chạy (tổng âm) nên không nằm trong danh sách đã điền: không báo lại
    # OTHER có trên phiếu nhưng không phải do Tool điền: cũng không báo
    report = verify.compare_adjustments(
        [item("7350107133709_D26.10", 1, 20)],
        [web("3", "NEG", 1), web("2", "OTHER", 5), web("1", "7350107133709_D26.10", 2)],
    )
    assert report.matched == 0
    assert report.mismatched == [("7350107133709_D26.10", ["1"], 2, 1)]
    assert report.missing == []
    assert not report.ok


def test_filled_item_not_on_stocktake_is_reported():
    report = verify.compare_adjustments([item("A", 3, 7)], [web("1", "B", 3)])
    assert report.missing == [("A", [7], 3)]
    assert not report.ok


def test_duplicate_codes_in_excel_are_summed():
    report = verify.compare_adjustments([item("A", 1, 4), item("a", 2, 9)], [web("1", "A", 3)])
    assert report.ok and report.matched == 1


def test_describe_report():
    report = verify.compare_adjustments(
        [item("A", 1, 4), item("MISS", 3, 7), item("OK", 2, 8)],
        [web("12", "A", 2), web("13", "OK", 2)],
    )
    lines = verify.describe_report(report)
    assert lines[0] == ("error", "SAI LỆCH: Mã A (STT 12 trên KiotViet) - SL lệch trên web 2, Excel 1")
    assert lines[1] == ("error", "SAI LỆCH: Mã MISS (dòng 7 Excel) đã điền nhưng không thấy trên phiếu - Excel 3")
    assert lines[-1] == ("error", "Khớp 1/3 mã đã điền | Sai lệch: 2")

    ok_lines = verify.describe_report(verify.compare_adjustments([item("A", 1, 4)], [web("1", "A", 1)]))
    assert ok_lines == [("ok", "Tất cả 1 mã đã điền đều khớp SL lệch với Excel.")]
