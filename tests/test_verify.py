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


def test_mismatch_missing_and_extra():
    report = verify.compare_adjustments(
        [item("7350107133709_D26.10", 1, 20), item("7350107133709", -1, 21), item("MISS", 3, 22)],
        [web("2", "7350107133709", -1), web("1", "7350107133709_D26.10", 2), web("5", "EXTRA", 1)],
    )
    assert report.matched == 1
    assert report.mismatched == [("7350107133709_D26.10", ["1"], 2, 1)]
    assert report.missing == [("MISS", [22], 3)]
    assert report.extra == [("EXTRA", ["5"], 1)]
    assert not report.ok


def test_duplicate_codes_in_excel_are_summed():
    report = verify.compare_adjustments([item("A", 1, 4), item("a", 2, 9)], [web("1", "A", 3)])
    assert report.ok and report.matched == 1


def test_invalid_excel_quantity_is_not_expected():
    report = verify.compare_adjustments([item("A", None, 4)], [])
    assert report.ok and report.matched == 0


def test_describe_report():
    report = verify.compare_adjustments(
        [item("A", 1, 4), item("MISS", 3, 7)],
        [web("12", "A", 2), web("13", "EXTRA", 1)],
    )
    lines = verify.describe_report(report)
    texts = [text for _, text in lines]
    assert lines[0] == ("error", "SAI: Mã A (STT 12 trên KiotViet) - SL lệch trên web 2, Excel 1")
    assert "KHÔNG CÓ TRÊN PHIẾU: Mã MISS (dòng 7 Excel)" in texts[1]
    assert "THỪA TRÊN PHIẾU: Mã EXTRA (STT 13 trên KiotViet)" in texts[2]
    assert lines[-1] == ("error", "Khớp 0/2 mã | Sai: 1 | Không có trên phiếu: 1 | Thừa trên phiếu: 1")

    ok_lines = verify.describe_report(verify.compare_adjustments([item("A", 1, 4)], [web("1", "A", 1)]))
    assert ok_lines == [("ok", "Khớp 1/1 mã | Sai: 0 | Không có trên phiếu: 0 | Thừa trên phiếu: 0")]
