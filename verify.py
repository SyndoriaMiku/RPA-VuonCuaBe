"""So số lượng cân chỉnh trong Excel với cột SL lệch trên phiếu kiểm kho KiotViet."""
from dataclasses import dataclass, field

TOLERANCE = 1e-6


@dataclass
class VerifyReport:
    matched: int = 0
    mismatched: list = field(default_factory=list)  # [(mã, [STT], SL lệch web, SL Excel)]
    missing: list = field(default_factory=list)     # [(mã, [dòng Excel], SL Excel)]
    extra: list = field(default_factory=list)       # [(mã, [STT], SL lệch web)]

    @property
    def ok(self):
        return not (self.mismatched or self.missing or self.extra)


def compare_adjustments(items, web_rows):
    """items: dữ liệu Excel [{code, qty, row}]; web_rows: [{stt, code, adjustment}].

    So theo mã hàng (không phân biệt hoa/thường). Mã trùng được cộng dồn ở cả hai phía.
    Dòng Excel có số lượng không phải số thì không tính vào phần mong đợi.
    """
    expected, excel_rows, display = {}, {}, {}
    for item in items:
        if item["qty"] is None:
            continue
        key = item["code"].strip().lower()
        display.setdefault(key, item["code"].strip())
        expected[key] = expected.get(key, 0.0) + item["qty"]
        excel_rows.setdefault(key, []).append(item["row"])

    actual, web_stts = {}, {}
    for row in web_rows:
        key = row["code"].strip().lower()
        display.setdefault(key, row["code"].strip())
        actual[key] = actual.get(key, 0.0) + row["adjustment"]
        web_stts.setdefault(key, []).append(row["stt"])

    report = VerifyReport()
    for key, qty in expected.items():
        if key not in actual:
            report.missing.append((display[key], excel_rows[key], qty))
        elif abs(actual[key] - qty) > TOLERANCE:
            report.mismatched.append((display[key], web_stts[key], actual[key], qty))
        else:
            report.matched += 1
    for key, adjustment in actual.items():
        if key not in expected:
            report.extra.append((display[key], web_stts[key], adjustment))
    return report


def describe_report(report, format_qty=lambda v: f"{v:g}"):
    """Các dòng (mức độ, nội dung) để ghi nhật ký."""
    lines = []
    for code, stts, web, excel in report.mismatched:
        lines.append(("error", f"SAI: Mã {code} (STT {', '.join(stts)} trên KiotViet) - "
                               f"SL lệch trên web {format_qty(web)}, Excel {format_qty(excel)}"))
    for code, rows, excel in report.missing:
        lines.append(("warn", f"KHÔNG CÓ TRÊN PHIẾU: Mã {code} (dòng {', '.join(map(str, rows))} Excel) - "
                              f"Excel {format_qty(excel)}"))
    for code, stts, web in report.extra:
        lines.append(("warn", f"THỪA TRÊN PHIẾU: Mã {code} (STT {', '.join(stts)} trên KiotViet) không có trong Excel - "
                              f"SL lệch {format_qty(web)}"))

    total = report.matched + len(report.mismatched) + len(report.missing)
    summary = (f"Khớp {report.matched}/{total} mã | Sai: {len(report.mismatched)} | "
               f"Không có trên phiếu: {len(report.missing)} | Thừa trên phiếu: {len(report.extra)}")
    lines.append(("ok" if report.ok else "error", summary))
    return lines
