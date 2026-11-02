"""So SL lệch trên phiếu kiểm kho KiotViet với số lượng trong Excel của các mã Tool đã điền."""
from dataclasses import dataclass, field

TOLERANCE = 1e-6


@dataclass
class VerifyReport:
    matched: int = 0
    mismatched: list = field(default_factory=list)  # [(mã, [STT], SL lệch web, SL Excel)]
    missing: list = field(default_factory=list)     # [(mã, [dòng Excel], SL Excel)]

    @property
    def ok(self):
        return not (self.mismatched or self.missing)


def compare_adjustments(filled_items, web_rows):
    """filled_items: các dòng Excel Tool đã điền thành công [{code, qty, row}];
    web_rows: toàn bộ dòng của phiếu [{stt, code, adjustment}].

    Chỉ kiểm tra các mã đã điền; mã bị bỏ qua (tổng âm, không tìm thấy) đã được báo
    lúc chạy nên không xét ở đây. So theo mã (không phân biệt hoa/thường), mã trùng cộng dồn.
    """
    expected, excel_rows, display = {}, {}, {}
    for item in filled_items:
        key = item["code"].strip().lower()
        display.setdefault(key, item["code"].strip())
        expected[key] = expected.get(key, 0.0) + item["qty"]
        excel_rows.setdefault(key, []).append(item["row"])

    actual, web_stts = {}, {}
    for row in web_rows:
        key = row["code"].strip().lower()
        if key not in expected:
            continue
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
    return report


def describe_report(report, format_qty=lambda v: f"{v:g}"):
    """Các dòng (mức độ, nội dung, mã hàng hoặc None) để ghi nhật ký."""
    lines = []
    for code, stts, web, excel in report.mismatched:
        lines.append(("error", f"Lệch {code}: web {format_qty(web)}, Excel {format_qty(excel)} "
                               f"(STT {', '.join(stts)})", code))
    for code, rows, excel in report.missing:
        lines.append(("error", f"Thiếu {code} trên phiếu, Excel {format_qty(excel)} "
                               f"(dòng {', '.join(map(str, rows))})", code))

    total = report.matched + len(report.mismatched) + len(report.missing)
    if report.ok:
        lines.append(("ok", f"Tất cả {total} mã khớp SL lệch với Excel.", None))
    else:
        lines.append(("error", f"Khớp {report.matched}/{total}, sai lệch: "
                               f"{len(report.mismatched) + len(report.missing)}", None))
    return lines
