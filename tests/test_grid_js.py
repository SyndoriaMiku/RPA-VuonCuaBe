"""Chạy các hàm JS tìm bảng/ô của kiotviet.py trên trang HTML mô phỏng KiotViet bằng Chrome headless.

Trang mô phỏng đúng cấu trúc thật (đã kiểm tra trên KiotViet 10/2026): bảng Kendo có tiêu đề và
thân là hai <table> riêng trong cùng .k-grid, ô Mã hàng mang class "cell-114" (không còn "cell-code"),
và một bảng khác (lịch sử) đứng trước bảng hàng hóa.
"""
import pathlib

import pytest

import kiotviet as kv

HEADER = [
    ("ProductId", "cell-del", ""), ("rowNumber", "cell-order txtC", "STT"), ("Image", "tdImage", ""),
    ("hiddenActualCount", "", ""), ("hiddenAdjustmentValue", "", ""), ("ProductCode", "cell-114", "Mã hàng"),
    ("ProductName", "cell-auto cell-180", "Tên hàng"), ("Unit", "cell-dvt", "ĐVT"),
    ("SystemCount", "cell-quantity txtR", "Tồn kho"), ("ActualCount", "cell-order-number txtR", "Thực tế"),
    ("AdjustmentValue", "cell-quantity txtR", "SL lệch"), ("AdjustmentPrice", "cell-total txtR", "Giá trị lệch"),
    ("ProductSame", "cell-productsame", ""),
]


def item_row(stt, code, on_hand, actual, adjustment, code_class="cell-114"):
    cells = ['<td class="cell-del"></td>', f'<td class="cell-order">{stt}</td>', '<td class="tdImage"></td>',
             '<td style="display:none">1</td>', '<td style="display:none">0</td>',
             f'<td class="{code_class}">{code}</td>', '<td class="cell-auto">Sữa</td>', '<td class="cell-dvt"></td>',
             f'<td class="cell-quantity txtR">{on_hand}</td>',
             f'<td class="cell-order-number"><input value="{actual}"></td>',
             f'<td class="cell-quantity txtR">{adjustment}</td>', '<td class="cell-total txtR">0</td>',
             '<td class="cell-productsame"></td>']
    return "<tr>" + "".join(cells) + "</tr>"


def page(rows, with_data_field=True, draft_dialog=False):
    def th(field, cls, text):
        attr = f' data-field="{field}"' if with_data_field else ""
        return f'<th class="{cls}"{attr}>{text}</th>'

    dialog = ('<div class="k-widget k-window k-window-alert"><span>Mở bản nháp</span>'
              '<p>Hệ thống tìm được 1 bản nháp chưa được lưu lên máy chủ.</p>'
              '<button class="kv-btn-cancel" onclick="this.closest(\'.k-window\').style.display=\'none\'">Bỏ qua</button>'
              '<button class="kv-btn-confirm">Đồng ý</button></div>') if draft_dialog else ""
    return f"""<!doctype html><html><body>
      <input id="productSearchInput">
      <div class="k-grid" id="grdAuditTrails">
        <table role="grid"><thead><tr><th data-field="ProductCode">Mã</th><th data-field="date">Thời gian</th></tr></thead></table>
        <table role="grid"><tbody><tr><td>KHAC-01</td><td>01/10</td></tr></tbody></table>
      </div>
      <div class="k-grid" id="cartGrid">
        <table role="grid"><thead><tr>{''.join(th(*h) for h in HEADER)}</tr></thead></table>
        <table role="grid"><tbody>{''.join(rows)}</tbody></table>
      </div>
      {dialog}
    </body></html>"""


@pytest.fixture(scope="module")
def browser():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    options = Options()
    options.add_argument("--headless=new")
    try:
        driver = webdriver.Chrome(options=options)
    except Exception as exc:
        pytest.skip(f"Không chạy được Chrome headless: {exc}")
    yield driver
    driver.quit()


@pytest.fixture
def open_page(browser, tmp_path):
    def _open(html, hash_path="#/StockTakes/0"):
        path = tmp_path / "kv.html"
        path.write_text(html, encoding="utf-8")
        browser.get(pathlib.Path(path).as_uri() + hash_path)
        return browser
    return _open


ROWS = [item_row("2", "7350107133709", "3", "2", "-1"), item_row("1", "7350107133709_D26.10", "-1", "1", "2")]


def test_reads_cells_by_data_field_in_the_item_grid(open_page):
    d = open_page(page(ROWS))
    assert kv.count_item_rows(d) == 2
    row = kv.first_item_row(d)
    assert kv.cell_text(d, row, kv.FIELD_CODE) == "7350107133709"
    assert kv.cell_text(d, row, kv.FIELD_ON_HAND) == "3"
    assert kv.read_row_order(d, row) == "2"
    assert kv.read_first_row_input(d) == "2"
    assert kv._read_visible_rows(d) == [
        {"stt": "2", "code": "7350107133709", "adjustment": "-1"},
        {"stt": "1", "code": "7350107133709_D26.10", "adjustment": "2"},
    ]


def test_find_matching_row_with_renamed_code_class(open_page):
    d = open_page(page(ROWS))
    assert kv.find_matching_row(d, "7350107133709")
    assert kv.find_matching_row(d, "7350107133709_D26.10") is False  # chỉ xét dòng đầu
    assert kv.find_matching_row(d, "KHAC-01") is False  # không lấy nhầm bảng lịch sử đứng trước


def test_falls_back_to_old_classes_without_data_field(open_page):
    d = open_page(page([item_row("5", "SP001", "4", "4", "0", code_class="cell-code")], with_data_field=False))
    row = kv.first_item_row(d)
    assert kv.cell_text(d, row, kv.FIELD_CODE) == "SP001"
    assert kv.cell_text(d, row, kv.FIELD_ADJUSTMENT) == "0"
    assert kv.find_matching_row(d, "SP001")


def test_draft_dialog_is_dismissed_on_stocktake_page(open_page):
    d = open_page(page(ROWS, draft_dialog=True))
    assert kv.ensure_stocktake_page(d, timeout=5) == ("ready", True)
    assert kv.ensure_stocktake_page(d, timeout=5) == ("ready", False)


def test_dialog_without_draft_text_is_left_alone(open_page):
    html = page(ROWS).replace("</body>", '<div class="k-window"><p>Thông báo khác</p>'
                                         '<button class="kv-btn-cancel">Bỏ qua</button></div></body>')
    d = open_page(html)
    assert kv.dismiss_draft_dialog(d) is False
