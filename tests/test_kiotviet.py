import pytest

import kiotviet as kv


def test_format_qty():
    assert kv.format_qty(5.0) == "5"
    assert kv.format_qty(-3) == "-3"
    assert kv.format_qty(2.5) == "2.5"
    assert kv.format_qty(1000000.0) == "1000000"


def test_parse_web_number():
    assert kv.parse_web_number("1,234") == 1234
    assert kv.parse_web_number(" 12 ") == 12
    assert kv.parse_web_number("") == 0
    with pytest.raises(ValueError):
        kv.parse_web_number("abc")


def test_text_contains_code():
    assert kv.text_contains_code("SP001 - Sữa tươi", "sp001")
    assert not kv.text_contains_code("SP002 - Bánh", "SP001")


def test_text_contains_code_rejects_codes_sharing_a_prefix():
    code = "7350107133709"
    assert not kv.text_contains_code("7350107133709_D26.10 Sữa Semper", code)
    assert not kv.text_contains_code("7350107133709-A Sữa", code)
    assert not kv.text_contains_code("17350107133709 Sữa", code)
    assert kv.text_contains_code("7350107133709 Sữa Semper", code)
    assert kv.text_contains_code("Mã: 7350107133709 (Hộp)", code)
    assert kv.text_contains_code("7350107133709_D26.10 Sữa Semper", "7350107133709_D26.10")


def test_code_equals():
    assert kv.code_equals(" 7350107133709 ", "7350107133709")
    assert kv.code_equals("sp001", "SP001")
    assert not kv.code_equals("7350107133709_D26.10", "7350107133709")


class FakeElement:
    def __init__(self, text, cells=None):
        self.text = text
        self.cells = cells or {}

    def get_attribute(self, name):
        return self.text

    def find_element(self, by, selector):
        if selector not in self.cells:
            raise LookupError(selector)
        return FakeElement(self.cells[selector])


class FakeDriver:
    def __init__(self, elements, exact=None):
        self.elements = elements
        self.exact = exact

    def execute_script(self, script, *args):
        return self.exact

    def find_elements(self, by, selector):
        return self.elements

    def find_element(self, by, selector):
        if not self.elements:
            raise LookupError
        return self.elements[0]


def test_find_matching_result_picks_exact_code_not_first_item():
    stale = FakeElement("OLD01 - Hàng cũ")
    wanted = FakeElement("SP001 - Sữa")
    assert kv.find_matching_result(FakeDriver([stale, wanted]), "SP001") is wanted
    assert kv.find_matching_result(FakeDriver([stale]), "SP001") is False


def test_find_matching_result_skips_code_with_suffix():
    with_suffix = FakeElement("7350107133709_D26.10 Sữa Semper")
    plain = FakeElement("7350107133709 Sữa Semper")
    assert kv.find_matching_result(FakeDriver([with_suffix, plain]), "7350107133709") is plain
    assert kv.find_matching_result(FakeDriver([with_suffix]), "7350107133709") is False


def test_find_matching_result_prefers_exact_element_match():
    exact = FakeElement("exact")
    assert kv.find_matching_result(FakeDriver([FakeElement("7350107133709 x")], exact=exact), "7350107133709") is exact


class GridRow:
    """Dòng bảng giả: cells là {data-field: chữ trong ô}."""

    def __init__(self, cells, tds=None):
        self.cells = cells
        self.tds = tds or []

    def find_elements(self, by, selector):
        return [FakeElement(t) for t in self.tds]


class GridDriver:
    """Driver giả mô phỏng các hàm JS kvItemRows/kvCell/kvText trên bảng hàng hóa."""

    def __init__(self, rows):
        self.rows = rows

    def execute_script(self, script, *args):
        if "kvItemRows()[0]" in script:
            return self.rows[0] if self.rows else None
        if "return kvCell(" in script:
            row, field = args
            return FakeElement(row.cells[field]) if field in row.cells else None
        if "return kvText(" in script:
            row, field = args
            return row.cells.get(field, "").strip()
        raise AssertionError("unexpected script")


def test_find_matching_row():
    row = GridRow({kv.FIELD_CODE: "SP001"})
    assert kv.find_matching_row(GridDriver([row]), "SP001") is row
    assert kv.find_matching_row(GridDriver([row]), "SP999") is False
    assert kv.find_matching_row(GridDriver([]), "SP001") is False


def test_find_matching_row_does_not_accept_previous_row_with_suffix():
    # Dòng đầu vẫn là mã vừa xử lý trước đó: phải chờ dòng mới, không được nhận nhầm
    previous = GridRow({kv.FIELD_CODE: "7350107133709_D26.10"})
    assert kv.find_matching_row(GridDriver([previous]), "7350107133709") is False
    assert kv.find_matching_row(GridDriver([previous]), "7350107133709_D26.10") is previous

    current = GridRow({kv.FIELD_CODE: " 7350107133709 "})
    assert kv.find_matching_row(GridDriver([current]), "7350107133709") is current


def test_row_without_code_column_needs_an_exact_cell():
    # Không xác định được cột Mã hàng: chỉ nhận khi có ô đúng bằng mã, không so gần đúng trên cả dòng
    row = GridRow({}, tds=["2", "7350107133709_D26.10", "Sữa"])
    assert kv.find_matching_row(GridDriver([row]), "7350107133709") is False
    row = GridRow({}, tds=["2", "7350107133709", "Sữa"])
    assert kv.find_matching_row(GridDriver([row]), "7350107133709") is row


def test_merge_note():
    assert kv.merge_note("", "Cân hàng lệch") == "Cân hàng lệch"
    assert kv.merge_note(None, "Cân hàng lệch") == "Cân hàng lệch"
    assert kv.merge_note("Kiểm kho tháng 10", "Cân hàng lệch") == "Kiểm kho tháng 10\nCân hàng lệch"
    assert kv.merge_note("Ghi chú: cân hàng lệch", "Cân hàng lệch") is None


def test_page_state_from_url():
    base = "https://vuoncuabe.kiotviet.vn/man/"
    assert kv.page_state_from_url(base + "#/StockTakes") == kv.PAGE_LIST
    assert kv.page_state_from_url(base + "#/StockTakes/") == kv.PAGE_LIST
    assert kv.page_state_from_url(base + "#/StockTakes/0") == kv.PAGE_EDIT
    assert kv.page_state_from_url(base + "#/StockTakes/12345") == kv.PAGE_EDIT
    assert kv.page_state_from_url(base + "#/Products") == kv.PAGE_OTHER
    assert kv.page_state_from_url(base) == kv.PAGE_OTHER


def test_read_row_order():
    assert kv.read_row_order(GridDriver([]), GridRow({kv.FIELD_STT: " 117 "})) == "117"
    assert kv.read_row_order(GridDriver([]), GridRow({kv.FIELD_STT: ""})) is None
    assert kv.read_row_order(GridDriver([]), GridRow({})) is None


class FakeSwitchTo:
    def window(self, handle):
        pass


class FakeBrowser:
    window_handles = ["tab-1"]
    current_url = "https://vuoncuabe.kiotviet.vn/man/#/StockTakes/0"

    def __init__(self, options):
        self.switch_to = FakeSwitchTo()


def test_connect_driver_skips_bring_to_front_when_focus_is_false(monkeypatch):
    focused = []
    monkeypatch.setattr(kv.webdriver, "Chrome", FakeBrowser)
    monkeypatch.setattr(kv, "bring_to_front", focused.append)

    kv.connect_driver(9999, focus=False)
    assert focused == []

    kv.connect_driver(9999)
    assert len(focused) == 1


class StatusDriver:
    def __init__(self, result):
        self.result = result

    def execute_script(self, script, *args):
        return self.result


def test_read_status_reads_page_branch_rows_and_draft():
    driver = StatusDriver({
        "url": "https://vuoncuabe.kiotviet.vn/man/#/StockTakes/0",
        "branch": ["152162", "CS08 - 22 Vũ Trọng Phụng"],
        "branches": [["152162", "CS08 - 22 Vũ Trọng Phụng"], ["151550", "CS02 - 68 Linh Lang"]],
        "rows": 12,
        "draft": True,
    })
    status = kv.read_status(driver)
    assert status == {
        "url": "https://vuoncuabe.kiotviet.vn/man/#/StockTakes/0",
        "page": kv.PAGE_EDIT,
        "branch": ("152162", "CS08 - 22 Vũ Trọng Phụng"),
        "branches": [("152162", "CS08 - 22 Vũ Trọng Phụng"), ("151550", "CS02 - 68 Linh Lang")],
        "rows": 12,
        "draft": True,
    }


def test_read_status_handles_missing_kiotviet_page():
    status = kv.read_status(StatusDriver({
        "url": "https://example.com/", "branch": None, "branches": [], "rows": 0, "draft": False}))
    assert status["page"] == kv.PAGE_OTHER
    assert status["branch"] is None and status["branches"] == []
    assert status["rows"] == 0 and status["draft"] is False


class RevealDriver:
    """Giả lập bảng Kendo: mỗi trang có sẵn danh sách dòng, đổi trang bằng dataSource.page()."""

    def __init__(self, pages, page=1):
        self.pages = pages
        self.current_page = page
        self.highlighted = []
        self.page_size = len(pages[0])

    def execute_script(self, script, *args):
        if "totalPages" in script:
            return {
                "page": self.current_page,
                "pageSize": self.page_size,
                "totalPages": len(self.pages),
                "total": self.page_size * len(self.pages),
            }
        if "dataSource.page(arguments[0])" in script:
            self.current_page = int(args[0])
            return None
        if "kvItemRows().map" in script:
            return self.pages[self.current_page - 1]
        if "scrollIntoView" in script:
            self.highlighted.append(args[0])
            return True
        raise AssertionError("unexpected script")


def reveal_row(stt, code):
    return {"stt": stt, "code": code, "adjustment": "-1"}


def test_reveal_item_turns_to_the_page_with_the_code():
    driver = RevealDriver([
        [reveal_row("2", "A"), reveal_row("1", "B")],
        [reveal_row("4", "C"), reveal_row("3", "D")],
    ])
    assert kv.reveal_item(driver, "c") == "4"
    assert driver.current_page == 2
    assert driver.highlighted == ["c"]


def test_reveal_item_returns_none_and_restores_page_when_not_found():
    driver = RevealDriver([
        [reveal_row("2", "A"), reveal_row("1", "B")],
        [reveal_row("4", "C"), reveal_row("3", "D")],
    ])
    assert kv.reveal_item(driver, "Z") is None
    assert driver.current_page == 1
    assert driver.highlighted == []


def test_reveal_item_requires_the_item_grid():
    class NoGrid:
        def execute_script(self, script, *args):
            return None

    with pytest.raises(RuntimeError):
        kv.reveal_item(NoGrid(), "A")


class TabDriver:
    """Driver giả có nhiều tab: tabs là {handle: url}."""

    def __init__(self, tabs, current):
        self.tabs = dict(tabs)
        self.current = current
        self.dead = False

    @property
    def window_handles(self):
        if self.dead:
            raise RuntimeError("chrome not reachable")
        return list(self.tabs)

    @property
    def current_window_handle(self):
        if self.current not in self.tabs:
            raise RuntimeError("no such window")
        return self.current

    @property
    def current_url(self):
        if self.current not in self.tabs:
            raise RuntimeError("no such window")
        return self.tabs[self.current]

    @property
    def switch_to(self):
        driver = self

        class Switch:
            def window(self, handle):
                driver.current = handle
        return Switch()


KV_URL = "https://vuoncuabe.kiotviet.vn/man/#/StockTakes"


def test_select_tab_keeps_current_kiotviet_tab():
    d = TabDriver({"a": KV_URL, "b": "https://google.com"}, "a")
    assert kv.select_kiotviet_tab(d) is True and d.current == "a"


def test_select_tab_moves_to_kiotviet_opened_in_new_tab():
    # Tool kết nối khi tab đang trống, sau đó người dùng mở KiotViet ở tab mới
    d = TabDriver({"blank": "about:blank", "new": KV_URL}, "blank")
    assert kv.select_kiotviet_tab(d) is True and d.current == "new"


def test_select_tab_recovers_when_current_tab_was_closed():
    d = TabDriver({"other": KV_URL}, "closed")
    assert kv.select_kiotviet_tab(d) is True and d.current == "other"


def test_select_tab_without_kiotviet_stays_on_a_live_tab():
    d = TabDriver({"x": "about:blank"}, "gone")
    assert kv.select_kiotviet_tab(d) is False and d.current == "x"


def test_select_tab_raises_when_chrome_is_gone():
    d = TabDriver({"a": KV_URL}, "a")
    d.dead = True
    try:
        kv.select_kiotviet_tab(d)
    except RuntimeError:
        pass
    else:
        raise AssertionError("phải báo lỗi để phiên dùng chung tạo lại driver")
