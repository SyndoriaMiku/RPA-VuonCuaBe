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


def test_find_matching_row():
    row = FakeElement("SP001 Sữa 10")
    assert kv.find_matching_row(FakeDriver([row]), "SP001") is row
    assert kv.find_matching_row(FakeDriver([row]), "SP999") is False
    assert kv.find_matching_row(FakeDriver([]), "SP001") is False


def test_find_matching_row_does_not_accept_previous_row_with_suffix():
    # Dòng đầu vẫn là mã vừa xử lý trước đó: phải chờ dòng mới, không được nhận nhầm
    previous = FakeElement("152 7350107133709_D26.10 Sữa Semper",
                           cells={kv.CODE_CELL_SELECTOR: "7350107133709_D26.10"})
    assert kv.find_matching_row(FakeDriver([previous]), "7350107133709") is False
    assert kv.find_matching_row(FakeDriver([previous]), "7350107133709_D26.10") is previous

    current = FakeElement("153 7350107133709 Sữa Semper", cells={kv.CODE_CELL_SELECTOR: " 7350107133709 "})
    assert kv.find_matching_row(FakeDriver([current]), "7350107133709") is current


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


class FakeRow:
    def __init__(self, order_text=None):
        self.order_text = order_text

    def find_element(self, by, selector):
        if self.order_text is None:
            raise LookupError
        return FakeElement(self.order_text)


def test_read_row_order():
    assert kv.read_row_order(FakeRow(" 117 ")) == "117"
    assert kv.read_row_order(FakeRow("")) is None
    assert kv.read_row_order(FakeRow(None)) is None
