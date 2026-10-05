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


class FakeElement:
    def __init__(self, text):
        self.text = text

    def get_attribute(self, name):
        return self.text


class FakeDriver:
    def __init__(self, elements):
        self.elements = elements

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


def test_find_matching_row():
    row = FakeElement("SP001 Sữa 10")
    assert kv.find_matching_row(FakeDriver([row]), "SP001") is row
    assert kv.find_matching_row(FakeDriver([row]), "SP999") is False
    assert kv.find_matching_row(FakeDriver([]), "SP001") is False


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
