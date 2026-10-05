import pytest

import branch_match as bm

BRANCHES = [
    ("152603", "CS01 - 38 Nguyễn Viết Xuân"),
    ("151550", "CS02 - 68 Linh Lang"),
    ("153113", "CS03 - 134 Lê Lợi"),
    ("152472", "CS04 - 322 Minh Khai"),
    ("152952", "CS05 - Số 70 Trần Bình"),
    ("153185", "CS06 - 97 Nguyễn Hữu Thọ"),
    ("151551", "CS07 - Vinhomes Ocean Park"),
    ("152162", "CS08 - 22 Vũ Trọng Phụng"),
    ("1000000059", "CS09 - 101C1 Nghĩa Tân"),
    ("1000000487", "CS11 - 241 Khâm Thiên"),
    ("323167", "CS13 - 109 Xuân Đỉnh"),
    ("1000000554", "CS16 - 516 Trần Cung"),
    ("120036", "Kênh Zalo"),
    ("323229", "Kiểm kho - 01"),
    ("1000001054", "Kiểm kho - 02"),
    ("323241", "Kho quà tặng"),
    ("323243", "Kho tổng (new)"),
    ("1000000185", "VCB - Ngoại tỉnh"),
    ("1000000961", "VCB - Tiktok"),
]


def ids(title):
    return [branch_id for branch_id, _ in bm.match_branches(title, BRANCHES)]


def test_normalize():
    assert bm.normalize("CS08 - 22 Vũ Trọng Phụng") == "cs08 22 vu trong phung"
    assert bm.normalize("Đỉnh") == "dinh"


@pytest.mark.parametrize("title, expected", [
    ("CÂN CHỈNH HÀNG HÓA CS08", ["152162"]),
    ("Cân chỉnh cs8 tháng 10", ["152162"]),
    ("Kiểm kho CS 05", ["152952"]),
    ("CS13 - Xuân Đỉnh", ["323167"]),
])
def test_match_by_cs_number(title, expected):
    assert ids(title) == expected


@pytest.mark.parametrize("title, expected", [
    ("Cân chỉnh hàng 22 Vũ Trọng Phụng", ["152162"]),
    ("Cân chỉnh Vũ Trọng Phụng", ["152162"]),
    ("Kiểm kho 70 Trần Bình", ["152952"]),
    ("Cơ sở Trần Bình", ["152952"]),
    ("Cân hàng Ocean Park", []),
    ("Cân hàng Vinhomes Ocean Park", ["151551"]),
    ("Cân hàng Kho tổng", ["323243"]),
    ("Đơn Kênh Zalo", ["120036"]),
])
def test_match_by_location(title, expected):
    assert ids(title) == expected


def test_cs_number_wins_over_location():
    assert ids("CS08 (chuyển từ Trần Bình)") == ["152162"]


def test_unknown_cs_number_falls_back_to_location():
    assert ids("CS99 Linh Lang") == ["151550"]


def test_no_match():
    assert ids("Cân chỉnh hàng hóa tháng 10") == []
    assert ids("") == []


def test_number_only_names_never_match_by_location():
    assert bm.location_keys("Kiểm kho - 01") == []
    assert ids("Kiểm kho 01") == []


def test_is_same_branch():
    assert bm.is_same_branch("Cân chỉnh CS08", "CS08 - 22 Vũ Trọng Phụng")
    assert bm.is_same_branch("Cân chỉnh Vũ Trọng Phụng", "CS08 - 22 Vũ Trọng Phụng")
    assert not bm.is_same_branch("Cân chỉnh CS01", "CS08 - 22 Vũ Trọng Phụng")
    assert not bm.is_same_branch("CS01 Vũ Trọng Phụng", "CS08 - 22 Vũ Trọng Phụng")
    assert not bm.is_same_branch("Cân chỉnh hàng hóa", "CS08 - 22 Vũ Trọng Phụng")
