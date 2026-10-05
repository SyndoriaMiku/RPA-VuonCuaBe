"""Khớp cơ sở (chi nhánh) KiotViet với tiêu đề trong file Excel.

Hai cách khớp, theo thứ tự ưu tiên:
1. Trùng số CS (CS8, CS08, cs 08 đều là cơ sở 8).
2. Trùng địa điểm: địa chỉ đầy đủ ("22 Vũ Trọng Phụng") hoặc chỉ tên đường ("Vũ Trọng Phụng").
"""
import re
import unicodedata

MIN_KEY_LENGTH = 4


def normalize(text):
    """Bỏ dấu, chữ thường, chỉ giữ chữ/số, gộp khoảng trắng."""
    text = str(text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return " ".join(text.split())


def parse_cs_number(text):
    match = re.search(r"\bcs\s*0*(\d+)\b", normalize(text))
    return int(match.group(1)) if match else None


def location_keys(branch_name):
    """Các cụm dùng để nhận diện địa điểm của một cơ sở, từ cụ thể đến chung."""
    name = re.sub(r"\(.*?\)", " ", str(branch_name))
    address = name.split(" - ", 1)[1] if " - " in name else name
    address = normalize(address)

    tokens = address.split()
    # Bỏ phần số nhà ở đầu ("so 70", "22", "101c1") để lấy riêng tên đường
    while tokens and (tokens[0] == "so" or any(ch.isdigit() for ch in tokens[0])):
        tokens.pop(0)
    street = " ".join(tokens)

    keys = []
    for key in (address, street):
        if key and key not in keys and len(key) >= MIN_KEY_LENGTH and re.search(r"[a-z]", key):
            keys.append(key)
    return keys


def _contains_phrase(text, phrase):
    return f" {phrase} " in f" {text} "


def match_branches(title, branches):
    """Trả về danh sách cơ sở khớp với tiêu đề. branches: [(id, tên hiển thị)]."""
    cs_number = parse_cs_number(title)
    if cs_number is not None:
        by_number = [b for b in branches if parse_cs_number(b[1]) == cs_number]
        if by_number:
            return by_number

    normalized_title = normalize(title)
    exact, by_street = [], []
    for branch in branches:
        keys = location_keys(branch[1])
        if not keys:
            continue
        if _contains_phrase(normalized_title, keys[0]):
            exact.append(branch)
        elif any(_contains_phrase(normalized_title, key) for key in keys[1:]):
            by_street.append(branch)
    return exact or by_street


def is_same_branch(title, branch_name):
    # Cả hai đều ghi số CS thì số CS quyết định, không xét địa điểm
    title_number, branch_number = parse_cs_number(title), parse_cs_number(branch_name)
    if title_number is not None and branch_number is not None:
        return title_number == branch_number
    return bool(match_branches(title, [(None, branch_name)]))
