import re
import time

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

SEARCH_INPUT_ID = "productSearchInput"
RESULT_ITEM_SELECTOR = ".output-complete ul li"
NOTE_TEXTAREA_ID = "stockTakeDes"
ADD_STOCKTAKE_SELECTOR = 'a[ng-click="addNew()"]'
BRANCH_SELECT_ID = "cbSelectBranch"

# Cột của bảng hàng hóa, tìm theo data-field trên tiêu đề bảng. KiotViet hay đổi class của ô
# (ví dụ ô Mã hàng từ "cell-code" thành "cell-114") nhưng data-field giữ nguyên.
FIELD_STT = "rowNumber"               # STT; đánh số giảm dần, dòng mới thêm có STT lớn nhất
FIELD_CODE = "ProductCode"            # Mã hàng
FIELD_ON_HAND = "SystemCount"         # Tồn kho
FIELD_ACTUAL = "ActualCount"          # Thực tế (ô nhập số)
FIELD_ADJUSTMENT = "AdjustmentValue"  # SL lệch

# Thời gian chờ thêm sau khi màn hình phiếu hiện ra, để bắt hộp thoại "Mở bản nháp" hiện chậm
DRAFT_DIALOG_GRACE_SECONDS = 1.5

# Hàm JS dùng chung để tìm bảng hàng hóa và ô theo data-field (class cũ chỉ là phương án dự phòng)
_GRID_LIB_JS = """
var KV_FALLBACK = {rowNumber: ['td.cell-order', 0], ProductCode: ['td.cell-code', 0],
                   SystemCount: ['td.cell-quantity', 0], ActualCount: ['td.cell-order-number', 0],
                   AdjustmentValue: ['td.cell-quantity', 1]};
function kvItemGridEl() {
    var th = document.querySelector('.k-grid thead th[data-field="ActualCount"]');
    if (th) return th.closest('.k-grid');
    var cell = document.querySelector('table[role=grid] tbody td.cell-order-number');
    return cell ? cell.closest('.k-grid') : null;
}
function kvItemGrid() {
    var el = kvItemGridEl();
    return el && window.jQuery ? jQuery(el).data('kendoGrid') : null;
}
function kvFieldIndex(gridEl, field) {
    var ths = gridEl.querySelectorAll('thead th');
    for (var i = 0; i < ths.length; i++) {
        if (ths[i].getAttribute('data-field') === field) return i;
    }
    return -1;
}
function kvCell(tr, field) {
    var tds = Array.prototype.filter.call(tr.children, function (c) { return c.tagName === 'TD'; });
    var gridEl = tr.closest('.k-grid');
    var i = gridEl ? kvFieldIndex(gridEl, field) : -1;
    if (i >= 0) return i < tds.length ? tds[i] : null;
    var fb = KV_FALLBACK[field];
    return fb ? (tr.querySelectorAll(fb[0])[fb[1]] || null) : null;
}
function kvText(tr, field) {
    var cell = kvCell(tr, field);
    return cell ? cell.textContent.trim() : '';
}
function kvItemRows() {
    var el = kvItemGridEl();
    if (!el) return [];
    return Array.prototype.filter.call(el.querySelectorAll('tbody tr'), function (tr) {
        var cell = kvCell(tr, 'ActualCount');
        return !!(cell && cell.querySelector('input'));
    });
}
"""

_DRAFT_DIALOG_JS = """
var wins = Array.prototype.filter.call(document.querySelectorAll('.k-window'), function (w) {
    return w.offsetParent !== null && /bản nháp/i.test(w.textContent);
});
for (var i = 0; i < wins.length; i++) {
    var btn = wins[i].querySelector('.kv-btn-cancel');
    if (btn && btn.offsetParent !== null) return btn;
}
return null;
"""

PAGE_LIST = "list"      # Danh sách phiếu kiểm kho: #/StockTakes
PAGE_EDIT = "edit"      # Đang tạo/sửa phiếu: #/StockTakes/0, #/StockTakes/<id>
PAGE_OTHER = "other"


class ItemError(Exception):
    """Lỗi của một mã hàng; web_stt là STT của dòng trên phiếu KiotViet nếu sản phẩm đã được thêm vào."""

    def __init__(self, message, web_stt=None):
        super().__init__(message)
        self.web_stt = web_stt


class SkipItem(ItemError):
    """Mã hàng bị bỏ qua vì lý do nghiệp vụ (không tìm thấy, số lượng âm, ...)."""


class VerifyError(ItemError):
    """Đã điền số lượng nhưng giá trị đọc lại trên web không khớp."""


def format_qty(value):
    value = round(float(value), 6)
    return str(int(value)) if value.is_integer() else str(value)


def parse_web_number(text):
    cleaned = str(text).replace(",", "").replace(" ", "").strip()
    return float(cleaned) if cleaned else 0.0


def code_equals(text, code):
    return str(text or "").strip().lower() == code.strip().lower()


def text_contains_code(text, code):
    """Mã xuất hiện như một cụm trọn vẹn trong text.

    Không được dính liền với ký tự hay gặp trong mã hàng (chữ, số, _ . - /), để mã
    "7350107133709" không khớp nhầm với "7350107133709_D26.10".
    """
    pattern = r"(?<![\w.\-/])" + re.escape(code.strip()) + r"(?![\w.\-/])"
    return re.search(pattern, str(text or ""), re.IGNORECASE) is not None


def connect_driver(port):
    """Gắn Selenium vào Chrome debug đang mở và chuyển sang tab KiotViet nếu có."""
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    driver = webdriver.Chrome(options=chrome_options)

    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if "kiotviet.vn" in driver.current_url.lower():
            break
    bring_to_front(driver)
    return driver


def bring_to_front(driver):
    """Mở lại cửa sổ nếu bị thu nhỏ và đưa tab lên trước.

    Khi cửa sổ Chrome bị thu nhỏ hoặc che, trang ở trạng thái 'hidden' và phím gõ
    không vào được ô tìm kiếm.
    """
    try:
        window = driver.execute_cdp_cmd("Browser.getWindowForTarget", {})
        if window["bounds"].get("windowState") == "minimized":
            driver.execute_cdp_cmd("Browser.setWindowBounds",
                                   {"windowId": window["windowId"], "bounds": {"windowState": "normal"}})
        driver.execute_cdp_cmd("Page.bringToFront", {})
    except Exception:
        pass  # Không đưa lên trước được thì vẫn thử chạy; bước gõ mã sẽ báo lỗi nếu không gõ được


_EXACT_RESULT_JS = """
var code = arguments[1].trim().toLowerCase();
var items = document.querySelectorAll(arguments[0]);
for (var i = 0; i < items.length; i++) {
    var nodes = [items[i]].concat(Array.prototype.slice.call(items[i].querySelectorAll('*')));
    for (var j = 0; j < nodes.length; j++) {
        // Chỉ lấy chữ nằm trực tiếp trong phần tử, không gộp chữ của phần tử con
        var own = Array.prototype.filter.call(nodes[j].childNodes, function (n) { return n.nodeType === 3; })
            .map(function (n) { return n.textContent; }).join('').trim().toLowerCase();
        if (own === code) return items[i];
    }
}
return null;
"""


def find_matching_result(driver, code):
    """Trả về gợi ý tìm kiếm đúng mã đang tìm, hoặc False nếu chưa có.

    Ưu tiên gợi ý có một phần tử hiển thị đúng bằng mã; sau đó mới đến gợi ý chứa mã
    như một cụm trọn vẹn. Không bao giờ chọn mã khác chỉ vì có chung phần đầu.
    """
    try:
        exact = driver.execute_script(_EXACT_RESULT_JS, RESULT_ITEM_SELECTOR, code)
    except Exception:
        exact = None
    if exact:
        return exact

    for li in driver.find_elements(By.CSS_SELECTOR, RESULT_ITEM_SELECTOR):
        try:
            if text_contains_code(li.get_attribute("textContent"), code):
                return li
        except StaleElementReferenceException:
            continue
    return False


def get_cell(driver, row, field):
    """Ô của cột `field` (data-field) trong một dòng của bảng hàng hóa, hoặc None."""
    return driver.execute_script(_GRID_LIB_JS + "return kvCell(arguments[0], arguments[1]);", row, field)


def cell_text(driver, row, field):
    return driver.execute_script(_GRID_LIB_JS + "return kvText(arguments[0], arguments[1]);", row, field)


def first_item_row(driver):
    """Dòng đầu tiên của bảng hàng hóa (sản phẩm vừa thêm), hoặc None."""
    return driver.execute_script(_GRID_LIB_JS + "return kvItemRows()[0] || null;")


def row_has_code(driver, row, code):
    """Dòng bảng kiểm kho có đúng mã hàng này không (so chính xác ô Mã hàng)."""
    cell = get_cell(driver, row, FIELD_CODE)
    if cell is not None:
        return code_equals(cell.get_attribute("textContent"), code)
    # Không xác định được cột Mã hàng: chỉ chấp nhận ô có chữ đúng bằng mã, không so gần đúng
    return any(code_equals(td.get_attribute("textContent"), code)
               for td in row.find_elements(By.CSS_SELECTOR, "td"))


def find_matching_row(driver, code):
    """Trả về dòng đầu tiên của bảng kiểm kho nếu nó đúng là mã đang xử lý."""
    try:
        row = first_item_row(driver)
        if row is not None and row_has_code(driver, row, code):
            return row
    except Exception:
        # Bảng chưa dựng xong hoặc dòng vừa bị render lại: coi như chưa khớp, để chờ tiếp
        pass
    return False


def read_row_order(driver, row):
    """STT của dòng trên phiếu kiểm kho, hoặc None nếu không đọc được."""
    try:
        return cell_text(driver, row, FIELD_STT) or None
    except Exception:
        return None


def read_first_row_input(driver):
    try:
        return driver.execute_script(_GRID_LIB_JS + """
            var tr = kvItemRows()[0];
            var input = tr && kvCell(tr, 'ActualCount').querySelector('input');
            return input ? input.value : null;
        """)
    except Exception:
        return None


def dismiss_draft_dialog(driver):
    """Bấm "Bỏ qua" ở hộp thoại "Mở bản nháp" nếu đang hiện; KiotViet sẽ xóa bản nháp chưa lưu.

    Trả về True nếu đã bấm.
    """
    button = driver.execute_script(_DRAFT_DIALOG_JS)
    if not button:
        return False
    driver.execute_script("arguments[0].click();", button)
    WebDriverWait(driver, 5, poll_frequency=0.2).until(lambda d: not d.execute_script(_DRAFT_DIALOG_JS))
    return True


def page_state_from_url(url):
    fragment = str(url).split("#", 1)[1] if "#" in str(url) else ""
    path = fragment.split("?", 1)[0].rstrip("/").lower()
    if path == "/stocktakes":
        return PAGE_LIST
    if path.startswith("/stocktakes/"):
        return PAGE_EDIT
    return PAGE_OTHER


def get_page_state(driver):
    return page_state_from_url(driver.current_url)


def _search_input_visible(driver):
    return driver.execute_script(
        "var s = document.getElementById(arguments[0]); return !!s && s.offsetParent !== null;",
        SEARCH_INPUT_ID)


def ensure_stocktake_page(driver, timeout=15):
    """Đảm bảo đang ở màn hình nhập phiếu kiểm kho.

    Đang ở danh sách kiểm kho thì bấm nút "Kiểm kho" để tạo phiếu mới. Nếu KiotViet hỏi
    "Mở bản nháp" (còn phiếu nháp chưa lưu) thì bấm "Bỏ qua" để bắt đầu phiếu mới.
    Trả về (trạng thái, đã_bỏ_bản_nháp); trạng thái là 'ready' (đã ở sẵn màn hình phiếu)
    hoặc 'created' (vừa tạo phiếu mới).
    """
    state = get_page_state(driver)
    if state == PAGE_EDIT:
        discarded = dismiss_draft_dialog(driver)
        WebDriverWait(driver, timeout).until(_search_input_visible)
        return "ready", discarded
    if state != PAGE_LIST:
        raise RuntimeError("Chrome không ở màn hình Kiểm kho. Hãy mở Hàng hóa > Kiểm kho trên KiotViet.")

    button = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ADD_STOCKTAKE_SELECTOR)))
    driver.execute_script("arguments[0].click();", button)

    discarded = False
    ready_since = None
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if dismiss_draft_dialog(driver):
            discarded = True
            ready_since = None
            continue
        if get_page_state(driver) == PAGE_EDIT and _search_input_visible(driver):
            ready_since = ready_since or time.monotonic()
            if time.monotonic() - ready_since >= DRAFT_DIALOG_GRACE_SECONDS:
                return "created", discarded
        else:
            ready_since = None
        time.sleep(0.2)
    raise RuntimeError("Đã bấm nút 'Kiểm kho' nhưng màn hình tạo phiếu không hiện ra.")


def count_item_rows(driver):
    return driver.execute_script(_GRID_LIB_JS + "return kvItemRows().length;")


_BRANCH_JS = "var dd = window.jQuery && jQuery('#' + arguments[0]).data('kendoDropDownList');"


def get_current_branch(driver):
    """(id, tên) của cơ sở đang làm việc, hoặc None nếu chưa đọc được."""
    result = driver.execute_script(
        _BRANCH_JS + "return dd ? [String(dd.value()), dd.text()] : null;", BRANCH_SELECT_ID)
    return tuple(result) if result else None


def get_branches(driver):
    return [
        (str(value), text)
        for value, text in driver.execute_script(
            "var s = document.getElementById(arguments[0]);"
            "return s ? Array.from(s.options).map(o => [o.value, o.text]) : [];",
            BRANCH_SELECT_ID)
    ]


def change_branch(driver, branch_id, timeout=45):
    """Đổi cơ sở làm việc qua Kendo DropDownList (giống người dùng chọn trong danh sách).

    Sau khi đổi, KiotViet tải lại toàn bộ trang (khoảng 10 giây). Đánh dấu trang cũ bằng
    window.__rpaBranchChange rồi chờ đến khi trang mới thay thế và hiển thị đúng cơ sở mới.
    """
    driver.execute_script(
        _BRANCH_JS + "if (!dd) throw new Error('Không tìm thấy ô chọn cơ sở');"
        "window.__rpaBranchChange = true; dd.value(arguments[1]); dd.trigger('change');",
        BRANCH_SELECT_ID, str(branch_id))

    def reloaded_with_new_branch(d):
        try:
            if d.execute_script("return !!window.__rpaBranchChange || document.readyState !== 'complete';"):
                return False
            current = get_current_branch(d)
        except Exception:
            return False
        return current if current and current[0] == str(branch_id) else False

    try:
        current = WebDriverWait(driver, timeout, poll_frequency=0.5).until(reloaded_with_new_branch)
    except TimeoutException:
        raise RuntimeError("Đã chọn cơ sở mới nhưng trang KiotViet chưa tải lại xong.")

    if get_page_state(driver) == PAGE_EDIT:
        WebDriverWait(driver, 15).until(_search_input_visible)
    return current


_ITEM_GRID_JS = _GRID_LIB_JS + "var grid = kvItemGrid();\n"


def _grid_state(driver):
    return driver.execute_script(
        _ITEM_GRID_JS + "if (!grid) return null; var ds = grid.dataSource;"
        "return {page: ds.page(), pageSize: ds.pageSize(), totalPages: ds.totalPages(), total: ds.total()};")


def _read_visible_rows(driver):
    return driver.execute_script(_GRID_LIB_JS + """
        return kvItemRows().map(function (tr) {
            return {stt: kvText(tr, 'rowNumber'), code: kvText(tr, 'ProductCode'),
                    adjustment: kvText(tr, 'AdjustmentValue')};
        });
    """)


def read_stocktake_rows(driver, timeout=10):
    """Đọc toàn bộ dòng của phiếu kiểm kho: [{stt, code, adjustment (SL lệch)}].

    KiotViet chỉ hiển thị 15 dòng mỗi trang, nên chuyển lần lượt từng trang để đọc đúng
    những gì đang hiển thị, xong thì quay lại trang ban đầu.
    """
    state = _grid_state(driver)
    if not state:
        raise RuntimeError("Không tìm thấy bảng hàng hóa của phiếu kiểm kho.")
    if not state["total"]:
        return []

    original_page = state["page"]
    page_size = state["pageSize"] or state["total"]
    rows = []
    previous_page = None
    try:
        for page in range(1, state["totalPages"] + 1):
            driver.execute_script(_ITEM_GRID_JS + "grid.dataSource.page(arguments[0]);", page)
            expected = min(page_size, state["total"] - (page - 1) * page_size)

            def page_ready(d):
                visible = _read_visible_rows(d)
                ok = (len(visible) == expected
                      and all(r["code"] and r["adjustment"] for r in visible)
                      and (previous_page is None or visible != previous_page))  # bảng đã chuyển sang trang mới
                return visible if ok else False

            try:
                previous_page = WebDriverWait(driver, timeout, poll_frequency=0.2).until(page_ready)
                rows.extend(previous_page)
            except TimeoutException:
                raise RuntimeError(f"Trang {page} của phiếu kiểm kho không hiển thị đủ {expected} dòng.")
    finally:
        try:
            driver.execute_script(_ITEM_GRID_JS + "if (grid) grid.dataSource.page(arguments[0]);", original_page)
        except Exception:
            pass

    if len(rows) != state["total"]:
        raise RuntimeError(f"Đọc được {len(rows)} dòng nhưng phiếu có {state['total']} mã hàng.")
    for row in rows:
        row["adjustment"] = parse_web_number(row["adjustment"])
    return rows


def merge_note(current, note_text):
    """Ghép ghi chú mới vào nội dung sẵn có; trả về None nếu đã có đúng nội dung đó."""
    current = (current or "").strip()
    if note_text.lower() in current.lower():
        return None
    return f"{current}\n{note_text}" if current else note_text


def fill_note(driver, note_text):
    """Điền ghi chú của phiếu kiểm kho. Giữ lại ghi chú cũ nếu có. Trả về 'filled' hoặc 'exists'."""
    textarea = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, NOTE_TEXTAREA_ID)))
    if not textarea.is_enabled():
        raise RuntimeError("ô Ghi chú đang bị khóa (phiếu đã hoàn thành?).")

    new_value = merge_note(textarea.get_attribute("value"), note_text)
    if new_value is None:
        return "exists"

    # Cập nhật qua AngularJS để ng-model nhận giá trị, blur để KiotViet gọi updateLocal()
    driver.execute_script("""
        var $el = angular.element(arguments[0]);
        $el.val(arguments[1]).triggerHandler('input');
        $el.triggerHandler('change');
        $el.triggerHandler('blur');
    """, textarea, new_value)

    try:
        WebDriverWait(driver, 3).until(
            lambda d: d.find_element(By.ID, NOTE_TEXTAREA_ID).get_attribute("value") == new_value)
    except TimeoutException:
        raise VerifyError("đã điền ghi chú nhưng đọc lại không khớp.")
    return "filled"


def process_item(driver, code, qty_excel, result_timeout):
    wait = WebDriverWait(driver, 10)

    # Bước 1: nhập mã vào ô tìm kiếm
    for attempt in range(2):
        search_input = wait.until(EC.element_to_be_clickable((By.ID, SEARCH_INPUT_ID)))
        search_input.click()
        search_input.send_keys(Keys.CONTROL + "a")
        search_input.send_keys(Keys.DELETE)
        search_input.send_keys(code)
        if search_input.get_attribute("value").strip() == code.strip():
            break
        bring_to_front(driver)
    else:
        raise RuntimeError("Không gõ được mã vào ô tìm kiếm. Hãy mở cửa sổ Chrome Debug, đừng thu nhỏ.")

    # Bước 2: chờ gợi ý chứa đúng mã rồi click (JS để né AdGuard)
    try:
        result = WebDriverWait(driver, result_timeout).until(
            lambda d: find_matching_result(d, code)
        )
    except TimeoutException:
        raise SkipItem("Không tìm thấy sản phẩm trên KiotViet.")
    driver.execute_script("arguments[0].click();", result)

    # Bước 3: chờ dòng đầu bảng chính là sản phẩm vừa chọn
    try:
        row = WebDriverWait(driver, 10).until(lambda d: find_matching_row(d, code))
    except TimeoutException:
        raise SkipItem("Không thấy sản phẩm xuất hiện ở dòng đầu bảng kiểm kho.")

    # Bước 4: đọc STT trên phiếu (để báo lỗi dễ tìm lại) và số lượng hiện có trên web
    web_stt = read_row_order(driver, row)
    qty_web = parse_web_number(cell_text(driver, row, FIELD_ON_HAND))

    # Bước 5: kiểm tra điều kiện
    total_value = qty_web + qty_excel
    if total_value < 0:
        raise SkipItem(f"số lượng cân chỉnh không hợp lệ (Tổng: {format_qty(total_value)})", web_stt)

    # Bước 6: điền số lượng bằng API AngularJS
    input_td = get_cell(driver, row, FIELD_ACTUAL).find_element(By.CSS_SELECTOR, "input")
    driver.execute_script("""
        var elm = arguments[0];
        var val = arguments[1];
        var $el = angular.element(elm);
        $el.val(val).triggerHandler('input');
        $el.triggerHandler('change');
        $el.triggerHandler('blur');
    """, input_td, total_value)

    # Bước 7: đọc lại giá trị để xác minh đã điền đúng
    def value_matches(d):
        raw = read_first_row_input(d)
        if raw is None or not str(raw).strip():
            return False
        try:
            return abs(parse_web_number(raw) - total_value) < 1e-6
        except ValueError:
            return False

    try:
        WebDriverWait(driver, 3).until(value_matches)
    except TimeoutException:
        raise VerifyError(
            f"đã điền nhưng đọc lại không khớp (mong đợi {format_qty(total_value)}, "
            f"thực tế {read_first_row_input(driver)!r})",
            web_stt,
        )

    return qty_web, total_value
