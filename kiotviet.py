import re

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

SEARCH_INPUT_ID = "productSearchInput"
RESULT_ITEM_SELECTOR = ".output-complete ul li"
FIRST_ROW_SELECTOR = "table[role='grid'] tbody tr:nth-child(1)"
CODE_CELL_SELECTOR = "td.cell-code"
QUANTITY_CELL_SELECTOR = "td.cell-quantity.txtR"
QUANTITY_INPUT_SELECTOR = "td.cell-order-number input"
ORDER_CELL_SELECTOR = "td.cell-order"  # Cột STT; KiotViet đánh số giảm dần, dòng mới thêm có STT lớn nhất
NOTE_TEXTAREA_ID = "stockTakeDes"
ADD_STOCKTAKE_SELECTOR = 'a[ng-click="addNew()"]'
BRANCH_SELECT_ID = "cbSelectBranch"
ITEM_ROW_SELECTOR = "table[role='grid'] tbody tr td.cell-order-number"

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


def row_has_code(row, code):
    """Dòng bảng kiểm kho có đúng mã hàng này không (so chính xác ô Mã hàng)."""
    try:
        cell_text = row.find_element(By.CSS_SELECTOR, CODE_CELL_SELECTOR).get_attribute("textContent")
    except Exception:
        # Không thấy ô Mã hàng (giao diện đổi): so cả dòng nhưng vẫn yêu cầu mã trọn vẹn
        return text_contains_code(row.get_attribute("textContent"), code)
    return code_equals(cell_text, code)


def find_matching_row(driver, code):
    """Trả về dòng đầu tiên của bảng kiểm kho nếu nó đúng là mã đang xử lý."""
    try:
        row = driver.find_element(By.CSS_SELECTOR, FIRST_ROW_SELECTOR)
        if row_has_code(row, code):
            return row
    except Exception:
        # Bảng chưa dựng xong hoặc dòng vừa bị render lại: coi như chưa khớp, để chờ tiếp
        pass
    return False


def read_row_order(row):
    """STT của dòng trên phiếu kiểm kho, hoặc None nếu không đọc được."""
    try:
        text = row.find_element(By.CSS_SELECTOR, ORDER_CELL_SELECTOR).get_attribute("textContent").strip()
    except Exception:
        return None
    return text or None


def read_first_row_input(driver):
    try:
        row = driver.find_element(By.CSS_SELECTOR, FIRST_ROW_SELECTOR)
        elm = row.find_element(By.CSS_SELECTOR, QUANTITY_INPUT_SELECTOR)
        return driver.execute_script("return arguments[0].value;", elm)
    except Exception:
        return None


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

    Đang ở danh sách kiểm kho thì bấm nút "Kiểm kho" để tạo phiếu mới.
    Trả về 'ready' (đã ở sẵn màn hình phiếu) hoặc 'created' (vừa tạo phiếu mới).
    """
    state = get_page_state(driver)
    if state == PAGE_EDIT:
        WebDriverWait(driver, timeout).until(_search_input_visible)
        return "ready"
    if state != PAGE_LIST:
        raise RuntimeError("Chrome không ở màn hình Kiểm kho. Hãy mở Hàng hóa > Kiểm kho trên KiotViet.")

    button = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ADD_STOCKTAKE_SELECTOR)))
    driver.execute_script("arguments[0].click();", button)
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: get_page_state(d) == PAGE_EDIT and _search_input_visible(d))
    except TimeoutException:
        raise RuntimeError("Đã bấm nút 'Kiểm kho' nhưng màn hình tạo phiếu không hiện ra.")
    return "created"


def count_item_rows(driver):
    return len(driver.find_elements(By.CSS_SELECTOR, ITEM_ROW_SELECTOR))


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
    web_stt = read_row_order(row)
    qty_cell = row.find_element(By.CSS_SELECTOR, QUANTITY_CELL_SELECTOR)
    qty_web = parse_web_number(qty_cell.text)

    # Bước 5: kiểm tra điều kiện
    total_value = qty_web + qty_excel
    if total_value < 0:
        raise SkipItem(
            f"số lượng cân chỉnh không hợp lệ (Tổng: {format_qty(total_value)}). Sản phẩm đã nằm trên phiếu "
            f"với Thực tế mặc định của KiotViet, cần sửa tay hoặc xóa dòng trước khi lưu", web_stt)

    # Bước 6: điền số lượng bằng API AngularJS
    input_td = row.find_element(By.CSS_SELECTOR, QUANTITY_INPUT_SELECTOR)
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
