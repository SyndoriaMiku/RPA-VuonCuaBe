# RPA-VuonCuaBe
# HƯỚNG DẪN SỬ DỤNG TOOL AUTO ROBOT

Phần mềm này giúp tự động hóa việc lấy dữ liệu (Mã hàng và Số lượng) từ file Excel và điền trực tiếp vào phần mềm quản lý kho KiotViet trên trình duyệt Chrome nhằm mục đích cân hàng nhanh và tự động hơn. Phần mềm chỉ giúp điền, bạn nên kiểm tra lại thông tin cuối cùng trước khi xác nhận.

## 1. CÀI ĐẶT BAN ĐẦU (Chỉ làm 1 lần duy nhất)

Để Tool có thể điều khiển được Chrome, bạn bắt buộc phải tạo một Shortcut Chrome dành riêng cho việc chạy Tool:

1. Ra màn hình Desktop, click chuột phải vào khoảng trống > Chọn **New** > **Shortcut**.
2. Copy và dán chính xác dòng lệnh sau vào ô *Type the location of the item*:
   `"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebug"`
3. Bấm **Next**, đặt tên là **Chrome Debug** rồi bấm **Finish**.

> **Lưu ý:** Nếu bạn có cài các tiện ích chặn quảng cáo (như AdGuard, AdBlock) trên Chrome Debug này, vui lòng **TẮT** chúng đi khi vào trang KiotViet để tránh lỗi hệ thống.

---

## 2. CHUẨN BỊ FILE EXCEL DỮ LIỆU

Tool được lập trình để đọc file Excel theo quy tắc cố định sau:
- **Cột B:** Chứa MÃ HÀNG HÓA.
- **Cột D:** Chứa SỐ LƯỢNG cần thay đổi (có thể là số dương để cộng thêm, hoặc số âm để trừ đi).
- Dữ liệu bắt buộc phải bắt đầu từ **hàng số 4** (tức là ô B4 và D4).
- Tool sẽ chạy liên tục từ trên xuống dưới và tự động DỪNG LẠI khi gặp một ô Mã hàng (cột B) bị trống.

---

## 3. CÁC BƯỚC CHẠY TOOL

**Bước 1:** Mở trình duyệt bằng cái shortcut **Chrome Debug** vừa tạo ở phần 1.
**Bước 2:** Truy cập vào trang KiotViet và mở sẵn màn hình nhập liệu/kiểm kho (nơi có ô tìm kiếm sản phẩm).
**Bước 3:** Mở file chạy của Tool (`auto_kiotviet.exe`).
**Bước 4:** Trên giao diện Tool:
   - Bấm nút **[Chọn File]** và trỏ đến file Excel dữ liệu của bạn.
   - (Tùy chọn) Nhập tên Sheet nếu dữ liệu không nằm ở Sheet đầu tiên.
**Bước 5:** Bấm nút **[▶ CHẠY TỰ ĐỘNG]** và để máy tính tự làm việc.

---

## 4. LƯU Ý QUAN TRỌNG TRONG QUÁ TRÌNH CHẠY

- **Không can thiệp chuột/bàn phím:** Khi Tool đang chạy (đang tự động gõ mã và click chọn hàng), hạn chế tối đa việc di chuột hoặc gõ phím vào cửa sổ Chrome đó để tránh làm sai lệch thao tác của phần mềm.
- **Cơ chế bỏ qua số âm:** Nếu tổng số lượng (Số trên web + Số trong Excel) cho ra kết quả < 0, Tool sẽ tự động bỏ qua không điền lên web để bảo vệ dữ liệu, đồng thời ghi chú lại vào khung Nhật ký (Log) để bạn kiểm tra sau.
- **Theo dõi Log:** Hãy luôn nhìn vào khung "Nhật ký chạy" trên phần mềm để biết tiến độ và phát hiện ngay các mã hàng bị lỗi không tìm thấy.

---
*Chúc bạn thao tác thành công và tiết kiệm được nhiều thời gian!*
