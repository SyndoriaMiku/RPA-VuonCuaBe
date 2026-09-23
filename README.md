# RPA Vườn Của Bé - Auto Robot KiotViet

Ứng dụng desktop hỗ trợ đọc mã hàng và số lượng cân chỉnh từ Excel, sau đó tự động nhập dữ liệu vào màn hình kiểm kho của KiotViet bằng Chrome.

> Tool chỉ hỗ trợ nhập dữ liệu. Hãy kiểm tra lại toàn bộ kết quả trên KiotViet trước khi xác nhận hoặc lưu phiếu.

## Yêu cầu

- Windows.
- Google Chrome được cài tại đường dẫn mặc định:
  `C:\Program Files\Google\Chrome\Application\chrome.exe`.
- File Excel định dạng `.xlsx`.
- Tài khoản KiotViet có quyền thao tác tại màn hình kiểm kho/cân chỉnh hàng hóa.

## Chuẩn bị file Excel

Tool đọc dữ liệu theo cấu trúc cố định:

- Ô `C2`: tiêu đề hiển thị trên cửa sổ Tool (không bắt buộc).
- Ô bắt đầu mặc định là `B4`; số lượng tương ứng được đọc tại `D4`.
- Có thể đổi ô bắt đầu, ví dụ thành `B18`, để bỏ qua các dòng lỗi hoặc dòng trống phía trên. Khi đó Tool sẽ đọc mã hàng từ `B18` và số lượng từ `D18`.
- Nếu chọn ô ở cột khác, cột số lượng luôn được tính cách cột mã hàng 2 cột; ví dụ mã hàng ở `C10` thì số lượng nằm tại `E10`.
- Từ ô bắt đầu, Tool quét đến cuối sheet và chỉ xử lý những dòng có mã hàng. Mọi ô trống xen giữa được tự động bỏ qua và không làm Tool dừng lại.
- Sheet mặc định là `Cân Chỉnh Hàng Hóa`. Có thể nhập tên sheet khác hoặc để trống để dùng sheet đang active trong workbook.

## Cách sử dụng

1. Mở `AutoApp.exe`.
2. Nhấn **Mở Chrome Debug**. Tool sẽ mở Chrome với cổng debug `9999` và profile riêng tại `C:\ChromeDebug`.
3. Trong cửa sổ Chrome vừa mở, đăng nhập KiotViet và truy cập màn hình kiểm kho/cân chỉnh hàng hóa có ô tìm kiếm sản phẩm.
4. Quay lại Tool, nhấn **Chọn File** và chọn file Excel `.xlsx`.
5. Sau khi load file, Tool chỉ ghi vào khung nhật ký: tên file, tên sheet, các khoảng dòng sẽ được xử lý và tổng số mã hàng. Tool không hiển thị popup cho thông tin này.
6. Kiểm tra tên sheet và ô bắt đầu. Giữ `B4` với file bình thường, hoặc nhập một ô thấp hơn như `B18` để bỏ qua các dòng phía trên.
7. Nếu cần, điều chỉnh thời gian chờ tải kết quả tìm kiếm.
8. Nhấn **▶ CHẠY TỰ ĐỘNG**.
9. Theo dõi khung **Nhật ký chạy** và kiểm tra lại dữ liệu trên KiotViet sau khi Tool hoàn tất.

### Mở Chrome Debug thủ công

Nếu không dùng nút **Mở Chrome Debug**, có thể tạo shortcut Chrome với lệnh sau:

```text
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9999 --user-data-dir="C:\ChromeDebug"
```

Cổng phải là `9999` để khớp với cấu hình kết nối trong Tool. Nên tắt các tiện ích chặn quảng cáo trên profile Chrome Debug nếu chúng làm gián đoạn thao tác ở KiotViet.

## Quy tắc xử lý

- Tool lấy số lượng hiện tại trên KiotViet và cộng với số lượng cân chỉnh trong Excel.
- Nếu tổng nhỏ hơn 0, mã hàng đó sẽ bị bỏ qua và được ghi vào nhật ký.
- Nếu không tìm thấy sản phẩm hoặc thao tác web gặp lỗi, Tool bỏ qua mã hiện tại, ghi lỗi và tiếp tục với mã tiếp theo.
- Không nên sử dụng chuột hoặc bàn phím trong cửa sổ Chrome khi Tool đang chạy.

## Môi trường phát triển

Dự án sử dụng môi trường Conda `p310` trên máy:

```powershell
conda activate p310
python --version
python app.py
```

Phiên bản Python hiện tại của môi trường là Python 3.10.18.

Các thư viện bên thứ ba cần thiết được khai báo trong `requirements.txt`:

```powershell
python -m pip install -r requirements.txt
```

Để đóng gói ứng dụng thành file `.exe`, chạy file spec:

```powershell
pyinstaller AutoApp.spec
```

`tkinter`, `time`, `threading`, `os` và `subprocess` thuộc Python/standard library nên không cần cài qua `pip`.

## Cấu trúc dự án

- `app.py`: giao diện và toàn bộ logic tự động hóa.
- `requirements.txt`: danh sách thư viện Python cần cài đặt.
- `AutoApp.spec`: cấu hình đóng gói bằng PyInstaller, bao gồm các DLL Tcl/Tk cần thiết từ môi trường Conda.
- `dist/AutoApp.exe`: bản executable sau khi build (đang được `.gitignore` bỏ qua).

## Lưu ý kỹ thuật

Tool phụ thuộc vào cấu trúc giao diện HTML hiện tại của KiotViet. Nếu KiotViet thay đổi ID, CSS selector hoặc cơ chế nhập liệu, phần tự động hóa có thể cần được cập nhật.
