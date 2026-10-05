# RPA Vườn Của Bé - Auto Robot KiotViet

Ứng dụng desktop hỗ trợ đọc mã hàng và số lượng cân chỉnh từ Excel, sau đó tự động nhập dữ liệu vào màn hình kiểm kho của KiotViet bằng Chrome.

> Tool chỉ hỗ trợ nhập dữ liệu. Hãy kiểm tra lại toàn bộ kết quả trên KiotViet trước khi xác nhận hoặc lưu phiếu.

## Yêu cầu

- Windows.
- Google Chrome (Tool tự dò các vị trí cài đặt phổ biến, hoặc cho phép chọn `chrome.exe` thủ công).
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

Giao diện chia thành 3 bước:

1. **① Mở Chrome Debug**: Tool tự dò Chrome trong `Program Files`, `Program Files (x86)` và `%LOCALAPPDATA%`. Nếu không thấy, Tool hỏi bạn chọn `chrome.exe` và ghi nhớ đường dẫn đó. Chrome mở với cổng debug `9999` và profile riêng `C:\ChromeDebug`. Đăng nhập KiotViet và mở **Hàng hóa > Kiểm kho**: có thể để ở trang danh sách phiếu (`#/StockTakes`) hoặc đã mở sẵn phiếu (`#/StockTakes/0`).
2. **② Chọn dữ liệu Excel**: chọn file `.xlsx`, tên sheet và ô bắt đầu. Dòng gợi ý bên cạnh cho biết Tool sẽ đọc mã hàng và số lượng ở đâu. Sau khi chọn file, tiêu đề lấy từ ô `C2`, và nhật ký ghi tên file, sheet, các khoảng dòng sẽ xử lý.
   - **Tự đổi cơ sở**: nếu Chrome Debug đang mở, Tool dò cơ sở theo tiêu đề `C2` và đổi cơ sở làm việc trên KiotViet cho khớp (trang tải lại khoảng 10 giây). Tool không tự đổi nếu tiêu đề khớp nhiều cơ sở, không khớp cơ sở nào, hoặc phiếu kiểm kho đang mở đã có hàng (để tránh mất dữ liệu); khi đó chỉ ghi cảnh báo vào nhật ký.
3. **③ Kiểm tra và chạy**:
   - **Kiểm tra file**: chỉ đọc Excel, không thao tác gì trên KiotViet. Báo mã trùng, số lượng không phải số và các dòng có số lượng bằng 0.
   - **▶ CHẠY TỰ ĐỘNG**: nếu file có mã trùng hoặc số lượng không hợp lệ, Tool hỏi xác nhận trước khi chạy. Thanh tiến độ hiển thị `đã xử lý/tổng`.
     - Nếu cơ sở trong tiêu đề file không khớp cơ sở đang làm việc trên web, Tool hỏi có tiếp tục không (vì tên trong file đôi khi khác tên trên web).
     - Nếu web đang ở trang danh sách kiểm kho, Tool tự bấm nút **Kiểm kho** để tạo phiếu mới; nếu đã ở màn hình phiếu thì làm luôn.
   - **■ DỪNG**: Tool dừng sau khi xong mã hàng hiện tại.
   - **Chờ kết quả tối đa**: Tool chỉ chờ đến khi thấy sản phẩm, không chờ đủ số giây cố định.

Nhật ký có màu: xanh là thành công, cam là bỏ qua/cảnh báo, đỏ là lỗi hoặc cần kiểm tra. Có nút **Sao chép tất cả** và **Xóa**. Cuối mỗi lần chạy có thống kê thành công / bỏ qua / lỗi.

Các thiết lập (file, sheet, ô bắt đầu, thời gian chờ, đường dẫn Chrome) được lưu vào `config.json` cạnh `AutoApp.exe` và tự nạp lại lần sau. Có thể sửa `debug_port` và `user_data_dir` trong file này nếu cần.

### Mở Chrome Debug thủ công

Nếu không dùng nút **Mở Chrome Debug**, có thể tạo shortcut Chrome với lệnh sau:

```text
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9999 --user-data-dir="C:\ChromeDebug"
```

Cổng phải khớp với `debug_port` trong `config.json` (mặc định `9999`). Nên tắt các tiện ích chặn quảng cáo trên profile Chrome Debug nếu chúng làm gián đoạn thao tác ở KiotViet.

## Quy tắc xử lý

- Khớp cơ sở theo tiêu đề `C2`, ưu tiên số CS (`CS8`, `CS08`, `cs 08` đều là CS08), sau đó đến địa điểm: địa chỉ đầy đủ (`22 Vũ Trọng Phụng`) hoặc tên đường (`Vũ Trọng Phụng`). Không phân biệt hoa/thường và dấu.

- Tool lấy số lượng hiện tại trên KiotViet và cộng với số lượng cân chỉnh trong Excel.
- Nếu tổng nhỏ hơn 0, mã hàng đó sẽ bị bỏ qua và được ghi vào nhật ký.
- Ô số lượng trống được tính là 0; ô không phải số thì mã đó bị bỏ qua.
- Mã xuất hiện nhiều lần trong Excel sẽ được cộng dồn lên web (Tool cảnh báo trước khi chạy).
- Tool chỉ chọn gợi ý tìm kiếm có chứa đúng mã hàng, và chỉ thao tác khi dòng đầu bảng kiểm kho là đúng sản phẩm đó.
- Sau khi điền, Tool đọc lại ô số lượng để xác minh; nếu không khớp sẽ ghi `CẦN KIỂM TRA` vào nhật ký.
- Nếu không tìm thấy sản phẩm hoặc thao tác web gặp lỗi, Tool bỏ qua mã hiện tại, ghi lỗi và tiếp tục với mã tiếp theo.
- Sau khi chạy xong toàn bộ danh sách (không bấm Dừng), Tool điền "Cân hàng lệch" vào ô **Ghi chú** của phiếu. Ghi chú cũ được giữ lại, nội dung mới thêm vào dòng dưới; nếu đã có sẵn thì không điền lại. Có thể đổi nội dung bằng khóa `note_text` trong `config.json` (để trống để tắt).
- Không nên sử dụng chuột hoặc bàn phím trong cửa sổ Chrome khi Tool đang chạy.

## Môi trường phát triển

Dự án sử dụng môi trường Conda `p310` (Python 3.10.18):

```powershell
conda activate p310
python -m pip install -r requirements-dev.txt
python app.py
python -m pytest tests
```

- `requirements.txt`: thư viện cần để chạy Tool (`openpyxl`, `selenium`).
- `requirements-dev.txt`: thêm `pyinstaller` và `pytest` cho phát triển và build.

Selenium (từ 4.6) tự tải chromedriver phù hợp với phiên bản Chrome qua Selenium Manager, không cần tải chromedriver thủ công.

### Build

```powershell
.uild.ps1
```

Script chạy test trong `p310`, sau đó build `dist/AutoApp.exe` bằng `AutoApp.spec`. Nếu test lỗi thì không build.

## Cấu trúc dự án

- `app.py`: điểm khởi chạy.
- `gui.py`: giao diện Tkinter, xử lý luồng và lưu cấu hình.
- `automation.py`: vòng lặp xử lý từng mã hàng, thống kê kết quả, xử lý dừng.
- `kiotviet.py`: thao tác Selenium với KiotViet (nhận diện trang, đổi cơ sở, tạo phiếu, tìm, điền, xác minh).
- `branch_match.py`: khớp tiêu đề file với cơ sở trên web (theo số CS hoặc địa điểm, không phân biệt dấu).
- `excel_reader.py`: đọc Excel, phân tích ô bắt đầu, kiểm tra mã trùng/số lượng.
- `config.py`: đọc/ghi `config.json`, dò đường dẫn Chrome.
- `tests/`: test `pytest` cho các module trên (không cần Chrome hay KiotViet).
- `AutoApp.spec`: cấu hình đóng gói bằng PyInstaller, bao gồm các DLL Tcl/Tk cần thiết từ môi trường Conda.
- `build.ps1`: script test và build.
- `dist/AutoApp.exe`: bản executable sau khi build (đang được `.gitignore` bỏ qua).

## Lưu ý kỹ thuật

Tool phụ thuộc vào cấu trúc giao diện HTML hiện tại của KiotViet. Các selector nằm ở đầu file `kiotviet.py`; nếu KiotViet thay đổi ID, CSS selector hoặc cơ chế nhập liệu, chỉ cần cập nhật file này.
