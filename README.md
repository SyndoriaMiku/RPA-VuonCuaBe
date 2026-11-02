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
   - **Cơ sở theo file**: danh sách cơ sở đọc từ web, tự chọn sẵn cơ sở khớp với tiêu đề `C2` (theo số CS hoặc địa điểm). Nếu chọn tay, Tool nhớ lựa chọn đó theo tiêu đề file cho lần sau. Dòng bên phải báo `✓ Khớp với web`, `⚠ Web đang ở ...`, hoặc `⚠ Không nhận ra cơ sở từ C2, hãy chọn`. Bấm **Đổi trên web** để đổi cơ sở KiotViet sang mục đang chọn.
   - **Tự đổi cơ sở**: sau khi chọn file, nếu Chrome Debug đang mở, Tool tự đổi cơ sở làm việc trên KiotViet cho khớp (trang tải lại khoảng 10 giây). Tool không tự đổi nếu tiêu đề khớp nhiều cơ sở, không khớp cơ sở nào, hoặc phiếu kiểm kho đang mở đã có hàng (để tránh mất dữ liệu); khi đó chỉ ghi cảnh báo và bạn có thể chọn tay rồi bấm **Đổi trên web**.
3. **③ Kiểm tra và chạy**:
   - **Kiểm tra file**: chỉ đọc Excel, không thao tác gì trên KiotViet. Báo mã trùng, số lượng không phải số và các dòng có số lượng bằng 0.
   - **▶ CHẠY TỰ ĐỘNG**: nếu file có mã trùng hoặc số lượng không hợp lệ, Tool hỏi xác nhận trước khi chạy. Thanh tiến độ hiển thị `đã xử lý/tổng`.
     - Nếu cơ sở trong tiêu đề file không khớp cơ sở đang làm việc trên web, Tool hỏi có tiếp tục không (vì tên trong file đôi khi khác tên trên web).
     - Nếu web đang ở trang danh sách kiểm kho, Tool tự bấm nút **Kiểm kho** để tạo phiếu mới; nếu đã ở màn hình phiếu thì làm luôn.
     - Nếu KiotViet hỏi **"Mở bản nháp"** (còn phiếu nháp chưa lưu), Tool tự chọn **Bỏ qua** để nhập file vào phiếu mới và ghi cảnh báo vào nhật ký. Lưu ý: Bỏ qua sẽ **xóa bản nháp cũ**, nên hãy lưu phiếu đang làm dở trước khi chạy Tool.
   - **■ DỪNG**: Tool dừng sau khi xong mã hàng hiện tại.
   - **Chờ kết quả tối đa**: Tool chỉ chờ đến khi thấy sản phẩm, không chờ đủ số giây cố định.

Nhật ký có màu: xanh là thành công, cam là bỏ qua/cảnh báo, đỏ là lỗi hoặc cần kiểm tra. Có nút **Sao chép tất cả** và **Xóa**. Cuối mỗi lần chạy có thống kê thành công / bỏ qua / lỗi.

Mã hàng trong các dòng `BỎ QUA`, `CẦN KIỂM TRA`, `LỖI Web`, `Lệch` và `Thiếu` được gạch chân màu xanh. **Bấm vào mã** để Tool mở Chrome, chuyển phiếu kiểm kho tới đúng trang chứa mã đó, cuộn tới và tô sáng dòng trong vài giây. Không tìm thấy thì bảng được trả về trang ban đầu và nhật ký báo `Không thấy mã ... trên phiếu đang mở`. Tool không nhảy trang khi đang chạy; hãy chờ chạy xong rồi bấm lại.

Khi lần chạy kết thúc (hoàn tất, dừng hoặc lỗi kết nối), Tool phát âm thanh hệ thống của Windows và nhấp nháy biểu tượng trên taskbar cho đến khi bạn mở lại cửa sổ: âm "OK" khi mọi mã đều thành công, âm cảnh báo khi có mã bỏ qua/sai lệch/lỗi. Có thể tắt hai thông báo này trong **Cài đặt**.

Các thiết lập (file, sheet, ô bắt đầu, thời gian chờ, đường dẫn Chrome) được lưu vào `config.json` cạnh `AutoApp.exe` và tự nạp lại lần sau. Các thiết lập còn lại sửa trong **⚙ Cài đặt** (xem bên dưới).

## Cài đặt

Bấm **⚙ Cài đặt** ở góc trên bên phải cửa sổ để sửa các thiết lập mà không cần mở `config.json`:

| Nhóm | Thiết lập | Khóa trong `config.json` |
|------|-----------|--------------------------|
| Chrome | Đường dẫn `chrome.exe` (Chọn… / Tự dò), cổng debug, thư mục profile | `chrome_path`, `debug_port`, `user_data_dir` |
| Chạy | Chờ kết quả tối đa (1–60 giây), nội dung ghi chú sau khi chạy (tối đa 2000 ký tự) | `wait_seconds`, `note_text` |
| Thông báo | Âm thanh khi xong, nhấp nháy taskbar | `notify_sound`, `notify_flash` |
| Trạng thái | Tần suất cập nhật thanh trạng thái (3–60 giây) | `status_interval` |

Lỗi hợp lệ hiện ngay cạnh ô nhập; không lưu được khi còn lỗi. Nút **Khôi phục mặc định** đưa các ô về giá trị gốc (chưa lưu cho đến khi bấm **Lưu**). Không mở được Cài đặt khi Tool đang chạy.

## Thanh trạng thái

Dải ngay dưới tiêu đề cho biết trước khi bấm Chạy là Tool có thao tác được không. Nó tự cập nhật mỗi 5 giây (đổi trong **Cài đặt**):

- ● xám **Chưa kết nối Chrome Debug**: kèm nút **Mở Chrome Debug**.
- ● cam **Chưa mở KiotViet** hoặc **Chưa ở màn hình Kiểm kho**.
- ● xanh **Danh sách kiểm kho – sẽ tự tạo phiếu mới** hoặc **Màn hình phiếu · N mã**, kèm cảnh báo nếu có bản nháp chưa lưu.

Thanh trạng thái chỉ đọc: không bấm, không gõ, không kéo cửa sổ Chrome lên trước. Khi Tool đang chạy hoặc đang đổi cơ sở, nó tạm dừng để không gửi lệnh chồng lên nhau.

### Mở Chrome Debug thủ công

Nếu không dùng nút **Mở Chrome Debug**, có thể tạo shortcut Chrome với lệnh sau:

```text
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9999 --user-data-dir="C:\ChromeDebug"
```

Cổng phải khớp với `debug_port` trong `config.json` (mặc định `9999`). Nên tắt các tiện ích chặn quảng cáo trên profile Chrome Debug nếu chúng làm gián đoạn thao tác ở KiotViet.

## Quy tắc xử lý

- Khớp cơ sở theo tiêu đề `C2`, ưu tiên lựa chọn tay đã lưu cho đúng tiêu đề đó; nếu chưa có thì ưu tiên số CS (`CS8`, `CS08`, `cs 08` đều là CS08), sau đó đến địa điểm: địa chỉ đầy đủ (`22 Vũ Trọng Phụng`) hoặc tên đường (`Vũ Trọng Phụng`). Không phân biệt hoa/thường và dấu. Lựa chọn tay trỏ tới cơ sở không còn trên web sẽ bị bỏ qua.

- Tool lấy số lượng hiện tại trên KiotViet và cộng với số lượng cân chỉnh trong Excel.
- Nếu tổng nhỏ hơn 0, mã hàng đó sẽ bị bỏ qua và được ghi vào nhật ký. Sản phẩm vẫn nằm trên phiếu với số lượng mặc định của KiotViet, cần sửa tay hoặc xóa trước khi lưu.
- Nhật ký ghi **STT trên phiếu KiotViet** cho các mã đã được thêm vào phiếu (tổng âm, cần kiểm tra), để dễ tìm lại; mã không tìm thấy trên web thì ghi số dòng Excel.
- Ô số lượng trống được tính là 0; ô không phải số thì mã đó bị bỏ qua.
- Mã xuất hiện nhiều lần trong Excel sẽ được cộng dồn lên web (Tool cảnh báo trước khi chạy).
- Tool chỉ chọn gợi ý tìm kiếm có đúng mã hàng, và chỉ thao tác khi ô **Mã hàng** ở dòng đầu bảng kiểm kho trùng khớp chính xác với mã đó. Mã có chung phần đầu không bị nhầm với nhau (ví dụ `7350107133709` và `7350107133709_D26.10` là hai mã khác nhau).
- Sau khi điền, Tool đọc lại ô số lượng để xác minh; nếu không khớp sẽ ghi `CẦN KIỂM TRA` vào nhật ký.
- Nếu không tìm thấy sản phẩm hoặc thao tác web gặp lỗi, Tool bỏ qua mã hiện tại, ghi lỗi và tiếp tục với mã tiếp theo.
- Sau khi chạy xong toàn bộ danh sách (không bấm Dừng), Tool **kiểm tra lại các mã đã điền thành công**: đọc cột **SL lệch** trên từng trang của phiếu (KiotViet hiển thị 15 mã mỗi trang) và so với số lượng trong Excel (cột D). Chỉ báo `SAI LỆCH` (kèm STT trên KiotViet) vào nhật ký cho mã đã điền số hợp lệ nhưng SL lệch vẫn khác Excel; các mã bị bỏ qua lúc chạy (tổng âm, không tìm thấy) đã được báo trước đó nên không báo lại.
- Sau đó Tool điền "Cân hàng lệch" vào ô **Ghi chú** của phiếu. Ghi chú cũ được giữ lại, nội dung mới thêm vào dòng dưới; nếu đã có sẵn thì không điền lại. Có thể đổi nội dung bằng khóa `note_text` trong `config.json` (để trống để tắt).
- Không nên sử dụng chuột hoặc bàn phím trong cửa sổ Chrome khi Tool đang chạy. Tool tự mở lại cửa sổ Chrome Debug nếu bị thu nhỏ (trang bị ẩn thì không gõ được mã vào ô tìm kiếm).
- Lệch trên phiếu chưa cân bằng được KiotViet tính theo tồn kho hiện tại. Nếu có bán hàng sau lúc chạy Tool, Lệch sẽ thay đổi dù Thực tế không đổi, nên hãy cân bằng phiếu ngay sau khi chạy.

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

Script chạy test trong `p310`, sau đó build `dist/AutoApp.exe` bằng `AutoApp.spec`. Nếu test lỗi thì không build. Trước khi đóng gói, script ghi mã commit ngắn và ngày build vào `version.py` (thêm `-dirty` nếu còn thay đổi chưa commit), rồi trả lại `BUILD = "dev"` sau khi build xong.

## Phiên bản

Số phiên bản nằm trong `version.py`:

- Tiêu đề cửa sổ hiển thị `Auto Robot v1.4.1` khi chạy từ mã nguồn, và `Auto Robot v1.4.1 (<mã commit> <ngày build>)` với bản `.exe`.
- Góc dưới bên phải cửa sổ và dòng đầu nhật ký luôn hiển thị phiên bản đang chạy, thuận tiện khi báo lỗi.
- Quy ước tăng số: sửa lỗi tăng số cuối (`1.4.0` → `1.4.1`), thêm tính năng tăng số giữa (`1.4.1` → `1.5.0`).

## Cấu trúc dự án

- `app.py`: điểm khởi chạy.
- `gui.py`: giao diện Tkinter, xử lý luồng và lưu cấu hình.
- `automation.py`: vòng lặp xử lý từng mã hàng, thống kê kết quả, xử lý dừng.
- `kiotviet.py`: thao tác Selenium với KiotViet (nhận diện trang, đổi cơ sở, tạo phiếu, tìm, điền, xác minh, chuyển tới mã trong nhật ký).
- `verify.py`: so SL lệch trên phiếu với số lượng trong Excel cho các mã đã điền.
- `branch_match.py`: khớp tiêu đề file với cơ sở trên web (theo số CS hoặc địa điểm, không phân biệt dấu).
- `excel_reader.py`: đọc Excel, phân tích ô bắt đầu, kiểm tra mã trùng/số lượng.
- `config.py`: đọc/ghi `config.json`, dò đường dẫn Chrome, kiểm tra thiết lập.
- `settings_dialog.py`: cửa sổ Cài đặt.
- `version.py`: số phiên bản và mã build hiển thị trên cửa sổ/nhật ký.
- `notifier.py`: âm thanh hệ thống và nhấp nháy taskbar khi chạy xong.
- `session.py`: giữ một phiên Selenium dùng chung, khóa để các luồng không chồng lệnh.
- `status_monitor.py`: đọc trạng thái KiotViet định kỳ và chuyển thành chữ/màu cho thanh trạng thái.
- `tests/`: test `pytest` cho các module trên (không cần Chrome hay KiotViet).
- `AutoApp.spec`: cấu hình đóng gói bằng PyInstaller, bao gồm các DLL Tcl/Tk cần thiết từ môi trường Conda.
- `build.ps1`: script test và build.
- `dist/AutoApp.exe`: bản executable sau khi build (đang được `.gitignore` bỏ qua).

## Lưu ý kỹ thuật

Tool phụ thuộc vào cấu trúc giao diện HTML hiện tại của KiotViet. Các cột của bảng hàng hóa (STT, Mã hàng, Tồn kho, Thực tế, SL lệch) được tìm theo thuộc tính `data-field` trên tiêu đề bảng (`rowNumber`, `ProductCode`, `SystemCount`, `ActualCount`, `AdjustmentValue`) nên không bị ảnh hưởng khi KiotViet đổi class của ô (ví dụ ô Mã hàng từ `cell-code` thành `cell-114`); class cũ chỉ dùng dự phòng. Các selector và hàm JS nằm ở đầu file `kiotviet.py`. `tests/test_grid_js.py` chạy các hàm JS đó trên trang HTML mô phỏng bằng Chrome headless.
