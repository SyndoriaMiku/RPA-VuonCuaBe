# Giao diện tối + chuyển động cho Auto Robot KiotViet

## 1. Mục đích và lý do

Cửa sổ hiện tại được người dùng mô tả là "khá sáng, đơn điệu, không có animation đẹp mắt".
Mục tiêu là làm Tool dễ nhìn hơn khi dùng nhiều giờ và bớt khô, **không đổi cách làm việc**.
Người dùng chính: chủ repo và nhân viên kiểm kho, trên Windows, thao tác liên tục trong buổi.

## 2. Mục tiêu

- Tông tối theo hướng "Xanh đêm" (bảng màu §7).
- Chuyển động ngắn, mượt, có mục đích (danh sách §6).
- Giữ nguyên toàn bộ chức năng: luồng 3 bước ①②③, nhật ký bấm được, thông báo khi xong,
  Cài đặt, tự đổi cơ sở, ghi log ra file, build PyInstaller onefile.
- Không làm cửa sổ treo/giật; không lặp lại lỗi "Not Responding" đã từng gặp.

## 3. Không làm (non-goals)

- Không đổi bố cục hay luồng thao tác (vẫn 3 thẻ bước như hiện tại).
- Không làm chế độ sáng/tối chuyển đổi qua lại.
- Không thêm tính năng nghiệp vụ mới.

## 4. Chống mục tiêu (anti-goals)

- Animation không được làm chậm thao tác, không chạy vô hạn khi không cần.
- Exe không phình to hay build phức tạp hơn đáng kể.
- Không phá 138 test hiện có và không sửa module logic (`automation`, `kiotviet`, `verify`,
  `excel_reader`, `branch_match`, `status_monitor`, `session`, `notifier`, `log_file`, `config`).

## 5. Ràng buộc

- Windows, Python 3.10 (conda `p310`), Tkinter; đóng gói PyInstaller onefile `AutoApp.spec`,
  `console=False`. PyInstaller hooks-contrib đã có `hook-customtkinter.py` (tự `collect_data_files`),
  nên không phải thêm dữ liệu theme bằng tay.
- Hộp thoại hệ thống (`filedialog`, `messagebox`) vẫn do Windows vẽ → không đổi được màu.
- `notifier.py` cần HWND cửa sổ chính; `customtkinter.CTk` vẫn là `Tk` nên giữ nguyên cách lấy.
- Tk chỉ làm được fade cửa sổ qua thuộc tính `-alpha` (Windows); không có hiệu ứng dựng sẵn.
- Mọi cập nhật widget phải ở main thread; giữ nguyên các hàm an toàn luồng `log`, `notify`,
  `update_progress`.
- `customtkinter` mới nhất là 6.0.0 (chưa cài trong `p310`); bản ổn định trước đó là 5.2.2.

## 6. Đã chốt với người dùng

- Hướng màu: **B — Xanh đêm** (bảng màu đầy đủ ở §7, đã xem mockup có chuyển động thật).
- Cách làm: **CustomTkinter** (thay ttk cho các widget chính).
- Chuyển động cần có:
  1. Chấm trạng thái: nhấp nháy nhẹ khi **chưa** kết nối (gợi ý cần mở Chrome), giữ yên khi đã kết
     nối; đổi màu mượt khi kết nối / mất kết nối; tạm dừng khi cửa sổ bị thu nhỏ.
  2. Nút **CHẠY**: sáng lên khi rê chuột, nhún nhẹ khi bấm; trạng thái tắt rõ ràng.
  3. Thanh tiến độ: giá trị chạy êm (nội suy) và có vệt sáng khi đang chạy.
  4. Nhật ký: dòng mới hiện dần; cuộn êm tới dòng mới thay vì nhảy.
  5. Thẻ bước: viền/nhãn đổi màu mượt khi bước hoàn tất hoặc đang chờ.
  6. Cửa sổ: fade-in khi mở (bỏ nếu gây cảm giác chậm).

## 7. Bảng màu và vai trò (hướng B)

| Vai trò | Mã | Dùng ở đâu |
|---|---|---|
| `bg` | `#0D1117` | nền cửa sổ |
| `surface` | `#161B22` | thẻ bước, thanh trạng thái |
| `surface_alt` | `#1C232D` | ô nhập, nút phụ, nền vệt tiến độ |
| `hover` | `#222C39` | rê chuột trên nút phụ / mục |
| `border` | `#263040` | viền thẻ, ô nhập, đường phân cách |
| `text` | `#E6EDF3` | chữ chính |
| `muted` | `#8B98A5` | chú thích, nhãn phụ |
| `accent` | `#4C8DFF` | nút CHẠY, tiêu điểm, bước đang làm |
| `accent_hover` | `#6BA1FF` | rê chuột nút CHẠY |
| `on_accent` | `#0B1220` | chữ trên nền accent |
| `ok` | `#43C77E` | log thành công, chấm đã kết nối |
| `warn` | `#E3B341` | log bỏ qua / cảnh báo |
| `error` | `#F0666C` | log lỗi, nút DỪNG, chấm mất kết nối |
| `link` | `#7CB6FF` | mã hàng bấm được trong nhật ký |

Chữ/nền và chữ/accent đạt tương phản WCAG ≥ 4.5:1. Ô nhập có nền `surface_alt`, chữ `text`,
viền `border`, khi focus viền đổi `accent`.

## 8. Cách làm và đánh đổi

Đã cân nhắc ba cách, người dùng chọn B:

- **A. ttk tối, không thêm thư viện** — an toàn nhất, giữ nguyên đóng gói, nhưng thẻ vẫn là hình
  chữ nhật viền 1px, không bo góc, vài chỗ (hover, bo góc, thanh cuộn nhật ký) khó đẹp.
- **B. CustomTkinter** — *được chọn*: gần mockup nhất (bo góc, hover, tông tối sẵn), ít code hơn để
  đạt thẩm mỹ. Đổi lại phải thay widget chính, hộp thoại và nhật ký làm lại, exe nặng thêm, và bản
  6.0.0 còn mới nên cần kiểm chứng tương thích.
- **C. Lai (ttk + Canvas vẽ tay)** — trung dung, nhưng thêm code bảo trì cho nút/thẻ vẽ tay.

Lý do chấp nhận rủi ro của B: người dùng ưu tiên thẩm mỹ; bề mặt giao diện nhỏ gọn (một màn hình
chính + một hộp thoại Cài đặt), nên phạm vi thay thế hữu hạn.

## 9. Thiết kế kỹ thuật

### 9.1 Phạm vi file

- **MỚI `ui_theme.py`** — bảng màu §7, tên font, bán kính bo góc, thời lượng animation, hàm
  `blend(color_a, color_b, t)` và một tiện ích tween dùng chung. Thuần Python, test được, không cần Tk.
- **MỚI `ui_widgets.py`** (có thể gộp vào `gui.py` nếu ngắn) — widget dùng lại: chấm trạng thái,
  nút chính có hover, thanh tiến độ êm, thẻ bo góc.
- **`gui.py`** — đổi cửa sổ gốc và các widget chính sang CustomTkinter; giữ nguyên tên thuộc tính
  (`self.btn_run`, `self.txt_log`, `self.lbl_status`, …) và toàn bộ hàm logic để phần còn lại của
  file không phải sửa nhiều.
- **`settings_dialog.py`** — `CTkToplevel` + ô/nhãn CTk; giữ nguyên hành vi kiểm tra và lưu.
- **`requirements.txt`** — thêm `customtkinter` (ghim bản chốt, mặc định 6.0.0; xem §12/R1).
- **`version.py`**, **`tests/test_version.py`** — bump `1.6.0`.
- **`README.md`** — cập nhật phần giao diện/tông tối, phụ thuộc mới, phiên bản.
- `AutoApp.spec` — dự kiến không phải sửa; chỉ kiểm tra lại sau build.

### 9.2 Cửa sổ và widget

- `tk.Tk()` → `customtkinter.CTk()`, đặt `ctk.set_appearance_mode("dark")`, nền `palette.bg`.
- Thẻ bước: `CTkFrame` bo góc, viền `border`, nền `surface`; tiêu đề bước dùng `accent` khi là
  bước đang làm.
- Ô nhập: `CTkEntry`; combobox: `CTkComboBox`; nút thường: `CTkButton` với `fg_color=surface_alt`,
  `hover_color=hover`; nút CHẠY `accent`/`accent_hover`; nút DỪNG `error`.
- Thanh tiến độ: `CTkProgressBar` (hoặc Canvas tự vẽ nếu cần vệt sáng).
- Nhật ký: widget `tkinter.Text` (vì cần nhiều màu + tag bấm được) đặt màu theo palette, ghép
  `CTkScrollbar`; giữ nguyên tag màu, tag link, Ctrl+C/Ctrl+A, nút **Xóa** / **Sao chép tất cả**.
- Thanh trạng thái: `CTkFrame` + nhãn chấm (dùng Canvas/Label) + nhãn chữ + nút **Mở Chrome Debug** /
  **Hiện Chrome** (giữ nguyên logic hiện/ẩn).
- `filedialog` / `messagebox` giữ nguyên; `notifier.top_window_handle(self.root)` không đổi.

### 9.3 Chuyển động

- Tất cả chạy bằng `root.after`, bước 16–25 ms, tổng 120–250 ms mỗi hiệu ứng; huỷ khi đóng cửa sổ;
  không bao giờ gọi hàm chặn (Selenium, subprocess) trong callback animation.
- Vòng lặp nền (chấm trạng thái, vệt tiến độ) chỉ chạy khi cần; khi cửa sổ bị thu nhỏ
  (`root.state() == "iconic"`) thì tạm dừng để đỡ tốn CPU.
- Đổi màu dùng `ui_theme.blend` để nội suy giữa hai màu theo `t` (0→1).

### 9.4 Luồng và an toàn

- `log`, `notify`, `update_progress` vẫn tự chuyển về main thread như hiện tại; animation cũng chỉ
  chạy trên main thread. Không thêm việc chặn nào vào luồng giao diện.

## 10. Kiểm thử

- Toàn bộ test hiện có phải xanh (không đụng logic nghiệp vụ).
- Test mới cho `ui_theme.py`: giá trị bảng màu; `blend` đúng ở hai đầu và ở giữa; tương phản chữ/nền
  và chữ/accent đạt ngưỡng (tính theo công thức WCAG ngay trong test).
- Smoke thủ công (TDD không phủ được giao diện Tk): chạy `app.py`, kiểm tra mọi nút/nhãn đúng tông
  tối, log bấm mã được, Cài đặt lưu được, chạy thử một mã, rồi chạy `dist/AutoApp.exe` bản build.

## 11. Đóng gói và phát hành

- `build.ps1` giữ nguyên quy trình (test trong `p310` → PyInstaller → đóng dấu `BUILD` rồi khôi phục).
- Kiểm tra exe: mở được, giao diện tối đúng, `customtkinter` có trong gói.
- Bump `1.6.0`; commit; push / phát hành release khi người dùng yêu cầu.

## 12. Rủi ro và cách xử lý

- **R1 — CTk 6.0.0 còn mới**, có thể lỗi tương thích với Tk 8.6 trên Windows: làm một spike nhỏ
  (một cửa sổ CTk có thẻ/ô/nút/nhật ký) trước khi thay toàn bộ; nếu vỡ thì ghim `5.2.2`.
- **R2 — Exe nặng thêm** (ước tính +2–5 MB do theme/dữ liệu CTk): chấp nhận.
- **R3 — Canh hàng widget CTk trong `CTkToplevel`** đôi khi lệch: kiểm tra bằng mắt, chỉnh padding.
- **R4 — Animation tốn CPU**: giới hạn số vòng lặp, dừng khi cửa sổ ẩn.
- **R5 — Nhật ký dùng `tkinter.Text` lẫn trong CTk**: chấp nhận vì cần tag nhiều màu và link bấm; chỉ
  chỉnh màu cho khớp, không đổi hành vi.

## 13. Việc còn để người làm quyết

Cả bốn việc dưới đây đã được chốt ở §14 (người đọc thiết kế có thể xem lại và phản đối):

- Tách `ui_widgets.py` riêng hay gộp vào `gui.py`.
- Chốt CTk `6.0.0` hay ghim `5.2.2` (theo kết quả spike R1).
- Có giữ fade-in cửa sổ hay bỏ.
- Thời lượng và biên độ cụ thể của từng animation.

## 14. Quyết định đã chốt sau khi soát thiết kế

Người soát độc lập nêu 5 chỗ phải đoán; không chỗ nào cần hỏi thêm người dùng, nên chốt như sau:

1. **Chấm trạng thái nhấp nháy khi nào.** `StatusMonitor` chỉ phát kết quả cuối (có/không kết nối)
   và không được sửa (§4), nên không có tín hiệu "đang dò". Chốt: chấm nhấp nháy nhẹ **liên tục khi
   chưa kết nối** (đúng nghĩa chỉ báo cần chú ý), giữ yên khi đã kết nối, và **tạm dừng khi cửa sổ
   thu nhỏ**. Đây là một vòng lặp rất nhẹ (một `after` đổi độ sáng mỗi ~60 ms), không vi phạm tinh
   thần "không chạy vô hạn khi không cần" vì luôn gắn với trạng thái đang hiển thị.
2. **Không thêm "trạng thái bước" cho thẻ ①②③.** Người dùng nêu vấn đề thị giác, không phải luồng
   làm việc, nên không suy diễn "bước đã xong" (dễ sai và gây nhiễu). Thẻ chỉ có: viền/nhãn đổi màu
   khi **rê chuột** hoặc khi **ô bên trong đang được nhập** (focus), cộng thêm bước ③ và thanh tiến độ
   phản ánh trạng thái đang chạy vốn đã có (`running`).
3. **Nhật ký khi dòng về dồn dập.** Mỗi dòng vẫn chèn ngay; chỉ chạy fade khi cách dòng trước
   ≥ 150 ms (chạy thật thì phần lớn dòng hiện tức thì). Cuộn êm gộp thành **một** tween cho mỗi đợt
   thay vì một tween mỗi dòng, tránh hàng đợi animation gây trễ — đúng bài học từ các lỗi treo trước.
4. **Có làm tối thanh tiêu đề Windows.** Vì mục tiêu là thẩm mỹ, thanh tiêu đề sáng trên thân tối sẽ
   bị coi là chưa xong. Dùng `ctypes` gọi `DwmSetWindowAttribute` (`DWMWA_USE_IMMERSIVE_DARK_MODE`,
   thử 20 rồi 19) trong `try/except`; máy đích là Windows 11 nên hỗ trợ. Không hỗ trợ thì bỏ qua êm.
5. **Tỉ lệ hiển thị (DPI).** Khoá thang đo của CTk (`set_widget_scaling(1.0)`,
   `set_window_scaling(1.0)`) và giữ nguyên cách tính kích thước/căn giữa theo pixel hiện có, để bố
   cục không lệch so với hôm nay. Nếu sau này chữ bị mờ trên màn hình có scaling, xử lý riêng bằng
   DPI-awareness (ngoài phạm vi lần này).

Chốt bốn việc ở §13:

- **Gộp widget vào `gui.py`**, chỉ tách `ui_theme.py` (bảng màu + `blend` + tween dùng chung) vì phần
  đó thuần Python và test được. Chỉ tách thêm `ui_widgets.py` nếu `gui.py` phình quá.
- **CTk 6.0.0 trước**, spike R1 quyết định; vỡ thì ghim `5.2.2`.
- **Giữ fade-in cửa sổ**, 120 ms, bọc `try/except`; nếu thấy giật thì bỏ.
- **Thời lượng:** chấm trạng thái chu kỳ ~1,2 s; hover nút 120 ms; nội suy tiến độ 180 ms/bước, vệt
  sáng 1,4 s; fade dòng nhật ký 160 ms; đổi viền thẻ 160 ms.
