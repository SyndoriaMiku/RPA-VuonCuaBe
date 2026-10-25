import os
import subprocess
import threading
import tkinter as tk
import tkinter.scrolledtext as scrolledtext
from tkinter import filedialog, messagebox, ttk

import automation
import branch_match
import config as app_config
import excel_reader
import kiotviet
import verify

FONT = "Segoe UI"
COLORS = {
    "bg": "#F3F5F7",
    "card": "#FFFFFF",
    "text": "#1F2933",
    "muted": "#6B7785",
    "primary": "#2E7D32",
    "primary_hover": "#1B5E20",
    "danger": "#C62828",
    "danger_hover": "#8E0000",
    "border": "#D9DEE3",
}
LOG_COLORS = {
    "info": "#374151",
    "ok": "#2E7D32",
    "warn": "#B45309",
    "error": "#C62828",
}


class App:
    def __init__(self):
        self.config = app_config.load_config()
        self.stop_event = threading.Event()
        self.running = False

        self.root = tk.Tk()
        self.root.title("Auto Robot")
        self.root.geometry("920x740")
        self.root.minsize(780, 620)
        self.root.configure(bg=COLORS["bg"])
        self._setup_style()
        self._build_ui()
        self._center_window()
        self._refresh_chrome_status()
        self._refresh_cell_hint()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # File đã dùng lần trước: load lại để hiện tiêu đề và tự đổi cơ sở như khi vừa chọn file
        remembered_file = self.config["file_path"]
        if remembered_file and os.path.isfile(remembered_file):
            self.root.after(300, lambda: self.load_file(remembered_file))

    # ------------------------------------------------------------------
    # Giao diện
    # ------------------------------------------------------------------
    def _setup_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=(FONT, 10), background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card"])
        style.configure("TLabel", background=COLORS["bg"])
        style.configure("Card.TLabel", background=COLORS["card"])
        style.configure("Muted.TLabel", background=COLORS["card"], foreground=COLORS["muted"], font=(FONT, 9))
        style.configure("Title.TLabel", font=(FONT, 17, "bold"), foreground=COLORS["text"])
        style.configure("Subtitle.TLabel", foreground=COLORS["muted"])
        style.configure("CardTitle.TLabel", background=COLORS["card"], font=(FONT, 11, "bold"))
        style.configure(
            "TEntry", fieldbackground="#FFFFFF", bordercolor=COLORS["border"], lightcolor=COLORS["border"],
            darkcolor=COLORS["border"], padding=5,
        )
        style.configure("TCombobox", padding=4)

        style.configure("TButton", padding=(12, 7), borderwidth=1, bordercolor=COLORS["border"],
                        background="#FFFFFF", focuscolor=COLORS["card"])
        style.map("TButton", background=[("active", "#E8ECEF"), ("disabled", "#EEF0F2")],
                  foreground=[("disabled", "#A0A8B0")])
        style.configure("Primary.TButton", font=(FONT, 11, "bold"), foreground="#FFFFFF",
                        background=COLORS["primary"], bordercolor=COLORS["primary"], padding=(18, 9))
        style.map("Primary.TButton",
                  background=[("active", COLORS["primary_hover"]), ("disabled", "#A5C8A7")],
                  foreground=[("disabled", "#FFFFFF")])
        style.configure("Danger.TButton", font=(FONT, 11, "bold"), foreground="#FFFFFF",
                        background=COLORS["danger"], bordercolor=COLORS["danger"], padding=(18, 9))
        style.map("Danger.TButton",
                  background=[("active", COLORS["danger_hover"]), ("disabled", "#E5A8A8")],
                  foreground=[("disabled", "#FFFFFF")])
        style.configure("Horizontal.TProgressbar", troughcolor="#E1E5E9", background=COLORS["primary"],
                        bordercolor=COLORS["border"], lightcolor=COLORS["primary"], darkcolor=COLORS["primary"])

    def _card(self, parent, title, step):
        outer = tk.Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"], highlightthickness=1)
        outer.pack(fill=tk.X, pady=(0, 10))
        inner = ttk.Frame(outer, style="Card.TFrame", padding=(14, 10))
        inner.pack(fill=tk.BOTH, expand=True)
        ttk.Label(inner, text=f"{step}  {title}", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))
        inner.columnconfigure(1, weight=1)
        return inner

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=(18, 14))
        main.pack(fill=tk.BOTH, expand=True)

        # Tiêu đề
        header = ttk.Frame(main)
        header.pack(fill=tk.X, pady=(0, 12))
        self.lbl_title = ttk.Label(header, text="Auto Robot KiotViet", style="Title.TLabel")
        self.lbl_title.pack(anchor="w")
        ttk.Label(
            header, style="Subtitle.TLabel",
            text="Nhập số lượng cân chỉnh từ Excel vào màn hình kiểm kho KiotViet. Luôn kiểm tra lại trước khi lưu phiếu.",
        ).pack(anchor="w")

        # Bước 1: Chrome
        card1 = self._card(main, "Mở Chrome Debug", "①")
        self.lbl_chrome = ttk.Label(card1, text="", style="Muted.TLabel")
        self.lbl_chrome.grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Button(card1, text="Mở Chrome Debug", command=self.open_chrome_debug).grid(
            row=1, column=3, sticky="e", padx=(10, 0))
        ttk.Label(
            card1, style="Muted.TLabel",
            text="Sau khi Chrome mở, đăng nhập KiotViet và vào màn hình kiểm kho có ô tìm kiếm sản phẩm.",
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # Bước 2: Excel
        card2 = self._card(main, "Chọn dữ liệu Excel", "②")
        ttk.Label(card2, text="File Excel", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        self.var_file = tk.StringVar(value=self.config["file_path"])
        ttk.Entry(card2, textvariable=self.var_file).grid(row=1, column=1, columnspan=2, sticky="ew", padx=8)
        ttk.Button(card2, text="Chọn file…", command=self.choose_file).grid(row=1, column=3, padx=(0, 0))

        ttk.Label(card2, text="Tên sheet", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        self.var_sheet = tk.StringVar(value=self.config["sheet_name"])
        ttk.Entry(card2, textvariable=self.var_sheet).grid(row=2, column=1, columnspan=2, sticky="ew", padx=8)
        ttk.Label(card2, text="Để trống = sheet đang active", style="Muted.TLabel").grid(
            row=2, column=3, sticky="w", padx=(10, 0))

        ttk.Label(card2, text="Ô bắt đầu", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=4)
        self.var_start = tk.StringVar(value=self.config["start_cell"])
        self.var_start.trace_add("write", lambda *_: self._refresh_cell_hint())
        ttk.Entry(card2, textvariable=self.var_start, width=10).grid(row=3, column=1, sticky="w", padx=8)
        self.lbl_cell_hint = ttk.Label(card2, text="", style="Muted.TLabel")
        self.lbl_cell_hint.grid(row=3, column=2, columnspan=2, sticky="w")

        # Bước 3: Chạy
        card3 = self._card(main, "Kiểm tra và chạy", "③")
        ttk.Label(card3, text="Chờ kết quả tối đa (giây)", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        self.combo_wait = ttk.Combobox(card3, values=["3.0", "4.0", "5.0", "7.0", "10.0"], width=6)
        self.combo_wait.set(str(self.config["wait_seconds"]))
        self.combo_wait.grid(row=1, column=1, sticky="w", padx=8)

        buttons = ttk.Frame(card3, style="Card.TFrame")
        buttons.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 4))
        self.btn_check = ttk.Button(buttons, text="Kiểm tra file", command=self.check_file)
        self.btn_check.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_run = ttk.Button(buttons, text="▶  CHẠY TỰ ĐỘNG", style="Primary.TButton", command=self.start_run)
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_stop = ttk.Button(buttons, text="■  DỪNG", style="Danger.TButton", command=self.request_stop,
                                   state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT)

        progress_row = ttk.Frame(card3, style="Card.TFrame")
        progress_row.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        self.progress_bar = ttk.Progressbar(progress_row, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_progress = ttk.Label(progress_row, text="", style="Card.TLabel", width=10, anchor="e")
        self.lbl_progress.pack(side=tk.LEFT, padx=(8, 0))

        # Nhật ký
        log_header = ttk.Frame(main)
        log_header.pack(fill=tk.X)
        ttk.Label(log_header, text="Nhật ký", font=(FONT, 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(log_header, text="Xóa", command=self.clear_log, padding=(8, 3)).pack(side=tk.RIGHT)
        ttk.Button(log_header, text="Sao chép tất cả", command=self.copy_all_log, padding=(8, 3)).pack(
            side=tk.RIGHT, padx=6)

        self.txt_log = scrolledtext.ScrolledText(
            main, height=10, wrap=tk.WORD, bg="#FFFFFF", fg=LOG_COLORS["info"], font=("Consolas", 10),
            relief=tk.FLAT, highlightbackground=COLORS["border"], highlightthickness=1, padx=8, pady=6,
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        for level, color in LOG_COLORS.items():
            self.txt_log.tag_config(level, foreground=color)
        self.txt_log.config(state=tk.DISABLED, insertwidth=0, takefocus=1)
        self.txt_log.bind("<Control-c>", self._copy_selected_log)
        self.txt_log.bind("<Control-a>", self._select_all_log)
        self.txt_log.bind("<Button-1>", lambda event: self.txt_log.focus_set())

    def _center_window(self):
        self.root.update_idletasks()
        width, height = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 3
        self.root.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _refresh_cell_hint(self):
        try:
            start_cell, _, quantity_column, start_row = excel_reader.parse_start_cell(self.var_start.get())
            self.lbl_cell_hint.config(
                text=f"Mã hàng đọc từ {start_cell}, số lượng đọc ở {quantity_column}{start_row} trở xuống")
        except ValueError:
            self.lbl_cell_hint.config(text="Ô bắt đầu chưa hợp lệ, ví dụ: B4")

    def _refresh_chrome_status(self):
        path = app_config.find_chrome(self.config["chrome_path"])
        if path:
            self.lbl_chrome.config(text=f"Chrome: {path}  ·  cổng debug {self.config['debug_port']}")
        else:
            self.lbl_chrome.config(text="Chưa tìm thấy Chrome. Bấm nút bên phải để chọn chrome.exe.")

    # ------------------------------------------------------------------
    # Tiện ích an toàn đa luồng
    # ------------------------------------------------------------------
    def _on_main_thread(self):
        return threading.current_thread() is threading.main_thread()

    def log(self, message, level="info"):
        def apply():
            self.txt_log.config(state=tk.NORMAL)
            self.txt_log.insert(tk.END, message + "\n", level)
            self.txt_log.see(tk.END)
            self.txt_log.config(state=tk.DISABLED)

        if self._on_main_thread():
            apply()
        else:
            self.root.after(0, apply)

    def notify(self, kind, title, message):
        """Hiện popup (showerror/showwarning/showinfo) an toàn từ bất kỳ luồng nào."""
        self.root.after(0, lambda: getattr(messagebox, kind)(title, message))

    def update_progress(self, done, total):
        def apply():
            self.progress_bar.config(maximum=max(total, 1), value=done)
            self.lbl_progress.config(text=f"{done}/{total}" if total else "")

        if self._on_main_thread():
            apply()
        else:
            self.root.after(0, apply)

    def _set_running(self, running, can_stop=True):
        self.running = running
        self.btn_run.config(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_check.config(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_stop.config(state=tk.NORMAL if running and can_stop else tk.DISABLED)

    def ask_yes_no(self, title, message):
        """Hỏi Có/Không từ luồng phụ: hiện hộp thoại trên main thread và chờ câu trả lời."""
        answer = {}
        answered = threading.Event()

        def ask():
            answer["value"] = messagebox.askyesno(title, message, icon="warning")
            answered.set()

        self.root.after(0, ask)
        answered.wait()
        return answer["value"]

    # ------------------------------------------------------------------
    # Nhật ký
    # ------------------------------------------------------------------
    def clear_log(self):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def copy_all_log(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.txt_log.get("1.0", tk.END).strip())

    def _copy_selected_log(self, event=None):
        try:
            selected_text = self.txt_log.selection_get()
        except tk.TclError:
            return "break"
        self.root.clipboard_clear()
        self.root.clipboard_append(selected_text)
        return "break"

    def _select_all_log(self, event=None):
        self.txt_log.tag_add(tk.SEL, "1.0", tk.END)
        self.txt_log.mark_set(tk.INSERT, "1.0")
        self.txt_log.see(tk.INSERT)
        return "break"

    # ------------------------------------------------------------------
    # Cấu hình
    # ------------------------------------------------------------------
    def _save_config(self):
        self.config.update({
            "file_path": self.var_file.get().strip(),
            "sheet_name": self.var_sheet.get(),
            "start_cell": self.var_start.get().strip(),
            "wait_seconds": self.combo_wait.get(),
        })
        try:
            app_config.save_config(self.config)
        except OSError:
            pass  # Không lưu được cấu hình (ví dụ thư mục chỉ đọc) không được làm gián đoạn công việc

    def _on_close(self):
        self._save_config()
        self.root.destroy()

    # ------------------------------------------------------------------
    # Chrome
    # ------------------------------------------------------------------
    def open_chrome_debug(self):
        chrome_path = app_config.find_chrome(self.config["chrome_path"])
        if not chrome_path:
            chrome_path = filedialog.askopenfilename(
                title="Chọn chrome.exe", filetypes=[("chrome.exe", "chrome.exe"), ("All Files", "*.*")])
            if not chrome_path:
                return
            self.config["chrome_path"] = chrome_path
            self._refresh_chrome_status()

        subprocess.run(["taskkill", "/f", "/im", "chromedriver.exe"], capture_output=True)

        port = self.config["debug_port"]
        try:
            subprocess.Popen([
                chrome_path,
                f"--remote-debugging-port={port}",
                f"--user-data-dir={self.config['user_data_dir']}",
            ])
        except OSError as exc:
            messagebox.showerror("Lỗi", f"Không thể mở Chrome Debug:\n{exc}")
            return
        self.log(f"Đã mở Chrome Debug (cổng {port}).", "ok")

    # ------------------------------------------------------------------
    # Excel
    # ------------------------------------------------------------------
    def choose_file(self):
        file_path = filedialog.askopenfilename(
            title="Chọn File Excel", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if not file_path:
            return
        self.load_file(file_path)

    def load_file(self, file_path):
        """Đọc tiêu đề C2, tóm tắt dữ liệu và tự đổi cơ sở trên web cho khớp file."""
        self.var_file.set(file_path)

        try:
            sheet = excel_reader.open_sheet(file_path, self.var_sheet.get().strip())
        except Exception:
            # Sheet mặc định không tồn tại trong file này: dùng sheet đang active để vẫn hiện được thông tin
            try:
                sheet = excel_reader.open_sheet(file_path, "")
            except Exception:
                self.lbl_title.config(text="[Lỗi đọc file Excel]")
                self.root.title("Auto Robot")
                messagebox.showerror("Lỗi đọc Excel", "Không thể đọc thông tin từ file Excel đã chọn.")
                return

        c2_value = sheet["C2"].value
        title = str(c2_value) if c2_value else "[Ô C2 trống]"
        self.lbl_title.config(text=title)
        self.root.title(f"Auto Robot - {c2_value}" if c2_value else "Auto Robot")

        try:
            _, scan_column, _, scan_start_row = excel_reader.parse_start_cell(self.var_start.get())
        except ValueError:
            scan_column, scan_start_row = "B", 4

        self.log("--- ĐÃ TẢI FILE ---", "ok")
        self.log(f"File: {os.path.basename(file_path)}")
        self.log(f"Sheet: {sheet.title}")
        self.log(excel_reader.summarize_data_column(sheet, scan_column, scan_start_row))
        self._save_config()

        if c2_value:
            self.auto_switch_branch(str(c2_value))

    def _load_items(self):
        """Đọc Excel theo cấu hình trên giao diện.

        Trả về (items, start_cell, quantity_column, tiêu đề C2) hoặc None nếu lỗi.
        """
        try:
            start_cell, code_column, quantity_column, start_row = excel_reader.parse_start_cell(self.var_start.get())
        except ValueError as exc:
            messagebox.showerror("Lỗi", str(exc))
            return None

        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file Excel trước.")
            return None

        try:
            sheet = excel_reader.open_sheet(file_path, self.var_sheet.get().strip())
            items = excel_reader.collect_excel_data(sheet, code_column, quantity_column, start_row)
            title = str(sheet["C2"].value or "").strip()
        except ValueError as exc:
            messagebox.showerror("Lỗi", str(exc))
            return None
        except Exception as exc:
            messagebox.showerror("Lỗi đọc Excel", f"Có lỗi xảy ra khi đọc file:\n{exc}")
            return None

        if not items:
            messagebox.showwarning("Cảnh báo", f"Không tìm thấy dữ liệu nào bắt đầu từ ô {start_cell}.")
            return None
        return items, start_cell, quantity_column, title

    def check_file(self):
        loaded = self._load_items()
        if not loaded:
            return
        items, start_cell, quantity_column, _ = loaded
        self.log(f"--- KIỂM TRA FILE (từ {start_cell}, số lượng ở cột {quantity_column}) ---", "info")
        for level, text in excel_reader.describe_precheck(excel_reader.precheck(items)):
            self.log(text, level)
        self.log("Bước này chỉ đọc Excel, không thao tác gì trên KiotViet.")
        self._save_config()

    # ------------------------------------------------------------------
    # Chạy tự động
    # ------------------------------------------------------------------
    def start_run(self):
        if self.running:
            return
        loaded = self._load_items()
        if not loaded:
            return
        items, start_cell, quantity_column, title = loaded

        report = excel_reader.precheck(items)
        if report["duplicates"] or report["invalid"]:
            warnings = [text for level, text in excel_reader.describe_precheck(report) if level in ("warn", "error")]
            if not messagebox.askyesno(
                "Phát hiện vấn đề trong file",
                "\n".join(warnings[:15]) + ("\n..." if len(warnings) > 15 else "") + "\n\nVẫn tiếp tục chạy?",
                icon="warning",
            ):
                return

        try:
            result_timeout = float(self.combo_wait.get())
        except ValueError:
            result_timeout = 5.0

        self._save_config()
        self.stop_event.clear()
        self._set_running(True)
        self.update_progress(0, 0)

        thread = threading.Thread(
            target=self._worker, args=(items, start_cell, quantity_column, result_timeout, title), daemon=True)
        thread.start()

    def request_stop(self):
        self.stop_event.set()
        self.btn_stop.config(state=tk.DISABLED)
        self.log("Đã nhận yêu cầu dừng, Tool sẽ dừng sau khi xong mã hàng hiện tại...", "warn")

    def _worker(self, items, start_cell, quantity_column, result_timeout, title):
        try:
            try:
                driver = kiotviet.connect_driver(self.config["debug_port"])
            except Exception as exc:
                self.notify("showerror", "Lỗi kết nối",
                            f"Không thể kết nối với Chrome Debug. Hãy bấm 'Mở Chrome Debug' trước.\n\n{exc}")
                return

            self.log("----------------------------------------")
            if not self._confirm_branch(driver, title):
                self.log("Đã hủy chạy vì cơ sở không khớp.", "warn")
                return

            try:
                page = kiotviet.ensure_stocktake_page(driver)
            except Exception as exc:
                self.notify("showerror", "Chưa ở màn hình Kiểm kho", str(exc).split("\n")[0])
                return
            if page == "created":
                self.log("Đang ở danh sách kiểm kho: đã bấm 'Kiểm kho' để tạo phiếu mới.", "ok")

            self.log(f"Đọc Excel từ {start_cell}; số lượng tại cột {quantity_column}.")
            self.log(f"Bắt đầu chạy {len(items)} mã hàng...")

            stats = automation.run_items(
                driver, items, result_timeout, self.stop_event, self.log, self.update_progress)

            self.log("--- ĐÃ DỪNG THEO YÊU CẦU ---" if stats.stopped else "--- HOÀN TẤT QUÁ TRÌNH ---",
                     "warn" if stats.stopped else "ok")
            self.log(
                f"Thành công: {stats.ok} | Bỏ qua: {stats.skipped} | Lỗi/cần kiểm tra: {stats.error} | "
                f"Tổng: {len(items)}",
                "ok" if not (stats.skipped or stats.error) else "info",
            )

            if not stats.stopped:
                self._verify_stocktake(driver, stats.succeeded)
                self._fill_note(driver)
        finally:
            self.root.after(0, lambda: self._set_running(False))

    # ------------------------------------------------------------------
    # Cơ sở làm việc
    # ------------------------------------------------------------------
    def auto_switch_branch(self, title):
        """Đổi cơ sở trên web cho khớp tiêu đề file (chạy nền, không chặn giao diện)."""
        if self.running:
            return
        self._set_running(True, can_stop=False)
        threading.Thread(target=self._switch_branch_worker, args=(title,), daemon=True).start()

    def _switch_branch_worker(self, title):
        try:
            try:
                driver = kiotviet.connect_driver(self.config["debug_port"])
                current = kiotviet.get_current_branch(driver)
                branches = kiotviet.get_branches(driver)
            except Exception:
                self.log("Chưa kết nối được Chrome Debug/KiotViet nên chưa tự đổi cơ sở.", "warn")
                return
            if not current or not branches:
                self.log("Không đọc được danh sách cơ sở trên web nên chưa tự đổi cơ sở.", "warn")
                return

            matches = branch_match.match_branches(title, branches)
            if not matches:
                self.log(f"Không xác định được cơ sở từ tiêu đề file. Giữ cơ sở hiện tại: {current[1]}.", "warn")
                return
            if len(matches) > 1:
                names = ", ".join(name for _, name in matches)
                self.log(f"Tiêu đề file khớp nhiều cơ sở ({names}). Giữ cơ sở hiện tại: {current[1]}.", "warn")
                return

            target_id, target_name = matches[0]
            if target_id == current[0]:
                self.log(f"Cơ sở đang làm việc đã đúng: {current[1]}.", "ok")
                return

            if kiotviet.get_page_state(driver) == kiotviet.PAGE_EDIT and kiotviet.count_item_rows(driver):
                self.log(
                    f"Phiếu kiểm kho đang mở đã có hàng, không tự đổi cơ sở để tránh mất dữ liệu. "
                    f"Cơ sở theo file: {target_name}; đang làm việc: {current[1]}.", "warn")
                return

            self.log(f"Đang đổi cơ sở: {current[1]} → {target_name}...")
            kiotviet.change_branch(driver, target_id)
            self.log(f"Đã đổi cơ sở làm việc sang {target_name}.", "ok")
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không đổi được cơ sở: {first_line or type(exc).__name__}", "error")
        finally:
            self.root.after(0, lambda: self._set_running(False))

    def _confirm_branch(self, driver, title):
        """So cơ sở trong file với cơ sở đang làm việc; không khớp thì hỏi có tiếp tục không."""
        if not title:
            self.log("Ô C2 trống, không kiểm tra được cơ sở.", "warn")
            return True
        try:
            current = kiotviet.get_current_branch(driver)
        except Exception:
            current = None
        if not current:
            self.log("Không đọc được cơ sở đang làm việc trên web, bỏ qua bước kiểm tra cơ sở.", "warn")
            return True

        if branch_match.is_same_branch(title, current[1]):
            self.log(f"Cơ sở khớp: {current[1]}.", "ok")
            return True

        self.log(f"Cơ sở có thể không khớp. File: {title} | Web: {current[1]}", "warn")
        return self.ask_yes_no(
            "Cơ sở không khớp",
            f"Tiêu đề file: {title}\nCơ sở đang làm việc trên web: {current[1]}\n\n"
            "Hai cơ sở có vẻ không khớp. Vẫn tiếp tục chạy?",
        )

    def _verify_stocktake(self, driver, filled_items):
        """Đọc lại SL lệch trên từng trang của phiếu và so với Excel cho các mã đã điền thành công."""
        if not filled_items:
            return
        self.log("--- KIỂM TRA LẠI SL LỆCH CÁC MÃ ĐÃ ĐIỀN ---")
        try:
            web_rows = kiotviet.read_stocktake_rows(driver)
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không kiểm tra lại được phiếu: {first_line or type(exc).__name__}", "error")
            return

        report = verify.compare_adjustments(filled_items, web_rows)
        for level, text in verify.describe_report(report, kiotviet.format_qty):
            self.log(text, level)

    def _fill_note(self, driver):
        note_text = str(self.config["note_text"]).strip()
        if not note_text:
            return
        try:
            result = kiotviet.fill_note(driver, note_text)
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không điền được ghi chú '{note_text}': {first_line or type(exc).__name__}", "error")
            return

        if result == "exists":
            self.log(f"Ghi chú đã có sẵn '{note_text}', không điền thêm.")
        else:
            self.log(f"Đã điền ghi chú: '{note_text}'.", "ok")

    def run(self):
        self.root.mainloop()


def main():
    App().run()
