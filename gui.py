import ctypes
import os
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import automation
import branch_match
import config as app_config
import excel_reader
import kiotviet
import log_file
import notifier
import session
import settings_dialog
import status_monitor
import ui_theme
import verify
import version

FONT = ui_theme.FONT
RADIUS = ui_theme.RADIUS
C = ui_theme.PALETTE
COLORS = {
    "bg": "#F3F5F7",
    "card": "#FFFFFF",
    "text": "#1F2933",
    "muted": "#6B7785",
    "primary": "#2E7D32",
    "primary_hover": "#1B5E20",
    "danger": "#C62828",
    "danger_hover": "#8E0000",
    "warn": "#B45309",
    "border": "#D9DEE3",
}
LOG_COLORS = {
    "info": C["text"],
    "ok": C["ok"],
    "warn": C["warn"],
    "error": C["error"],
}
LINK_COLOR = C["link"]


def _dwm_set_dark(hwnd, attribute, value):
    """Làm tối thanh tiêu đề Windows; máy không hỗ trợ sẽ ném lỗi để bỏ qua êm."""
    ctypes.windll.dwmapi.DwmSetWindowAttribute(
        ctypes.c_void_p(hwnd), ctypes.c_uint(attribute),
        ctypes.byref(ctypes.c_int(value)), ctypes.sizeof(ctypes.c_int))


class App:
    def __init__(self):
        self.config = app_config.load_config()
        # File log của lần mở Tool này; tự dọn log cũ hơn số ngày giữ
        self.file_log = log_file.open_session(app_config.APP_DIR, self.config["log_retention_days"])
        self.stop_event = threading.Event()
        self.running = False
        self._last_status = None
        self._branches = []
        self._current_title = ""
        self._link_seq = 0
        self.session = session.DriverSession(
            lambda: kiotviet.connect_driver(self.config["debug_port"], focus=False),
            prepare=kiotviet.select_kiotviet_tab)

        # Giao diện tối: khoá thang đo để giữ đúng bố cục pixel hiện có.
        # Tắt tự-scaling DPI của CTk, nếu không nó tự đổi kích thước cửa sổ sau khi khởi động.
        ctk.deactivate_automatic_dpi_awareness()
        ctk.set_appearance_mode("dark")
        self.root = ctk.CTk()
        self.root.title(version.window_title())
        # Đặt kích thước qua wm_* của Tk: lớp override geometry/minsize của CTk bị nhiễu
        # bởi sự kiện <Configure> của widget con nên tự đổi lại kích thước cửa sổ.
        self.root.wm_geometry("920x740")
        self.root.wm_minsize(780, 620)
        self.root.configure(fg_color=C["bg"])
        self._setup_style()
        self._build_ui()
        self._center_window()
        self._refresh_chrome_status()
        self._refresh_cell_hint()
        self.log(f"Phiên bản {version.build_text()}")
        self.hwnd = notifier.top_window_handle(self.root)
        self.monitor = status_monitor.StatusMonitor(
            self.session,
            self._on_status_update,
            interval=float(self.config["status_interval"]),
            available=lambda: session.debug_port_available(self.config["debug_port"]),
        )
        self.monitor.start()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._enable_dark_title_bar()

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
        style.configure("Footer.TLabel", background=COLORS["bg"], foreground=COLORS["muted"], font=(FONT, 8))

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

    def _enable_dark_title_bar(self):
        """Làm tối thanh tiêu đề Windows; máy không hỗ trợ thì bỏ qua êm."""
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()
            ui_theme.enable_dark_title_bar(
                hwnd, lambda attribute, value: _dwm_set_dark(hwnd, attribute, value))
        except Exception:
            pass

    def _card(self, parent, title, step):
        outer = ctk.CTkFrame(parent, fg_color=C["surface"], corner_radius=RADIUS,
                             border_width=1, border_color=C["border"])
        outer.pack(fill=tk.X, pady=(0, 10))
        inner = ctk.CTkFrame(outer, fg_color="transparent")
        inner.pack(fill=tk.BOTH, expand=True, padx=14, pady=10)
        ctk.CTkLabel(inner, text=f"{step}  {title}", text_color=C["text"], font=(FONT, 11, "bold"),
                     fg_color="transparent").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))
        inner.columnconfigure(1, weight=1)
        return inner

    def _build_ui(self):
        main = ctk.CTkFrame(self.root, fg_color=C["bg"], corner_radius=0)
        main.pack(fill=tk.BOTH, expand=True, padx=18, pady=14)

        # Tiêu đề
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill=tk.X, pady=(0, 12))
        header_top = ctk.CTkFrame(header, fg_color="transparent")
        header_top.pack(fill=tk.X)
        self.lbl_title = ctk.CTkLabel(header_top, text="Auto Robot KiotViet", text_color=C["text"],
                                      font=(FONT, 17, "bold"), fg_color="transparent")
        self.lbl_title.pack(side=tk.LEFT, anchor="w")
        self.btn_settings = ctk.CTkButton(
            header_top, text="⚙  Cài đặt", command=self.open_settings, fg_color=C["surface_alt"],
            hover_color=C["hover"], text_color=C["text"], corner_radius=RADIUS,
            border_width=1, border_color=C["border"])
        self.btn_settings.pack(side=tk.RIGHT, anchor="n")
        ctk.CTkLabel(
            header, text="Nhập số lượng cân chỉnh từ Excel vào màn hình kiểm kho KiotViet. "
                         "Luôn kiểm tra lại trước khi lưu phiếu.",
            text_color=C["muted"], font=(FONT, 10), fg_color="transparent",
        ).pack(anchor="w")

        # Thanh trạng thái KiotViet (chỉ đọc, cập nhật định kỳ)
        status_bar = ctk.CTkFrame(main, fg_color=C["surface"], corner_radius=RADIUS,
                                  border_width=1, border_color=C["border"])
        status_bar.pack(fill=tk.X, pady=(0, 10))
        status_inner = ctk.CTkFrame(status_bar, fg_color="transparent")
        status_inner.pack(fill=tk.X, padx=12, pady=6)
        self.lbl_status_dot = ctk.CTkLabel(
            status_inner, text="●", text_color=status_monitor.CONNECTION_COLORS["disconnected"],
            font=(FONT, 12), fg_color="transparent")
        self.lbl_status_dot.pack(side=tk.LEFT)
        self.lbl_status = ctk.CTkLabel(
            status_inner, text="Chưa kết nối Chrome Debug", text_color=C["text"],
            font=(FONT, 9), fg_color="transparent", anchor="w")
        self.lbl_status.pack(side=tk.LEFT, padx=(6, 0))
        self.btn_status_chrome = ctk.CTkButton(
            status_inner, text="Mở Chrome Debug", command=self.open_chrome_debug,
            fg_color=C["surface_alt"], hover_color=C["hover"], text_color=C["text"],
            corner_radius=8, border_width=1, border_color=C["border"], height=28, width=132)
        self.btn_show_chrome = ctk.CTkButton(
            status_inner, text="Hiện Chrome", command=self.show_chrome,
            fg_color=C["surface_alt"], hover_color=C["hover"], text_color=C["text"],
            corner_radius=8, border_width=1, border_color=C["border"], height=28, width=110)

        # Bước 1: Chrome
        card1 = self._card(main, "Mở Chrome Debug", "①")
        self.lbl_chrome = ctk.CTkLabel(card1, text="", text_color=C["muted"], font=(FONT, 9),
                                       fg_color="transparent", anchor="w")
        self.lbl_chrome.grid(row=1, column=0, columnspan=3, sticky="w")
        ctk.CTkButton(card1, text="Mở Chrome Debug", command=self.open_chrome_debug,
                      fg_color=C["surface_alt"], hover_color=C["hover"], text_color=C["text"],
                      corner_radius=8, border_width=1, border_color=C["border"]).grid(
            row=1, column=3, sticky="e", padx=(10, 0))
        ctk.CTkLabel(
            card1, text="Sau khi Chrome mở, đăng nhập KiotViet và vào màn hình kiểm kho có ô tìm kiếm sản phẩm.",
            text_color=C["muted"], font=(FONT, 9), fg_color="transparent", justify="left",
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # Bước 2: Excel
        card2 = self._card(main, "Chọn dữ liệu Excel", "②")
        ctk.CTkLabel(card2, text="File Excel", text_color=C["text"], fg_color="transparent").grid(
            row=1, column=0, sticky="w", pady=4)
        self.var_file = tk.StringVar(value=self.config["file_path"])
        ctk.CTkEntry(card2, textvariable=self.var_file, fg_color=C["surface_alt"],
                     border_color=C["border"], text_color=C["text"], corner_radius=8).grid(
            row=1, column=1, columnspan=2, sticky="ew", padx=8)
        ctk.CTkButton(card2, text="Chọn file…", command=self.choose_file, fg_color=C["surface_alt"],
                      hover_color=C["hover"], text_color=C["text"], corner_radius=8,
                      border_width=1, border_color=C["border"]).grid(row=1, column=3, padx=(0, 0))

        ctk.CTkLabel(card2, text="Tên sheet", text_color=C["text"], fg_color="transparent").grid(
            row=2, column=0, sticky="w", pady=4)
        self.var_sheet = tk.StringVar(value=self.config["sheet_name"])
        ctk.CTkEntry(card2, textvariable=self.var_sheet, fg_color=C["surface_alt"],
                     border_color=C["border"], text_color=C["text"], corner_radius=8).grid(
            row=2, column=1, columnspan=2, sticky="ew", padx=8)
        ctk.CTkLabel(card2, text="Để trống = sheet đang active", text_color=C["muted"], font=(FONT, 9),
                     fg_color="transparent").grid(row=2, column=3, sticky="w", padx=(10, 0))

        ctk.CTkLabel(card2, text="Ô bắt đầu", text_color=C["text"], fg_color="transparent").grid(
            row=3, column=0, sticky="w", pady=4)
        self.var_start = tk.StringVar(value=self.config["start_cell"])
        self.var_start.trace_add("write", lambda *_: self._refresh_cell_hint())
        ctk.CTkEntry(card2, textvariable=self.var_start, width=80, fg_color=C["surface_alt"],
                     border_color=C["border"], text_color=C["text"], corner_radius=8).grid(
            row=3, column=1, sticky="w", padx=8)
        self.lbl_cell_hint = ctk.CTkLabel(card2, text="", text_color=C["muted"], font=(FONT, 9),
                                          fg_color="transparent", anchor="w")
        self.lbl_cell_hint.grid(row=3, column=2, columnspan=2, sticky="w")

        ctk.CTkLabel(card2, text="Cơ sở theo file", text_color=C["text"], fg_color="transparent").grid(
            row=4, column=0, sticky="w", pady=4)
        self.var_branch = tk.StringVar()
        self.combo_branch = ctk.CTkComboBox(
            card2, variable=self.var_branch, values=[], command=lambda _value: self._on_branch_selected(),
            fg_color=C["surface_alt"], border_color=C["border"], text_color=C["text"],
            button_color=C["surface_alt"], button_hover_color=C["hover"], dropdown_fg_color=C["surface"],
            dropdown_hover_color=C["hover"], dropdown_text_color=C["text"], corner_radius=8)
        self.combo_branch.grid(row=4, column=1, sticky="ew", padx=8)
        self.btn_change_branch = ctk.CTkButton(
            card2, text="Đổi trên web", command=self.change_branch_clicked, fg_color=C["surface_alt"],
            hover_color=C["hover"], text_color=C["text"], corner_radius=8,
            border_width=1, border_color=C["border"])
        self.btn_change_branch.grid(row=4, column=2, sticky="w")
        self.lbl_branch = ctk.CTkLabel(card2, text="Chọn file Excel để nhận cơ sở.", text_color=C["muted"],
                                       font=(FONT, 9), fg_color="transparent", anchor="w")
        self.lbl_branch.grid(row=4, column=3, sticky="w", padx=(10, 0))

        # Bước 3: Chạy
        card3 = self._card(main, "Kiểm tra và chạy", "③")
        ctk.CTkLabel(card3, text="Chờ kết quả tối đa (giây)", text_color=C["text"],
                     fg_color="transparent").grid(row=1, column=0, sticky="w", pady=4)
        self.combo_wait = ctk.CTkComboBox(
            card3, values=["3.0", "4.0", "5.0", "7.0", "10.0"], width=100, fg_color=C["surface_alt"],
            border_color=C["border"], text_color=C["text"], button_color=C["surface_alt"],
            button_hover_color=C["hover"], dropdown_fg_color=C["surface"],
            dropdown_hover_color=C["hover"], dropdown_text_color=C["text"], corner_radius=8)
        self.combo_wait.set(str(self.config["wait_seconds"]))
        self.combo_wait.grid(row=1, column=1, sticky="w", padx=8)

        buttons = ctk.CTkFrame(card3, fg_color="transparent")
        buttons.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 4))
        self.btn_check = ctk.CTkButton(
            buttons, text="Kiểm tra file", command=self.check_file, fg_color=C["surface_alt"],
            hover_color=C["hover"], text_color=C["text"], corner_radius=8,
            border_width=1, border_color=C["border"])
        self.btn_check.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_run = ctk.CTkButton(
            buttons, text="▶  CHẠY TỰ ĐỘNG", command=self.start_run, fg_color=C["accent"],
            hover_color=C["accent_hover"], text_color=C["on_accent"], font=(FONT, 11, "bold"),
            corner_radius=8)
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_stop = ctk.CTkButton(
            buttons, text="■  DỪNG", command=self.request_stop, fg_color=C["error"],
            hover_color=C["bad_hover"], text_color=C["on_accent"], font=(FONT, 11, "bold"),
            corner_radius=8, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT)

        progress_row = ctk.CTkFrame(card3, fg_color="transparent")
        progress_row.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        self.progress_bar = ctk.CTkProgressBar(progress_row, progress_color=C["accent"],
                                               fg_color=C["surface_alt"], corner_radius=6)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar.set(0)
        self.lbl_progress = ctk.CTkLabel(progress_row, text="", text_color=C["text"], width=70,
                                         anchor="e", fg_color="transparent")
        self.lbl_progress.pack(side=tk.LEFT, padx=(8, 0))

        # Nhật ký
        log_header = ctk.CTkFrame(main, fg_color="transparent")
        log_header.pack(fill=tk.X)
        ctk.CTkLabel(log_header, text="Nhật ký", text_color=C["text"], font=(FONT, 11, "bold"),
                     fg_color="transparent").pack(side=tk.LEFT)
        ctk.CTkButton(log_header, text="Xóa", command=self.clear_log, fg_color=C["surface_alt"],
                      hover_color=C["hover"], text_color=C["text"], corner_radius=8, height=26,
                      border_width=1, border_color=C["border"]).pack(side=tk.RIGHT)
        ctk.CTkButton(log_header, text="Sao chép tất cả", command=self.copy_all_log,
                      fg_color=C["surface_alt"], hover_color=C["hover"], text_color=C["text"],
                      corner_radius=8, height=26, border_width=1, border_color=C["border"]).pack(
            side=tk.RIGHT, padx=6)

        log_box = ctk.CTkFrame(main, fg_color="transparent")
        log_box.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self.txt_log = tk.Text(
            log_box, height=10, wrap=tk.WORD, bg=C["surface"], fg=LOG_COLORS["info"],
            font=("Consolas", 10), relief=tk.FLAT, highlightbackground=C["border"],
            highlightthickness=1, insertbackground=C["text"], selectbackground=C["hover"],
            padx=8, pady=6,
        )
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_scroll = ctk.CTkScrollbar(
            log_box, command=self.txt_log.yview, fg_color=C["surface"], button_color=C["border"],
            button_hover_color=C["hover"])
        self.log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.configure(yscrollcommand=self.log_scroll.set)
        for level, color in LOG_COLORS.items():
            self.txt_log.tag_config(level, foreground=color)
        self.txt_log.config(state=tk.DISABLED, insertwidth=0, takefocus=1)
        self.txt_log.bind("<Control-c>", self._copy_selected_log)
        self.txt_log.bind("<Control-a>", self._select_all_log)
        self.txt_log.bind("<Button-1>", lambda event: self.txt_log.focus_set())

        # Đặt ở đáy và pack trước khung nhật ký, để khi cửa sổ thấp thì nhật ký co lại chứ không che mất dòng này
        ttk.Label(main, text=version.build_text(), style="Footer.TLabel").pack(
            side=tk.BOTTOM, anchor="e", pady=(6, 0), before=log_header)

    def _center_window(self):
        self.root.update_idletasks()
        # Cao vừa đủ cho toàn bộ giao diện (gồm ~10 dòng nhật ký), nhưng không vượt quá màn hình.
        # CTk trả winfo_width() mặc định (200) khi cửa sổ chưa hiện, nên lấy theo bề rộng đặt trước.
        width = max(920, self.root.winfo_reqwidth())
        height = min(max(740, self.root.winfo_reqheight() + 20), self.root.winfo_screenheight() - 80)
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 3
        self.root.wm_geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")

    def _refresh_cell_hint(self):
        try:
            start_cell, _, quantity_column, start_row = excel_reader.parse_start_cell(self.var_start.get())
            self.lbl_cell_hint.configure(
                text=f"Mã hàng đọc từ {start_cell}, số lượng đọc ở {quantity_column}{start_row} trở xuống")
        except ValueError:
            self.lbl_cell_hint.configure(text="Ô bắt đầu chưa hợp lệ, ví dụ: B4")

    def _refresh_chrome_status(self):
        path = app_config.find_chrome(self.config["chrome_path"])
        if path:
            self.lbl_chrome.configure(text=f"Chrome: {path}  ·  cổng debug {self.config['debug_port']}")
        else:
            self.lbl_chrome.configure(text="Chưa tìm thấy Chrome. Bấm nút bên phải để chọn chrome.exe.")

    def _on_status_update(self, status):
        """Nhận trạng thái từ luồng nền và chuyển về main thread của Tk."""
        try:
            self.root.after(0, lambda: self._apply_status(status))
        except (tk.TclError, RuntimeError):
            pass

    def _apply_status(self, status, running=False):
        view = status_monitor.describe_status(status, running)
        self.lbl_status.configure(text=view["text"])
        self.lbl_status_dot.configure(
            text_color=status_monitor.CONNECTION_COLORS["connected" if view["connected"] else "disconnected"])
        if view["can_open_chrome"]:
            self.btn_show_chrome.pack_forget()
            self.btn_status_chrome.pack(side=tk.RIGHT)
        else:
            self.btn_status_chrome.pack_forget()
            self.btn_show_chrome.pack(side=tk.RIGHT)
        if status is not None:
            self._last_status = status
            self._update_branches(status.get("branches") or [])
        self._update_branch_hint()

    def open_settings(self):
        if self.running:
            messagebox.showwarning("Đang chạy", "Không thể đổi cài đặt khi Tool đang chạy.")
            return
        self._save_config()

        dialog = settings_dialog.SettingsDialog(self.root, self.config)
        self.root.wait_window(dialog)
        if dialog.result is None:
            return

        old_port = self.config.get("debug_port")
        self.config.update(dialog.result)
        # Cập nhật ô trên màn hình chính trước: _save_config() đọc giá trị từ ô này
        self.combo_wait.set(str(self.config["wait_seconds"]))
        self._save_config()
        self._refresh_chrome_status()
        self.monitor.set_interval(self.config["status_interval"])
        log_file.purge_old(log_file.log_dir(app_config.APP_DIR), self.config["log_retention_days"])
        if str(old_port) != str(self.config["debug_port"]):
            # Ngắt phiên cũ phải chờ lock và chromedriver: chạy nền để cửa sổ không bị treo
            threading.Thread(target=self.session.close, daemon=True).start()
            self.log("Đã đổi cổng debug; lần kết nối Chrome tiếp theo sẽ dùng cổng mới.", "warn")
        self.log("Đã lưu cài đặt.", "ok")
        self.monitor.refresh()

    # ------------------------------------------------------------------
    # Tiện ích an toàn đa luồng
    # ------------------------------------------------------------------
    def _on_main_thread(self):
        return threading.current_thread() is threading.main_thread()

    def log(self, message, level="info", code=None):
        """Ghi nhật ký ra màn hình và file log.

        Nếu có `code`, mã hàng trong dòng được gạch chân và bấm được để tìm trên phiếu.
        """
        self.file_log.write(message)

        def apply():
            self.txt_log.config(state=tk.NORMAL)
            anchor = self.txt_log.index(tk.END + "-1c")
            self.txt_log.insert(tk.END, message + "\n", level)
            offset = message.find(code) if code else -1
            if offset >= 0:
                self._link_seq += 1
                tag = f"link{self._link_seq}"
                start = f"{anchor}+{offset}c"
                end = f"{start}+{len(code)}c"
                self.txt_log.tag_add(tag, start, end)
                self.txt_log.tag_config(tag, foreground=LINK_COLOR, underline=True)
                self.txt_log.tag_bind(tag, "<Button-1>", lambda event, c=code: self.reveal_code(c))
                self.txt_log.tag_bind(tag, "<Enter>", lambda event: self.txt_log.config(cursor="hand2"))
                self.txt_log.tag_bind(tag, "<Leave>", lambda event: self.txt_log.config(cursor=""))
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
            self.progress_bar.set(done / total if total else 0)
            self.lbl_progress.configure(text=f"{done}/{total}" if total else "")

        if self._on_main_thread():
            apply()
        else:
            self.root.after(0, apply)

    def _set_running(self, running, can_stop=True):
        self.running = running
        self.btn_run.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_check.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_settings.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_change_branch.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_show_chrome.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_stop.configure(state=tk.NORMAL if running and can_stop else tk.DISABLED)
        if running:
            self._apply_status(None, running=True)
        else:
            self._apply_status(self._last_status)
            self.monitor.refresh()

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
    # Tìm mã trên phiếu
    # ------------------------------------------------------------------
    def reveal_code(self, code):
        """Bấm mã hàng trong nhật ký: chuyển phiếu KiotViet tới dòng chứa mã đó."""
        if self.running:
            self.log("Tool đang chạy, chờ xong rồi bấm lại mã để tìm trên phiếu.", "warn")
            return
        threading.Thread(target=self._reveal_worker, args=(code,), daemon=True).start()

    def _reveal_worker(self, code):
        try:
            with self.session.exclusive() as driver:
                if driver is None:
                    self.log(f"Chưa kết nối được Chrome Debug nên chưa tìm được mã {code}.", "warn")
                    return
                kiotviet.bring_to_front(driver)
                stt = kiotviet.reveal_item(driver, code)
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không tìm được mã {code}: {first_line or type(exc).__name__}", "error")
            return

        if stt:
            self.log(f"Đã chuyển tới mã {code} (STT {stt} trên KiotViet).", "ok")
        else:
            self.log(f"Không thấy mã {code} trên phiếu đang mở.", "warn")

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
        # Không được chờ gì ở đây: lần chạy/thanh trạng thái có thể đang giữ phiên Selenium
        # rất lâu, chờ lock hay driver.quit() sẽ làm cửa sổ "Not Responding".
        if self.running and not messagebox.askyesno(
                "Tool đang chạy", "Tool đang chạy. Dừng lại và thoát?", icon="warning"):
            return
        self.stop_event.set()
        self._save_config()
        self.monitor.stop(wait=False)
        self.file_log.close()
        self.session.terminate()
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

        # Ngắt phiên cũ và mở Chrome có thể chờ chromedriver vài chục giây; chạy nền để cửa sổ
        # Tool không bị treo.
        threading.Thread(target=self._open_chrome_worker, args=(chrome_path,), daemon=True).start()

    def _open_chrome_worker(self, chrome_path):
        port = self.config["debug_port"]
        try:
            # Dừng chromedriver trước: driver.quit() sẽ không phải chờ driver tự thoát (tối đa 30 giây)
            subprocess.run(["taskkill", "/f", "/im", "chromedriver.exe"], capture_output=True)
            self.session.close()
            subprocess.Popen([
                chrome_path,
                f"--remote-debugging-port={port}",
                f"--user-data-dir={self.config['user_data_dir']}",
            ])
        except OSError as exc:
            self.notify("showerror", "Lỗi", f"Không thể mở Chrome Debug:\n{exc}")
            return
        self.log(f"Đã mở Chrome Debug (cổng {port}).", "ok")
        self.monitor.refresh()

    def show_chrome(self):
        """Đưa cửa sổ Chrome Debug đang kết nối lên trên cùng."""
        threading.Thread(target=self._show_chrome_worker, daemon=True).start()

    def _show_chrome_worker(self):
        with self.session.exclusive(blocking=False) as driver:
            if driver is None:
                return  # đang chạy/đang bận: Tool tự đưa Chrome lên trong lúc chạy
            kiotviet.bring_to_front(driver)

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
                self.lbl_title.configure(text="[Lỗi đọc file Excel]")
                self.root.title(version.window_title())
                messagebox.showerror("Lỗi đọc Excel", "Không thể đọc thông tin từ file Excel đã chọn.")
                return

        c2_value = sheet["C2"].value
        title = str(c2_value) if c2_value else "[Ô C2 trống]"
        self.lbl_title.configure(text=title)
        self.root.title(version.window_title(str(c2_value).strip() if c2_value else ""))
        self._current_title = str(c2_value).strip() if c2_value else ""
        self._update_branches(self._branches)
        self._update_branch_hint()

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
        self.btn_stop.configure(state=tk.DISABLED)
        self.log("Đã nhận yêu cầu dừng, Tool sẽ dừng sau khi xong mã hàng hiện tại...", "warn")

    def _worker(self, items, start_cell, quantity_column, result_timeout, title):
        success = False
        try:
            with self.session.exclusive() as driver:
                if driver is None:
                    self.notify("showerror", "Lỗi kết nối",
                                "Không thể kết nối với Chrome Debug. Hãy bấm 'Mở Chrome Debug' trước.")
                    return
                kiotviet.bring_to_front(driver)

                self.log("----------------------------------------")
                if not self._confirm_branch(driver, title):
                    self.log("Đã hủy chạy vì cơ sở không khớp.", "warn")
                    return

                try:
                    page, discarded_draft = kiotviet.ensure_stocktake_page(driver)
                except Exception as exc:
                    self.notify("showerror", "Chưa ở màn hình Kiểm kho", str(exc).split("\n")[0])
                    return
                if page == "created":
                    self.log("Đang ở danh sách kiểm kho, tạo phiếu mới.", "ok")
                if discarded_draft:
                    self.log("Đã chọn 'Bỏ qua' bản nháp cũ, bắt đầu phiếu mới.", "warn")

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
                    verified = self._verify_stocktake(driver, stats.succeeded)
                    self._fill_note(driver)
                    success = verified and not (stats.skipped or stats.error)
        finally:
            self.root.after(0, lambda: self._set_running(False))
            notifier.notify_done(
                success,
                sound=self.config.get("notify_sound", True),
                flash=self.config.get("notify_flash", True),
                hwnd=self.hwnd,
            )

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
            with self.session.exclusive() as driver:
                if driver is None:
                    self.log("Chưa kết nối được Chrome Debug/KiotViet nên chưa tự đổi cơ sở.", "warn")
                    return
                try:
                    current = kiotviet.get_current_branch(driver)
                    branches = kiotviet.get_branches(driver)
                except Exception:
                    self.session.invalidate()
                    self.log("Chưa kết nối được Chrome Debug/KiotViet nên chưa tự đổi cơ sở.", "warn")
                    return
                if not branches:
                    self.log("Không đọc được danh sách cơ sở trên web nên chưa tự đổi cơ sở.", "warn")
                    return

                branch, reason = branch_match.resolve_branch(
                    title, branches, self.config.get("branch_overrides", {}))
                if branch is None:
                    if reason == "ambiguous":
                        self.log("Tiêu đề file khớp nhiều cơ sở, hãy chọn tay.", "warn")
                    else:
                        self.log("Không xác định được cơ sở từ file, hãy chọn tay.", "warn")
                    return

                self._show_branch_choice(branch)
                if current and str(branch[0]) == str(current[0]):
                    self.log(f"Cơ sở đã đúng: {current[1]}.", "ok")
                    return

                if kiotviet.get_page_state(driver) == kiotviet.PAGE_EDIT and kiotviet.count_item_rows(driver):
                    self.log("Phiếu đã có hàng, không đổi cơ sở (tránh mất dữ liệu).", "warn")
                    return

                working = current[1] if current else "không rõ"
                self.log(f"Đang đổi cơ sở: {working} → {branch[1]}...")
                kiotviet.change_branch(driver, branch[0])
                self.log(f"Đã đổi cơ sở làm việc sang {branch[1]}.", "ok")
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không đổi được cơ sở: {first_line or type(exc).__name__}", "error")
        finally:
            self.root.after(0, lambda: self._set_running(False))

    def change_branch_clicked(self):
        """Đổi cơ sở trên web sang lựa chọn trong danh sách Cơ sở theo file."""
        if self.running:
            return
        branch = self._selected_branch()
        if branch is None:
            messagebox.showwarning("Cơ sở", "Hãy chọn cơ sở trong danh sách trước.")
            return
        self._set_running(True, can_stop=False)
        threading.Thread(target=self._change_branch_worker, args=(branch,), daemon=True).start()

    def _change_branch_worker(self, branch):
        try:
            with self.session.exclusive() as driver:
                if driver is None:
                    self.log("Chưa kết nối được Chrome/KiotViet, không đổi được cơ sở.", "warn")
                    return
                try:
                    current = kiotviet.get_current_branch(driver)
                except Exception:
                    self.session.invalidate()
                    self.log("Chưa kết nối được Chrome/KiotViet, không đổi được cơ sở.", "warn")
                    return

                if current and str(branch[0]) == str(current[0]):
                    self.log(f"Cơ sở đã đúng: {current[1]}.", "ok")
                    return

                if kiotviet.get_page_state(driver) == kiotviet.PAGE_EDIT and kiotviet.count_item_rows(driver):
                    self.log("Phiếu đã có hàng, không đổi cơ sở (tránh mất dữ liệu).", "warn")
                    return

                working = current[1] if current else "không rõ"
                self.log(f"Đang đổi cơ sở: {working} → {branch[1]}...")
                kiotviet.change_branch(driver, branch[0])
                self.log(f"Đã đổi cơ sở sang {branch[1]}.", "ok")
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không đổi được cơ sở: {first_line or type(exc).__name__}", "error")
        finally:
            self.root.after(0, lambda: self._set_running(False))

    def _update_branches(self, branches):
        branches = [tuple(b) for b in branches]
        if not branches:
            return
        if branches != self._branches:
            self._branches = branches
            self.combo_branch.configure(values=[name for _, name in branches])

        # Chọn sẵn cơ sở theo file (ưu tiên lựa chọn tay đã lưu) khi danh sách vừa có
        if self._current_title and self._selected_branch() is None:
            branch, _ = branch_match.resolve_branch(
                self._current_title, branches, self.config.get("branch_overrides", {}))
            if branch:
                self.var_branch.set(branch[1])

    def _selected_branch(self):
        name = self.var_branch.get().strip()
        if not name:
            return None
        return next((branch for branch in self._branches if branch[1] == name), None)

    def _on_branch_selected(self, event=None):
        branch = self._selected_branch()
        if branch and self._current_title:
            overrides = self.config.setdefault("branch_overrides", {})
            overrides[branch_match.normalize(self._current_title)] = branch[0]
            self._save_config()
            self.log(f"Đã ghi nhớ cơ sở cho file này: {branch[1]}.", "ok")
        self._update_branch_hint()

    def _show_branch_choice(self, branch):
        def apply():
            if branch[1] in self.combo_branch.cget("values"):
                self.var_branch.set(branch[1])
            self._update_branch_hint()
        self.root.after(0, apply)

    def _update_branch_hint(self):
        selected = self._selected_branch()
        if not self._current_title:
            self.lbl_branch.configure(text="Chọn file Excel để nhận cơ sở.", text_color=C["muted"])
            return
        if selected is None:
            if self._branches:
                self.lbl_branch.configure(text="⚠ Không nhận ra cơ sở từ C2, hãy chọn", text_color=C["warn"])
            else:
                self.lbl_branch.configure(text="Chưa đọc được danh sách cơ sở trên web", text_color=C["muted"])
            return

        web = (self._last_status or {}).get("branch")
        if not web:
            self.lbl_branch.configure(text="Chưa kết nối web", text_color=C["muted"])
        elif str(web[0]) == str(selected[0]):
            self.lbl_branch.configure(text="✓ Khớp với web", text_color=C["ok"])
        else:
            self.lbl_branch.configure(text=f"⚠ Web đang ở {web[1]}", text_color=C["warn"])

    def _confirm_branch(self, driver, title):
        """So cơ sở đang chọn trong danh sách với cơ sở làm việc trên web."""
        selected = self._selected_branch()
        if not selected and not title:
            self.log("Chưa có cơ sở để kiểm tra.", "warn")
            return True
        try:
            current = kiotviet.get_current_branch(driver)
        except Exception:
            current = None
        if not current:
            self.log("Không đọc được cơ sở trên web, bỏ qua kiểm tra.", "warn")
            return True

        if selected:
            if str(selected[0]) == str(current[0]):
                self.log(f"Cơ sở khớp: {current[1]}.", "ok")
                return True
            file_label = selected[1]
        else:
            if branch_match.is_same_branch(title, current[1]):
                self.log(f"Cơ sở khớp: {current[1]}.", "ok")
                return True
            file_label = title

        self.log(f"Cơ sở có thể không khớp: File {file_label} | Web {current[1]}", "warn")
        return self.ask_yes_no(
            "Cơ sở không khớp",
            f"Cơ sở theo file: {file_label}\nCơ sở đang làm việc trên web: {current[1]}\n\n"
            "Hai cơ sở có vẻ không khớp. Vẫn tiếp tục chạy?",
        )

    def _verify_stocktake(self, driver, filled_items):
        """Đọc lại SL lệch trên từng trang của phiếu và so với Excel cho các mã đã điền thành công.

        Trả về True nếu mọi mã đã điền đều khớp (hoặc không có mã nào để kiểm tra).
        """
        if not filled_items:
            return True
        self.log("--- KIỂM TRA LẠI SL LỆCH CÁC MÃ ĐÃ ĐIỀN ---")
        try:
            web_rows = kiotviet.read_stocktake_rows(driver)
        except Exception as exc:
            first_line = str(exc).split("\n")[0]
            self.log(f"Không kiểm tra lại được phiếu: {first_line or type(exc).__name__}", "error")
            return False

        report = verify.compare_adjustments(filled_items, web_rows)
        for level, text, code in verify.describe_report(report, kiotviet.format_qty):
            self.log(text, level, code)
        return report.ok

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
