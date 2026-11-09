"""Cửa sổ Cài đặt: sửa các thiết lập đang lưu trong config.json."""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import config as app_config

ERROR_COLOR = "#C62828"


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, current_config):
        super().__init__(parent)
        self.result = None
        self.config = dict(current_config)

        self.title("Cài đặt")
        self.transient(parent)
        self.resizable(False, False)
        self._style = ttk.Style(self)
        self._style.configure("Error.TLabel", foreground=ERROR_COLOR)
        self._style.configure("Hint.TLabel", foreground="#6B7785")

        self._build()
        self._fill(self.config)
        self._center(parent)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.grab_set()

    def _build(self):
        body = ttk.Frame(self, padding=(18, 14))
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(1, weight=1)

        self.var_chrome = tk.StringVar()
        self.var_port = tk.StringVar()
        self.var_profile = tk.StringVar()
        self.var_wait = tk.StringVar()
        self.var_note = tk.StringVar()
        self.var_sound = tk.BooleanVar()
        self.var_flash = tk.BooleanVar()
        self.var_interval = tk.StringVar()
        self.var_retention = tk.StringVar()
        self._errors = {}

        row = 0
        ttk.Label(body, text="Đường dẫn chrome.exe").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.var_chrome, width=48).grid(row=row, column=1, sticky="ew", padx=6)
        chrome_buttons = ttk.Frame(body)
        chrome_buttons.grid(row=row, column=2, sticky="w")
        ttk.Button(chrome_buttons, text="Chọn…", command=self._choose_chrome, padding=(8, 3)).pack(side=tk.LEFT)
        ttk.Button(chrome_buttons, text="Tự dò", command=self._detect_chrome, padding=(8, 3)).pack(
            side=tk.LEFT, padx=(4, 0))
        self._error(body, row, "chrome_path")
        ttk.Label(body, text="Để trống để Tool tự dò vị trí cài đặt.", style="Hint.TLabel").grid(
            row=row + 1, column=0, columnspan=4, sticky="w")
        row += 2

        self._field(body, row, "Cổng debug", self.var_port, "debug_port", width=10)
        row += 1
        self._field(body, row, "Thư mục profile Chrome", self.var_profile, "user_data_dir", width=48)
        row += 1
        self._field(body, row, "Chờ kết quả tối đa (giây)", self.var_wait, "wait_seconds", width=10)
        row += 1
        self._field(body, row, "Ghi chú sau khi chạy", self.var_note, "note_text", width=48)
        row += 1
        self._field(body, row, "Cập nhật trạng thái mỗi (giây)", self.var_interval, "status_interval", width=10)
        row += 1
        self._field(body, row, "Giữ nhật ký (ngày)", self.var_retention, "log_retention_days", width=10)
        row += 1

        ttk.Label(body, text="Thông báo khi xong").grid(row=row, column=0, sticky="nw", pady=(8, 2))
        notify = ttk.Frame(body)
        notify.grid(row=row, column=1, columnspan=3, sticky="w", pady=(8, 2))
        ttk.Checkbutton(notify, text="Âm thanh", variable=self.var_sound).pack(side=tk.LEFT)
        ttk.Checkbutton(notify, text="Nhấp nháy taskbar", variable=self.var_flash).pack(side=tk.LEFT, padx=(14, 0))
        row += 1

        ttk.Separator(body).grid(row=row, column=0, columnspan=4, sticky="ew", pady=12)
        row += 1

        buttons = ttk.Frame(body)
        buttons.grid(row=row, column=0, columnspan=4, sticky="e")
        ttk.Button(buttons, text="Khôi phục mặc định", command=self._restore_defaults).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Hủy", command=self._cancel).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Lưu", style="Primary.TButton", command=self._save).pack(
            side=tk.LEFT, padx=(8, 0))

    def _field(self, parent, row, label, variable, key, width):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=variable, width=width).grid(row=row, column=1, sticky="w", padx=6)
        self._error(parent, row, key)

    def _error(self, parent, row, key):
        label = ttk.Label(parent, text="", style="Error.TLabel", wraplength=220)
        label.grid(row=row, column=3, sticky="w", padx=(8, 0))
        self._errors[key] = label

    def _fill(self, cfg):
        self.var_chrome.set(str(cfg.get("chrome_path") or ""))
        self.var_port.set(str(cfg.get("debug_port", "")))
        self.var_profile.set(str(cfg.get("user_data_dir") or ""))
        self.var_wait.set(str(cfg.get("wait_seconds", "")))
        self.var_note.set(str(cfg.get("note_text") or ""))
        self.var_sound.set(bool(cfg.get("notify_sound", True)))
        self.var_flash.set(bool(cfg.get("notify_flash", True)))
        self.var_interval.set(str(cfg.get("status_interval", "")))
        self.var_retention.set(str(cfg.get("log_retention_days", "")))
        self._show_errors({})

    def _collect(self):
        return {
            "chrome_path": self.var_chrome.get().strip(),
            "debug_port": self.var_port.get().strip(),
            "user_data_dir": self.var_profile.get().strip(),
            "wait_seconds": self.var_wait.get().strip(),
            "note_text": self.var_note.get(),
            "notify_sound": bool(self.var_sound.get()),
            "notify_flash": bool(self.var_flash.get()),
            "status_interval": self.var_interval.get().strip(),
            "log_retention_days": self.var_retention.get().strip(),
        }

    def _show_errors(self, errors):
        for key, label in self._errors.items():
            label.config(text=errors.get(key, ""))

    def _save(self):
        values = self._collect()
        errors = app_config.validate_config(values)
        self._show_errors(errors)
        if errors:
            return

        values["debug_port"] = int(values["debug_port"])
        values["status_interval"] = int(values["status_interval"])
        values["log_retention_days"] = int(values["log_retention_days"])
        values["wait_seconds"] = str(float(values["wait_seconds"]))
        self.result = values
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

    def _restore_defaults(self):
        self._fill(app_config.DEFAULTS)

    def _choose_chrome(self):
        path = filedialog.askopenfilename(
            parent=self, title="Chọn chrome.exe",
            filetypes=[("chrome.exe", "chrome.exe"), ("All Files", "*.*")],
            initialdir=self.var_chrome.get() or None,
        )
        if path:
            self.var_chrome.set(path)

    def _detect_chrome(self):
        path = app_config.find_chrome("")
        if path:
            self.var_chrome.set(path)
        else:
            messagebox.showinfo(
                "Tự dò Chrome", "Không tìm thấy Chrome ở các vị trí quen thuộc. Hãy bấm 'Chọn…'.", parent=self)

    def _center(self, parent):
        self.update_idletasks()
        width, height = self.winfo_width(), self.winfo_height()
        x = parent.winfo_rootx() + max((parent.winfo_width() - width) // 2, 0)
        y = parent.winfo_rooty() + max((parent.winfo_height() - height) // 3, 0)
        self.geometry(f"+{x}+{y}")
