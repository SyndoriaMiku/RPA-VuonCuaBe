"""Cửa sổ Cài đặt: sửa các thiết lập đang lưu trong config.json."""
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

import config as app_config
import ui_theme

C = ui_theme.PALETTE
FONT = ui_theme.FONT


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_config):
        super().__init__(parent)
        self.result = None
        self.config = dict(current_config)

        self.title("Cài đặt")
        self.configure(fg_color=C["bg"])
        self.transient(parent)
        self.resizable(False, False)

        self._build()
        self._fill(self.config)
        self._center(parent)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.grab_set()

    def _build(self):
        body = ctk.CTkFrame(self, fg_color=C["bg"])
        body.pack(fill=tk.BOTH, expand=True, padx=18, pady=14)
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
        ctk.CTkLabel(body, text="Đường dẫn chrome.exe", text_color=C["text"],
                     fg_color="transparent").grid(row=row, column=0, sticky="w", pady=4)
        ctk.CTkEntry(body, textvariable=self.var_chrome, width=340, fg_color=C["surface_alt"],
                     border_color=C["border"], text_color=C["text"], corner_radius=8).grid(
            row=row, column=1, sticky="ew", padx=6)
        chrome_buttons = ctk.CTkFrame(body, fg_color="transparent")
        chrome_buttons.grid(row=row, column=2, sticky="w")
        self._secondary_button(chrome_buttons, "Chọn…", self._choose_chrome).pack(side=tk.LEFT)
        self._secondary_button(chrome_buttons, "Tự dò", self._detect_chrome).pack(
            side=tk.LEFT, padx=(4, 0))
        self._error(body, row, "chrome_path")
        ctk.CTkLabel(body, text="Để trống để Tool tự dò vị trí cài đặt.", text_color=C["muted"],
                     font=(FONT, 9), fg_color="transparent").grid(
            row=row + 1, column=0, columnspan=4, sticky="w")
        row += 2

        self._field(body, row, "Cổng debug", self.var_port, "debug_port", 90)
        row += 1
        self._field(body, row, "Thư mục profile Chrome", self.var_profile, "user_data_dir", 340)
        row += 1
        self._field(body, row, "Chờ kết quả tối đa (giây)", self.var_wait, "wait_seconds", 90)
        row += 1
        self._field(body, row, "Ghi chú sau khi chạy", self.var_note, "note_text", 340)
        row += 1
        self._field(body, row, "Cập nhật trạng thái mỗi (giây)", self.var_interval, "status_interval", 90)
        row += 1
        self._field(body, row, "Giữ nhật ký (ngày)", self.var_retention, "log_retention_days", 90)
        row += 1

        ctk.CTkLabel(body, text="Thông báo khi xong", text_color=C["text"], fg_color="transparent").grid(
            row=row, column=0, sticky="nw", pady=(8, 2))
        notify = ctk.CTkFrame(body, fg_color="transparent")
        notify.grid(row=row, column=1, columnspan=3, sticky="w", pady=(8, 2))
        ctk.CTkCheckBox(notify, text="Âm thanh", variable=self.var_sound, text_color=C["text"],
                        fg_color=C["accent"], hover_color=C["accent_hover"],
                        border_color=C["border"], checkmark_color=C["on_accent"]).pack(side=tk.LEFT)
        ctk.CTkCheckBox(notify, text="Nhấp nháy taskbar", variable=self.var_flash, text_color=C["text"],
                        fg_color=C["accent"], hover_color=C["accent_hover"],
                        border_color=C["border"], checkmark_color=C["on_accent"]).pack(
            side=tk.LEFT, padx=(14, 0))
        row += 1

        ctk.CTkFrame(body, height=1, fg_color=C["border"]).grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=12)
        row += 1

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.grid(row=row, column=0, columnspan=4, sticky="e")
        self._secondary_button(buttons, "Khôi phục mặc định", self._restore_defaults).pack(side=tk.LEFT)
        self._secondary_button(buttons, "Hủy", self._cancel).pack(side=tk.LEFT, padx=(8, 0))
        ctk.CTkButton(buttons, text="Lưu", command=self._save, fg_color=C["accent"],
                      hover_color=C["accent_hover"], text_color=C["on_accent"], font=(FONT, 11, "bold"),
                      corner_radius=8).pack(side=tk.LEFT, padx=(8, 0))

    def _secondary_button(self, parent, text, command):
        return ctk.CTkButton(parent, text=text, command=command, fg_color=C["surface_alt"],
                             hover_color=C["hover"], text_color=C["text"], corner_radius=8,
                             border_width=1, border_color=C["border"])

    def _field(self, parent, row, label, variable, key, width):
        ctk.CTkLabel(parent, text=label, text_color=C["text"], fg_color="transparent").grid(
            row=row, column=0, sticky="w", pady=4)
        ctk.CTkEntry(parent, textvariable=variable, width=width, fg_color=C["surface_alt"],
                     border_color=C["border"], text_color=C["text"], corner_radius=8).grid(
            row=row, column=1, sticky="w", padx=6)
        self._error(parent, row, key)

    def _error(self, parent, row, key):
        label = ctk.CTkLabel(parent, text="", text_color=C["error"], font=(FONT, 9), wraplength=220,
                             fg_color="transparent", anchor="w")
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
            label.configure(text=errors.get(key, ""))

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
        # Dùng bề rộng yêu cầu: CTk trả kích thước mặc định khi cửa sổ chưa hiện.
        width, height = self.winfo_reqwidth(), self.winfo_reqheight()
        x = parent.winfo_rootx() + max((parent.winfo_width() - width) // 2, 0)
        y = parent.winfo_rooty() + max((parent.winfo_height() - height) // 3, 0)
        self.wm_geometry(f"+{x}+{y}")
