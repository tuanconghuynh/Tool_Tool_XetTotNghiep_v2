# -*- coding: utf-8 -*-
import os
import sys
import traceback
from pathlib import Path
import tkinter as tk
from engine import process_excel, EMAIL_FIELDS, validate_email_metadata
NAME_AUTHOR = "HuynhCongTuan"
APP_VERSION = "2.0"
# GIAO DIỆN PHẦN MỀM
# ============================================================

# Yêu cầu:
#   pip install customtkinter openpyxl
#
# Nếu máy chưa có CustomTkinter, phần mềm sẽ báo hướng dẫn cài đặt.

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

import subprocess
import threading
import time
from datetime import datetime


APP_TITLE = "PHẦN MỀM XÉT TỐT NGHIỆP"
APP_SUBTITLE = "HUTECH – Viện Đại học Mở HUTECH"
NAME_AUTHOR = "Huỳnh Công Tuấn - THC0091102"


def open_file_or_folder(path):
    """Mở file hoặc thư mục bằng ứng dụng mặc định của Windows."""
    path = str(path)
    if os.name == "nt":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class EmailInformationDialog(ctk.CTkToplevel if ctk else object):
    """Nhập thông tin chung trước khi xét; Hủy không xuất file."""
    def __init__(self, parent, previous=None):
        super().__init__(parent)
        self.title('Thông tin đợt tốt nghiệp và trả email')
        self.geometry('780x720')
        self.minsize(680, 500)
        self.transient(parent)
        self.result = None
        self.entries = {}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self, text='THÔNG TIN TRẢ EMAIL',
                     font=ctk.CTkFont(size=20, weight='bold')).grid(row=0, column=0, padx=24, pady=(18, 8))
        form = ctk.CTkScrollableFrame(self)
        form.grid(row=1, column=0, sticky='nsew', padx=20, pady=8)
        form.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(form, text='* Bắt buộc. Các mục hạn được phép để trống.\nThông tin sẽ áp dụng cho tất cả học viên trong file email.',
                     justify='left').grid(row=0, column=0, sticky='w', padx=12, pady=10)
        for index, (key, label, required, default) in enumerate(EMAIL_FIELDS):
            ctk.CTkLabel(form, text=label + (' *' if required else ' (không bắt buộc)'),
                         anchor='w').grid(row=index*2+1, column=0, sticky='ew', padx=12, pady=(8, 2))
            entry = ctk.CTkEntry(form, height=36)
            entry.grid(row=index*2+2, column=0, sticky='ew', padx=12, pady=(0, 4))
            entry.insert(0, (previous or {}).get(key, default))
            self.entries[key] = entry
        self.error_label = ctk.CTkLabel(self, text='', text_color='#B91C1C', wraplength=700)
        self.error_label.grid(row=2, column=0, sticky='ew', padx=20, pady=4)
        buttons = ctk.CTkFrame(self, fg_color='transparent')
        buttons.grid(row=3, column=0, pady=(8, 18))
        ctk.CTkButton(buttons, text='Hủy', command=self.destroy,
                       fg_color='#64748B').grid(row=0, column=0, padx=8)
        ctk.CTkButton(buttons, text='Tiếp tục xét', command=self.submit).grid(row=0, column=1, padx=8)
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        self.after(100, self.activate)

    def activate(self):
        if self.winfo_exists():
            self.grab_set()
            self.entries[EMAIL_FIELDS[0][0]].focus_set()

    def submit(self):
        try:
            self.result = validate_email_metadata({key: entry.get() for key, entry in self.entries.items()})
        except ValueError as exc:
            self.error_label.configure(text=str(exc))
            return
        self.destroy()


class GraduationApp(ctk.CTk if ctk else object):
    def __init__(self):
        super().__init__()

        self.title(f"{APP_TITLE} | {APP_VERSION}")
        self.geometry("1080x760")
        self.minsize(980, 700)
        self.resizable(True, True)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.selected_file = None
        self.output_dir = None
        self.result_info = None
        self.processing = False
        self.email_metadata = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.build_header()
        self.build_body()
        self.build_footer()

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------
    def build_header(self):
        header = ctk.CTkFrame(
            self,
            corner_radius=0,
            fg_color="#0B5CAD",
            height=110
        )
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        logo_box = ctk.CTkFrame(
            header,
            width=72,
            height=72,
            corner_radius=16,
            fg_color="#FFFFFF"
        )
        logo_box.grid(row=0, column=0, padx=(28, 18), pady=19)
        logo_box.grid_propagate(False)

        ctk.CTkLabel(
            logo_box,
            text="🎓",
            font=ctk.CTkFont(size=34)
        ).place(relx=0.5, rely=0.5, anchor="center")

        title_frame = ctk.CTkFrame(
            header,
            fg_color="transparent"
        )
        title_frame.grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(
            title_frame,
            text=APP_TITLE,
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color="#FFFFFF"
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text=APP_SUBTITLE,
            font=ctk.CTkFont(size=14),
            text_color="#DDEBFA"
        ).pack(anchor="w", pady=(4, 0))

        ctk.CTkLabel(
            header,
            text=f"Phiên bản: {APP_VERSION}",
            font=ctk.CTkFont(size=12),
            text_color="#DDEBFA"
        ).grid(row=0, column=2, padx=28)

    # --------------------------------------------------------
    # BODY
    # --------------------------------------------------------
    def build_body(self):
        body = ctk.CTkScrollableFrame(
            self,
            fg_color="#F5F7FA",
            corner_radius=0
        )
        body.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        body.grid_columnconfigure(0, weight=1)

        # FILE DỮ LIỆU
        file_card = ctk.CTkFrame(
            body,
            corner_radius=14,
            fg_color="#FFFFFF",
            border_width=1,
            border_color="#E2E8F0"
        )
        file_card.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 12))
        file_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            file_card,
            text="📂  FILE DỮ LIỆU",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#1E293B"
        ).grid(row=0, column=0, sticky="w", padx=22, pady=(18, 12))

        file_row = ctk.CTkFrame(file_card, fg_color="transparent")
        file_row.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 20))
        file_row.grid_columnconfigure(0, weight=1)

        self.file_entry = ctk.CTkEntry(
            file_row,
            height=42,
            corner_radius=9,
            placeholder_text="Chưa chọn file Excel..."
        )
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.file_entry.configure(state="disabled")

        self.choose_button = ctk.CTkButton(
            file_row,
            text="📁  Chọn file",
            width=135,
            height=42,
            corner_radius=9,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.choose_file
        )
        self.choose_button.grid(row=0, column=1)

        # NÚT CHẠY
        self.run_button = ctk.CTkButton(
            body,
            text="🚀  BẮT ĐẦU XÉT TỐT NGHIỆP",
            height=54,
            corner_radius=12,
            font=ctk.CTkFont(size=17, weight="bold"),
            fg_color="#0B5CAD",
            hover_color="#084B8A",
            command=self.start_process
        )
        self.run_button.grid(row=1, column=0, sticky="ew", padx=28, pady=12)

        # TRẠNG THÁI
        status_card = ctk.CTkFrame(
            body,
            corner_radius=14,
            fg_color="#FFFFFF",
            border_width=1,
            border_color="#E2E8F0"
        )
        status_card.grid(row=2, column=0, sticky="ew", padx=28, pady=12)
        status_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            status_card,
            text="⚙️  TRẠNG THÁI XỬ LÝ",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1E293B"
        ).grid(row=0, column=0, sticky="w", padx=22, pady=(17, 8))

        self.status_label = ctk.CTkLabel(
            status_card,
            text="Sẵn sàng – Vui lòng chọn file Excel.",
            font=ctk.CTkFont(size=13),
            text_color="#475569"
        )
        self.status_label.grid(row=1, column=0, sticky="w", padx=22)

        self.progress = ctk.CTkProgressBar(
            status_card,
            height=10,
            corner_radius=5,
            mode="determinate"
        )
        self.progress.grid(row=2, column=0, sticky="ew", padx=22, pady=(12, 20))
        self.progress.set(0)

        # KẾT QUẢ
        result_card = ctk.CTkFrame(
            body,
            corner_radius=14,
            fg_color="#FFFFFF",
            border_width=1,
            border_color="#E2E8F0"
        )
        result_card.grid(row=3, column=0, sticky="ew", padx=28, pady=12)
        result_card.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            result_card,
            text="📊  KẾT QUẢ XÉT",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1E293B"
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=22, pady=(17, 14))

        self.total_value = self.make_stat_box(
            result_card, 0, "👥", "TỔNG HỌC VIÊN", "—", "#334155"
        )
        self.eligible_value = self.make_stat_box(
            result_card, 1, "✅", "ĐỦ ĐIỀU KIỆN", "—", "#15803D"
        )
        self.not_eligible_value = self.make_stat_box(
            result_card, 2, "❌", "KHÔNG ĐỦ ĐIỀU KIỆN", "—", "#B91C1C"
        )

        self.rate_label = ctk.CTkLabel(
            result_card,
            text="Tỷ lệ đủ điều kiện: —    |    Tỷ lệ không đủ: —",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#475569"
        )
        self.rate_label.grid(row=2, column=0, columnspan=3, pady=(10, 20))

        # FILE ĐẦU RA
        output_card = ctk.CTkFrame(
            body,
            corner_radius=14,
            fg_color="#FFFFFF",
            border_width=1,
            border_color="#E2E8F0"
        )
        output_card.grid(row=4, column=0, sticky="ew", padx=28, pady=(12, 24))
        output_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            output_card,
            text="📁  FILE ĐÃ XUẤT",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1E293B"
        ).grid(row=0, column=0, sticky="w", padx=22, pady=(17, 12))

        self.output_buttons = []

        self.add_output_row(
            output_card, 1,
            "📄  File kết quả xét tốt nghiệp",
            "Chưa có file"
        )
        self.add_output_row(
            output_card, 2,
            "📊  File báo cáo thống kê",
            "Chưa có file"
        )
        self.add_output_row(
            output_card, 3,
            "📋  File trả kết quả tốt nghiệp",
            "Chưa có file"
        )

        self.add_output_row(
            output_card, 4,
            "✉  File trả email (DATA 30 cột)",
            "Chưa có file"
        )

        self.open_folder_button = ctk.CTkButton(
            output_card,
            text="📂  MỞ THƯ MỤC CHỨA FILE",
            height=40,
            corner_radius=9,
            fg_color="#E8F1FB",
            hover_color="#D8E8F8",
            text_color="#0B5CAD",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.open_output_folder,
            state="disabled"
        )
        self.open_folder_button.grid(
            row=5, column=0, sticky="ew",
            padx=22, pady=(12, 20)
        )

    def make_stat_box(self, parent, column, icon, title, value, text_color):
        box = ctk.CTkFrame(
            parent,
            corner_radius=12,
            fg_color="#F8FAFC"
        )
        box.grid(row=1, column=column, sticky="ew", padx=8)
        box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            box,
            text=icon,
            font=ctk.CTkFont(size=25)
        ).grid(row=0, column=0, pady=(14, 3))

        value_label = ctk.CTkLabel(
            box,
            text=value,
            font=ctk.CTkFont(size=25, weight="bold"),
            text_color=text_color
        )
        value_label.grid(row=1, column=0)

        ctk.CTkLabel(
            box,
            text=title,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#64748B"
        ).grid(row=2, column=0, pady=(2, 14))

        return value_label

    def add_output_row(self, parent, row, title, filename):
        frame = ctk.CTkFrame(
            parent,
            fg_color="#F8FAFC",
            corner_radius=9
        )
        frame.grid(row=row, column=0, sticky="ew", padx=22, pady=4)
        frame.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(
            frame,
            text=title,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#334155"
        )
        label.grid(row=0, column=0, sticky="w", padx=14, pady=12)

        button = ctk.CTkButton(
            frame,
            text="Mở file",
            width=90,
            height=32,
            corner_radius=7,
            state="disabled"
        )
        button.grid(row=0, column=1, padx=10)

        self.output_buttons.append((label, button))

    # --------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------
    def build_footer(self):
        footer = ctk.CTkFrame(
            self,
            height=38,
            corner_radius=0,
            fg_color="#EAF0F6"
        )
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)

        ctk.CTkLabel(
            footer,
            text="HUTECH • Viện Đại học Mở HUTECH",
            font=ctk.CTkFont(size=11),
            text_color="#64748B"
        ).pack(side="left", padx=20)

        ctk.CTkLabel(
            footer,
            text=f"Code by: {NAME_AUTHOR}",
            font=ctk.CTkFont(size=11),
            text_color="#64748B"
        ).pack(side="right", padx=20)

    # --------------------------------------------------------
    # FILE
    # --------------------------------------------------------
    def choose_file(self):
        from tkinter import filedialog

        path = filedialog.askopenfilename(
            title="Chọn file Excel xét tốt nghiệp",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro-Enabled Workbook", "*.xlsm"),
                ("All files", "*.*")
            ]
        )

        if path:
            self.selected_file = path
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, path)
            self.file_entry.configure(state="disabled")

            self.status_label.configure(
                text="Đã chọn file – Sẵn sàng xét tốt nghiệp.",
                text_color="#0B5CAD"
            )
            self.progress.set(0)

    # --------------------------------------------------------
    # XỬ LÝ
    # --------------------------------------------------------
    def start_process(self):
        if self.processing:
            return

        if not self.selected_file:
            from tkinter import messagebox
            messagebox.showwarning(
                "Chưa chọn file",
                "Anh vui lòng chọn file Excel trước khi bắt đầu."
            )
            return

        dialog = EmailInformationDialog(self, self.email_metadata)
        self.wait_window(dialog)
        if dialog.result is None:
            self.status_label.configure(text='Đã hủy – Chưa xác nhận thông tin trả email.', text_color='#B45309')
            return
        self.email_metadata = dialog.result

        # Cho người dùng chọn thư mục lưu kết quả ngay trước khi chạy.
        from tkinter import filedialog, messagebox

        output_dir = filedialog.askdirectory(
            title="Chọn thư mục lưu 4 file kết quả"
        )

        if not output_dir:
            self.status_label.configure(
                text="Đã hủy – Chưa chọn thư mục lưu kết quả.",
                text_color="#B45309"
            )
            return

        self.output_dir = Path(output_dir)

        self.processing = True
        self.run_button.configure(
            state="disabled",
            text="⏳  ĐANG XỬ LÝ..."
        )
        self.choose_button.configure(state="disabled")

        self.total_value.configure(text="...")
        self.eligible_value.configure(text="...")
        self.not_eligible_value.configure(text="...")
        self.rate_label.configure(
            text="Đang xử lý dữ liệu..."
        )

        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.status_label.configure(
            text="Đang đọc và xét dữ liệu Excel...",
            text_color="#0B5CAD"
        )

        thread = threading.Thread(
            target=self.run_process,
            daemon=True
        )
        thread.start()

    def run_process(self):
        try:
            info = process_excel(
                self.selected_file,
                output_dir=self.output_dir,
                email_metadata=self.email_metadata
            )
            self.after(0, lambda: self.process_success(info))
        except Exception as exc:
            error_text = f"{exc}\n\n{traceback.format_exc()}"
            self.after(0, lambda: self.process_error(error_text))

    def process_success(self, info):
        self.processing = False
        self.result_info = info

        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)

        self.run_button.configure(
            state="normal",
            text="🚀  BẮT ĐẦU XÉT TỐT NGHIỆP"
        )
        self.choose_button.configure(state="normal")

        total = info["total_count"]
        eligible = info["eligible_count"]
        not_eligible = info["not_eligible_count"]

        eligible_rate = eligible / total * 100 if total else 0
        not_eligible_rate = not_eligible / total * 100 if total else 0

        self.total_value.configure(text=f"{total:,}")
        self.eligible_value.configure(text=f"{eligible:,}")
        self.not_eligible_value.configure(text=f"{not_eligible:,}")

        self.rate_label.configure(
            text=(
                f"Tỷ lệ đủ điều kiện: {eligible_rate:.2f}%"
                f"    |    "
                f"Tỷ lệ không đủ: {not_eligible_rate:.2f}%"
            )
        )

        self.status_label.configure(
            text="✓ Đã xuất 4 file, gồm file trả email – Xem báo cáo để kiểm tra cảnh báo.",
            text_color="#15803D"
        )

        self.set_output_file(
            0,
            "📄  File kết quả xét tốt nghiệp",
            info["output_path"]
        )
        self.set_output_file(
            1,
            "📊  File báo cáo thống kê",
            info["report_path"]
        )

        # File thứ 3 được tạo bởi bản nâng cấp mới.
        # Tương thích cả trường hợp tên key khác nhau.
        tra_path = (
            info.get("return_result_path")
            or info.get("tra_ket_qua_path")
            or info.get("return_path")
            or info.get("tra_result_path")
        )

        if tra_path:
            self.set_output_file(
                2,
                "📋  File trả kết quả tốt nghiệp",
                tra_path
            )

        self.set_output_file(3, "✉  File trả email (DATA 30 cột)", info['email_result_path'])
        self.open_folder_button.configure(state="normal")

        from tkinter import messagebox
        messagebox.showinfo(
            "Xét tốt nghiệp hoàn tất",
            "ĐÃ XỬ LÝ XONG!\n\n"
            f"Tổng số trả kết quả: {total} học viên\n"
            f"Đủ điều kiện: {eligible} học viên ({eligible_rate:.2f}%)\n"
            f"Không đủ điều kiện: {not_eligible} học viên ({not_eligible_rate:.2f}%)\n"
            f"Cần kiểm tra: {info.get('review_count', 0)} học viên\n\n"
            f"Đã tạo 4 file tại:\n{Path(info['output_path']).parent}\n\n"
            f"File trả email:\n{info['email_result_path']}"
        )

    def process_error(self, error_text):
        self.processing = False

        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)

        self.run_button.configure(
            state="normal",
            text="🚀  BẮT ĐẦU XÉT TỐT NGHIỆP"
        )
        self.choose_button.configure(state="normal")

        self.status_label.configure(
            text="✕ Có lỗi xảy ra – Vui lòng kiểm tra thông báo.",
            text_color="#B91C1C"
        )

        from tkinter import messagebox
        messagebox.showerror(
            "Lỗi xét tốt nghiệp",
            error_text
        )

    def set_output_file(self, index, title, path):
        label, button = self.output_buttons[index]
        label.configure(text=title)

        button.configure(
            state="normal",
            text="Mở file",
            command=lambda p=path: open_file_or_folder(p)
        )

    def open_output_folder(self):
        if not self.result_info:
            return

        path = self.result_info["output_path"]
        open_file_or_folder(Path(path).parent)


def run_gui():
    if ctk is None:
        import tkinter.messagebox as mb
        root = tk.Tk()
        root.withdraw()
        mb.showerror(
            "Thiếu thư viện",
            "Chưa cài CustomTkinter.\n\n"
            "Vui lòng chạy:\n\n"
            "pip install customtkinter openpyxl"
        )
        root.destroy()
        return

    app = GraduationApp()
    app.mainloop()


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) > 1:
        import argparse
        import json
        parser = argparse.ArgumentParser(description="Xét tốt nghiệp linh hoạt")
        parser.add_argument("file")
        parser.add_argument("--output-dir")
        parser.add_argument("--sheet")
        parser.add_argument("--config")
        args = parser.parse_args()
        print(json.dumps(process_excel(args.file, args.output_dir, args.sheet, args.config), ensure_ascii=False, default=str, indent=2))
    else:
        run_gui()


if __name__ == "__main__":
    main()
