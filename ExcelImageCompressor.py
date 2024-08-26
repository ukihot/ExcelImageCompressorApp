import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
from io import BytesIO
from datetime import datetime
import subprocess


def select_folder():
    return filedialog.askdirectory()


def find_excel_files(folder):
    return [
        os.path.join(root, file)
        for root, _, files in os.walk(folder)
        for file in files
        if file.endswith((".xlsx", ".xlsm")) and not file.startswith("~$")
    ]


def compress_image(image):
    with Image.open(BytesIO(image._data())) as img:
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format=img.format if img.format else "PNG", dpi=(96, 96))
        return OpenpyxlImage(img_byte_arr)


def get_file_size_in_kb(file_path):
    return os.path.getsize(file_path) / 1024


def process_file(file):
    start_time = time.time()
    original_size = get_file_size_in_kb(file)
    error_message = None

    try:
        workbook = load_workbook(file)
        total_images = sum(len(ws._images) for ws in workbook.worksheets)
        updated_images = 0

        for ws in workbook.worksheets:
            images = ws._images
            compressed_images = [compress_image(img) for img in images]
            updated_images += len(compressed_images)
            ws._images[:] = compressed_images

        if updated_images > 0:
            workbook.save(file)

    except Exception as e:
        error_message = str(e)

    elapsed_time = (time.time() - start_time) * 1000
    new_size = get_file_size_in_kb(file) if not error_message else original_size
    return {
        "original_size": original_size,
        "new_size": new_size,
        "updated_images": updated_images,
        "elapsed_time": elapsed_time,
        "error_message": error_message,
        "file": file,
    }


def create_report(report_data, folder):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = os.path.join(folder, f"report_{timestamp}.xlsx")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "レポート"

    headers = [
        "Date",
        "Orig Size (KB)",
        "New Size (KB)",
        "Img Count",
        "Path",
        "File",
        "Time (ms)",
        "Status",
        "Error Message",
    ]
    sheet.append(headers)

    def format_row(row):
        return [
            f"{value:.1f}" if isinstance(value, (int, float)) else value
            for value in row
        ]

    total_original_size = 0
    total_new_size = 0
    total_elapsed_time = 0

    for data in report_data:
        path, file = os.path.relpath(data["file"], folder), os.path.basename(
            data["file"]
        )
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data["original_size"],
            data["new_size"],
            data["updated_images"],
            path if path != file else ".",
            file,
            data["elapsed_time"],
            "成功" if not data["error_message"] else "失敗",
            data["error_message"] or "",
        ]
        sheet.append(format_row(row))

        total_original_size += data["original_size"]
        total_new_size += data["new_size"]
        total_elapsed_time += data["elapsed_time"]

    # 数値列のフォーマットを設定
    for col in range(1, 10):
        for cell in sheet.iter_cols(min_col=col, max_col=col, min_row=2):
            for c in cell:
                if isinstance(c.value, (int, float)):
                    c.number_format = "0.0"

    # 合計行を追加
    compression_ratio = (
        (1 - total_new_size / total_original_size) * 100
        if total_original_size > 0
        else 0
    )
    summary_row = [
        "",
        f"{total_original_size / 1024:.2f} MB",
        f"{total_new_size / 1024:.2f} MB ({compression_ratio:.2f}%)",
        "",
        "",
        f"{total_elapsed_time / 1000:.2f} s",
        "",
    ]
    sheet.append(summary_row)

    workbook.save(report_file)
    return report_file


def compress_images_in_folder(folder, progress_callback):
    files = find_excel_files(folder)
    report_data = []

    for i, file in enumerate(files, start=1):
        data = process_file(file)
        report_data.append(data)
        progress_callback(i, len(files))

    report_file = create_report(report_data, folder)
    return len(files), report_file


class ExcelImageCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Image Compressor")
        self.selected_folder = tk.StringVar()

        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.folder_label = tk.Label(frame, text="フォルダ選択:")
        self.folder_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        self.folder_entry = tk.Entry(frame, textvariable=self.selected_folder, width=40)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)

        self.select_button = tk.Button(frame, text="参照", command=self.set_folder)
        self.select_button.grid(row=0, column=2, padx=5, pady=5)

        self.start_button = tk.Button(
            frame, text="圧縮開始", command=self.compress_images
        )
        self.start_button.grid(row=1, column=0, columnspan=3, pady=10)

        self.progress = Progressbar(
            frame, orient="horizontal", length=300, mode="determinate"
        )
        self.progress.grid(row=2, column=0, columnspan=2, pady=10)

        self.progress_label = tk.Label(frame, text="0.00%")
        self.progress_label.grid(row=2, column=2, padx=5, pady=10, sticky=tk.W)

    def set_folder(self):
        self.selected_folder.set(select_folder())

    def compress_images(self):
        folder = self.selected_folder.get()
        if not folder:
            messagebox.showerror("エラー", "フォルダが選択されていません。")
            return

        try:
            total_files, report_file = compress_images_in_folder(
                folder, self.update_progress
            )
            messagebox.showinfo(
                "完了",
                f"{total_files} 件のファイルの圧縮が完了しました。レポートが作成されました:\n{report_file}",
            )
            subprocess.run(["start", "excel", report_file], shell=True)
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました: {e}")

    def update_progress(self, current, total):
        percentage = (current / total) * 100
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.progress_label.config(text=f"{percentage:.2f}%")
        self.root.update_idletasks()


def main():
    root = tk.Tk()
    app = ExcelImageCompressorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
