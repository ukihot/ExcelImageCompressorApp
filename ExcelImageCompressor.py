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


# フォルダ選択ダイアログを表示し、選択したフォルダのパスを返す関数
def select_folder():
    return filedialog.askdirectory()


# 指定されたフォルダ内の全てのExcelファイルを検索する関数
def find_excel_files(folder):
    return [
        os.path.join(root, file)
        for root, _, files in os.walk(folder)
        for file in files
        if file.endswith((".xlsx", ".xlsm")) and not file.startswith("~$")
    ]


# 画像を圧縮する関数
def compress_image(image):
    with Image.open(BytesIO(image._data())) as img:
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format=img.format if img.format else "PNG", dpi=(96, 96))
        return OpenpyxlImage(img_byte_arr)


# ファイルサイズをKB単位で取得する関数
def get_file_size_in_kb(file_path):
    return os.path.getsize(file_path) / 1024


# Excelファイル内の画像を圧縮する関数
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


# 圧縮結果を報告するレポートを作成する関数
def create_report(report_data, folder):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = os.path.join(folder, f"report_{timestamp}.xlsx")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"

    # レポートのヘッダーを設定
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

    # データ行をフォーマットして追加
    def format_row(row):
        return [
            f"{value:.1f}" if isinstance(value, (int, float)) else value
            for value in row
        ]

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
            "Success" if not data["error_message"] else "Failed",
            data["error_message"] or "",
        ]
        sheet.append(format_row(row))

    # 数値列のフォーマットを設定
    for col in range(1, 10):
        for cell in sheet.iter_cols(min_col=col, max_col=col, min_row=2):
            for c in cell:
                if isinstance(c.value, (int, float)):
                    c.number_format = "0.0"

    workbook.save(report_file)
    return report_file


# フォルダ内の全てのExcelファイルの画像を圧縮し、レポートを作成する関数
def compress_images_in_folder(folder, progress_callback):
    files = find_excel_files(folder)
    report_data = []

    for i, file in enumerate(files, start=1):
        data = process_file(file)
        report_data.append(data)
        progress_callback(i, len(files))

    report_file = create_report(report_data, folder)
    return len(files), report_file


# アプリケーションのGUIを定義するクラス
class ExcelImageCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Image Compressor")
        self.selected_folder = tk.StringVar()

        # GUIの構成要素を作成
        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.folder_label = tk.Label(frame, text="Select Folder:")
        self.folder_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        self.folder_entry = tk.Entry(frame, textvariable=self.selected_folder, width=40)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)

        self.select_button = tk.Button(frame, text="Browse", command=self.set_folder)
        self.select_button.grid(row=0, column=2, padx=5, pady=5)

        self.start_button = tk.Button(
            frame, text="Start Compression", command=self.compress_images
        )
        self.start_button.grid(row=1, column=0, columnspan=3, pady=10)

        self.progress = Progressbar(
            frame, orient="horizontal", length=300, mode="determinate"
        )
        self.progress.grid(row=2, column=0, columnspan=2, pady=10)

        self.progress_label = tk.Label(frame, text="0.00%")
        self.progress_label.grid(row=2, column=2, padx=5, pady=10, sticky=tk.W)

    # フォルダ選択ボタンの処理
    def set_folder(self):
        self.selected_folder.set(select_folder())

    # 圧縮処理を開始するメソッド
    def compress_images(self):
        folder = self.selected_folder.get()
        if not folder:
            messagebox.showerror("Error", "No folder selected.")
            return

        try:
            total_files, report_file = compress_images_in_folder(
                folder, self.update_progress
            )
            messagebox.showinfo(
                "Completed",
                f"Compression completed for {total_files} files. Report generated at:\n{report_file}",
            )
            # レポートファイルをExcelで開き、アプリケーションを終了
            subprocess.run(["start", "excel", report_file], shell=True)
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    # プログレスバーを更新するメソッド
    def update_progress(self, current, total):
        percentage = (current / total) * 100
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.progress_label.config(text=f"{percentage:.2f}%")
        self.root.update_idletasks()


# メイン関数
def main():
    root = tk.Tk()
    app = ExcelImageCompressorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
