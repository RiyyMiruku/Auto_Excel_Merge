import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def main():

    # 建立主程式視窗
    root = tk.Tk()
    root.title("Excel Sheet 合併工具")
    root.geometry("500x250")

    # 儲存選擇的資料夾路徑
    folder_path = tk.StringVar()
    sheet_index = tk.IntVar(value=1)
    output_filename = tk.StringVar()

    # 資料夾選擇功能
    def browse_folder():
        path = filedialog.askdirectory()
        if path:
            folder_path.set(path)
            # 自動設定預設輸出檔名為資料夾內第一個檔名
            files = [f for f in os.listdir(path) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
            if files:
                output_filename.set(os.path.splitext(files[0])[0] + "_merged.xlsx")

    # 合併處理主功能
    def merge_sheets():
        path = folder_path.get()
        index = sheet_index.get() - 1
        outname = output_filename.get()

        if not path or not outname:
            messagebox.showerror("錯誤", "請選擇資料夾並輸入輸出檔名")
            return

        files = [f for f in os.listdir(path) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        if not files:
            messagebox.showerror("錯誤", "資料夾內沒有找到符合的 Excel 檔案")
            return

        merged_wb = Workbook()
        merged_wb.remove(merged_wb.active)  # 刪除預設空白工作表

        progress["maximum"] = len(files)
        progress["value"] = 0

        for i, file in enumerate(files):
            try:
                full_path = os.path.join(path, file)
                df = pd.read_excel(full_path, sheet_name=index, engine='openpyxl' if file.endswith('.xlsx') else 'xlrd')
                ws = merged_wb.create_sheet(title=f"Sheet{i+1}")
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
            except Exception as e:
                messagebox.showwarning("警告", f"無法處理檔案：{file}\n錯誤：{e}")
            progress["value"] = i + 1
            root.update_idletasks()

        save_path = os.path.join(path, outname)
        merged_wb.save(save_path)
        messagebox.showinfo("完成", f"合併完成，儲存為：{save_path}")

    # GUI 元件配置
    tk.Label(root, text="選擇資料夾：").pack(anchor='w', padx=10, pady=(10,0))
    tk.Entry(root, textvariable=folder_path, width=60).pack(padx=10)
    tk.Button(root, text="瀏覽", command=browse_folder).pack(padx=10, pady=5)

    tk.Label(root, text="合併第幾個 Sheet（預設為 1）：").pack(anchor='w', padx=10)
    tk.Entry(root, textvariable=sheet_index, width=10).pack(padx=10)

    tk.Label(root, text="輸出檔名：").pack(anchor='w', padx=10)
    tk.Entry(root, textvariable=output_filename, width=60).pack(padx=10)

    progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress.pack(pady=10)

    tk.Button(root, text="開始合併", command=merge_sheets).pack(pady=5)

    # 啟動 GUI 主迴圈
    root.mainloop()



if __name__ == "__main__":
    main()
