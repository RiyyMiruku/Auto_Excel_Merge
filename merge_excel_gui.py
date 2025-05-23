import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from itertools import islice, tee

def main():

    # 建立主程式視窗
    root = tk.Tk()
    root.title("Excel Sheet 合併工具")
    root.geometry("500x330")

    folder_path = tk.StringVar()
    sheet_index = tk.IntVar(value=1)
    output_filename = tk.StringVar()
    output_path = tk.StringVar()

    def browse_folder():
        path = filedialog.askdirectory()
        if path:
            folder_path.set(path)
            # 只取前6個符合條件的檔案
            files_gen = (f for f in os.listdir(path) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~'))
            files = list(islice(files_gen, 6))
            if files:
                # 顯示檔案清單給使用者確認
                file_list = "\n".join(files)
                messagebox.showinfo("資料夾內容", f"該資料夾內的 Excel 檔案：\n{file_list}\n...")
            else:
                messagebox.showwarning("提醒", "此資料夾內沒有 Excel 檔案")

    def browse_output_file():
        file_path = filedialog.askdirectory()

        if file_path:
            output_path.set(file_path)

    def merge_sheets():
  
        outname = output_filename.get()
        
        #確保輸出檔名不為空
        if not outname:
            messagebox.showerror("錯誤", "請輸入檔案名稱")
            return
        else:   
            outname = outname + ".xlsx"

        #確保合併的sheet_index為整數
        try:
            index = int(sheet_index.get()) - 1
            if index < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("錯誤", "請輸入正整數作為合併的 Sheet 編號")
            return
        
        #檢查輸入與輸出路徑
        path = folder_path.get()
        outpath = output_path.get()
        if not path or not outpath:
            messagebox.showerror("錯誤", "請選擇資料夾與輸出位置")
            return
        
        #確保輸入的資料夾內有 Excel 檔案
        files_gen = (f for f in os.listdir(path) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~'))


        # 統計個數並生成副本 iterator，使用一個統計總數量，另一個用於實際處理
        files_gen, files_iter = tee(files_gen)
        file_count = sum(1 for _ in files_gen)

        # 如果沒有檔案(0個符合條件)，顯示錯誤訊息
        if file_count == 0:
            messagebox.showerror("提示", "資料夾內沒有找到符合的 Excel 檔案")
            return

        #創建新的工作簿，清除第一頁空白頁
        merged_wb = Workbook()
        merged_wb.remove(merged_wb.active)

        #創建進度條
        progress["maximum"] = file_count
        progress["value"] = 0

        #處理每一個excel檔案
        for i, file in enumerate(files_iter):
            try:
                #取得Excel檔案路徑
                full_path = os.path.join(path, file)

                #使用openpyxl讀取.xlsx檔案，使用xlrd讀取.xls檔案來分別處裡
                df = pd.read_excel(full_path, sheet_name=index, engine='openpyxl' if file.endswith('.xlsx') else 'xlrd')

                #取得檔名寫入新的工作表
                sheet_name = os.path.splitext(file)[0]
                ws = merged_wb.create_sheet(title=f"{sheet_name}")

                #將DataFrame寫入新工作表
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
            except Exception as e:
                messagebox.showwarning("警告", f"無法處理檔案：{file}\n錯誤：{e}")
                return
            
            # 更新進度條
            progress["value"] = i + 1
            root.update_idletasks()

        # 儲存合併後的檔案
        try:
            save_path = os.path.join(outpath, outname)
            merged_wb.save(save_path)
            messagebox.showinfo("完成", f"合併完成，儲存到：\n{outpath}\n檔名：{outname}")
        except Exception as e:
            messagebox.showwarning("警告", f"無法處理檔案：{file}\n錯誤：{e}")

        

    tk.Label(root, text="選擇資料夾：").pack(anchor='w', padx=10, pady=(10,0))
    tk.Entry(root, textvariable=folder_path, width=60).pack(padx=10)
    tk.Button(root, text="瀏覽", command=browse_folder).pack(padx=10, pady=5)

    tk.Label(root, text="合併第幾個 Sheet（預設為 1）：").pack(anchor='w', padx=10)
    tk.Entry(root, textvariable=sheet_index, width=10).pack(padx=10)

    tk.Label(root, text="儲存路徑：").pack(anchor='w', padx=10)
    tk.Entry(root, textvariable=output_path, width=60).pack(padx=10)
    tk.Button(root, text="選擇輸出位置", command=browse_output_file).pack(padx=10, pady=5)

    tk.Label(root, text="輸出檔案名稱：").pack(anchor='w', padx=10)
    tk.Entry(root, textvariable=output_filename, width=60).pack(padx=10)   

   
    progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress.pack(pady=10)

    tk.Button(root, text="開始合併", command=merge_sheets).pack(pady=5)

    root.mainloop()



if __name__ == "__main__":
    main()
