# Auto_Excel_Merge
A little tool that can combine all excels into sheets in a single excel

# Excel Sheet 合併工具

這是一個使用 Python 開發的桌面應用程式，可在 Windows 環境下執行，協助使用者將一個資料夾內的所有 Excel 檔案（支援 `.xlsx` 和 `.xls`）的指定工作表合併到一個新的 Excel 檔案中。

---

## 📦 功能說明

- ✅ 選擇資料夾
- ✅ 自訂要合併的工作表索引（預設為第 1 張 Sheet）
- ✅ 自訂輸出檔名（預設為資料夾內第一個檔案的名稱 + `_merged.xlsx`）
- ✅ 自動將每個 Excel 的資料儲存在 `Sheet1`、`Sheet2`、... 中
- ✅ 圖形介面 + 合併進度條
- ✅ 支援 `.xlsx` 與 `.xls`

---

## 🚀 執行方式

### 🔧 安裝相依套件（第一次使用前需執行一次）：

```bash
pip install pandas openpyxl xlrd
```

### ▶️ 執行程式：

```bash
python merge_excel_gui.py
```

---

## 💡 打包為 .exe 可執行檔（可供非技術使用者）

1. 安裝打包工具：

```bash
pip install pyinstaller
```

2. 打包指令：

```bash
pyinstaller --noconsole --onefile merge_excel_gui.py
```

3. 完成後，`dist/merge_excel_gui.exe` 即為可執行檔案，可直接雙擊開啟使用。

---

## ⚠️ 注意事項

- 所有來源 Excel 檔案應為格式正確、未受密碼保護。
- 輸出檔會儲存在原始資料夾中。
- Sheet 索引為從 1 開始計算（非 0）。

---

## 🧑‍💻 作者與貢獻
由 ChatGPT 與使用者共同設計
