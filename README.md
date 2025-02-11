# Excel to PDF Conversion Tool

## 簡介
這是一個用 Python 編寫的工具，用來將 `.xlsx` 文件轉換為 `.pdf` 文件。程式會自動掃描指定目錄中的 Excel 文件，讀取其內容，並生成對應的 PDF 文件。適合用於批量處理 Excel 文件並快速生成 PDF 報表。

---

## 功能特色
- 自動檢測目錄下的 `.xlsx` 文件。
- 將每個 Excel 文件轉換為結構化的 PDF 文檔。
- 支援多平台路徑處理，生成的 PDF 保存到指定文件夾。

---

## 使用的技術與庫
- **`glob`**: 搜索目錄中的 `.xlsx` 文件。
- **`pathlib.Path`**: 處理文件和路徑，保證跨平台兼容性。
- **`fpdf.FPDF`**: 用於生成和格式化 PDF 文件。

---

## 安裝方式

1. 確保已安裝 **Python 3.x**。
2. 安裝所需的 Python 模組：
   ```bash
   pip install fpdf
   ```

---

## 使用方式

1. **準備文件**：
   - 將所有 `.xlsx` 文件放入 `input/` 資料夾中。

2. **運行腳本**：
   - 執行以下指令：
     ```bash
     python main.py
     ```

3. **查看輸出**：
   - 轉換後的 `.pdf` 文件將存放在 `output/` 資料夾中。

---

## 專案結構
```

project/
│
├── main.py               # 主程式，用於執行轉換
├── input/                # 放置 .xlsx 文件的資料夾
│     ├── PDFs/               # 生成 .pdf 文件的資料夾
└── README.md             # 專案說明文件
```

---

## 未來改進
- 支援更多文件格式，例如 `.csv` 或 `.xls`。
- 增加自定義 PDF 格式的功能（例如顏色、字體、表格佈局）。
- 提升錯誤處理能力，避免因文件損壞導致程式崩潰。

---