# PDF Tools

視覺化 PDF 頁面編輯器，所有操作皆在本機完成，不上傳任何資料。

## 功能

- 開啟 PDF，以縮圖預覽所有頁面（大檔案先顯示版面、縮圖在背景載入）
- 圖片轉 PDF：可直接開啟 / 附加 PNG、JPG、BMP、GIF、TIFF（依 EXIF 自動轉正，多頁 TIFF 拆成多頁）
- 支援開啟有密碼的 PDF
- 直接從檔案總管把 PDF 或圖片拖進視窗開啟 / 附加，開檔對話框也可多選
- 拖拉頁面自由排序，雙擊縮圖放大預覽
- 附加更多 PDF（合併入當前頁面）
- 單頁或批次旋轉 / 刪除，可復原 / 重做
- 插入空白頁（尺寸比照相鄰頁面）
- 匯出：擷取選取頁為 PDF、分割成多個 PDF（依頁碼範圍 / 每 N 頁 / 每頁）、轉成圖片（PNG / JPG，96 – 300 DPI）
- 輸出保留書籤，並自動去除重複資源、壓縮檔案
- 輸出選項（套用於所有輸出的檔案）：
  - 加密（AES-256，設定開啟密碼、禁止複製文字）
  - 文字浮水印（大小、透明度、斜放可調，支援中文）
  - 頁碼（位置、格式、起始數字可調，可略過封面）
  - 壓縮圖片（標準 150 DPI / 強力 96 DPI）
- 介面縮放（70% – 400%），縮放比例與上次使用的資料夾會自動記住

### 快捷鍵

| 按鍵 | 功能 |
|------|------|
| Ctrl+O / Ctrl+S | 開啟 / 儲存 |
| Ctrl+Z / Ctrl+Y | 復原 / 重做 |
| Ctrl+A / Esc | 全選 / 取消選取 |
| Shift+點擊 | 選取範圍 |
| R / Delete | 旋轉 / 刪除選取頁 |
| 雙擊縮圖 | 預覽；← → 換頁，Esc 關閉 |

## 使用方式

### 一般使用者（Windows）

直接下載 `PDF_Tools.exe` 執行，無需安裝任何東西。  
需要 Windows 10 / 11（內建 Edge WebView2）。

---

## 開發

### 環境需求（Linux / WSL2）

```bash
# 系統套件
sudo apt-get install -y python3-gi python3-gi-cairo \
    gir1.2-gtk-3.0 gir1.2-webkit2-4.1 libwebkit2gtk-4.1-0 \
    fonts-wqy-zenhei xfonts-wqy

# Python 套件（使用系統 Python 3.12）
/usr/bin/python3.12 -m pip install --user --break-system-packages -r requirements.txt
```

### 執行（開發測試）

```bash
/usr/bin/python3.12 pdf_tools.py
```

### 打包成 Windows .exe（從 WSL2 執行）

需要先在 Windows 安裝 Python（[python.org](https://www.python.org/)，安裝時勾選 Add to PATH）。

```bash
# 安裝 Windows 端套件（只需執行一次）
python.exe -m pip install -r requirements-dev.txt

# 打包
python.exe -m PyInstaller PDF_Tools.spec
```

Linux 版本則使用 `PDF_Tools_Linux.spec`。

產出位於 `dist/PDF_Tools.exe`。

### 測試

```bash
python -m pip install -r requirements-dev.txt
python -m pytest tests
```

### 專案結構

```
pdf_tools.py   # 主程式（Python 後端 + PyWebView 視窗）
tests/         # 後端測試（pytest）
web/
  index.html   # 介面結構
  style.css    # 樣式
  app.js       # 前端邏輯（拖拉、縮放、API 呼叫）
```

### 主要依賴

| 套件 | 用途 |
|------|------|
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | PDF 讀寫、頁面渲染、加密 |
| [pywebview](https://pywebview.flowrl.com/) | 以 HTML/CSS/JS 建立原生視窗 |
| [PyInstaller](https://pyinstaller.org/) | 打包成單一執行檔 |
