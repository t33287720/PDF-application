# PDF Tools

視覺化 PDF 頁面編輯器，所有操作皆在本機完成，不上傳任何資料。

## 功能

- 開啟 PDF，以縮圖預覽所有頁面
- 拖拉頁面自由排序
- 附加更多 PDF（合併入當前頁面）
- 單頁或批次旋轉 / 刪除
- 擷取選取頁面另存新檔
- 輸出時可選擇加密（設定開啟密碼）
- 介面縮放（70% – 400%）

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
/usr/bin/python3.12 -m pip install --user --break-system-packages \
    pikepdf pymupdf pywebview
```

### 執行（開發測試）

```bash
/usr/bin/python3.12 pdf_tools.py
```

### 打包成 Windows .exe（從 WSL2 執行）

需要先在 Windows 安裝 Python（[python.org](https://www.python.org/)，安裝時勾選 Add to PATH）。

```bash
# 安裝 Windows 端套件（只需執行一次）
python.exe -m pip install pikepdf pymupdf pywebview pyinstaller

# 打包
python.exe -m PyInstaller --onefile --windowed --add-data "web;web" --name "PDF_Tools" pdf_tools.py
```

產出位於 `dist/PDF_Tools.exe`。

### 專案結構

```
pdf_tools.py   # 主程式（Python 後端 + PyWebView 視窗）
web/
  index.html   # 介面結構
  style.css    # 樣式
  app.js       # 前端邏輯（拖拉、縮放、API 呼叫）
```

### 主要依賴

| 套件 | 用途 |
|------|------|
| [pikepdf](https://github.com/pikepdf/pikepdf) | PDF 讀寫、加密 |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | 頁面渲染成縮圖 |
| [pywebview](https://pywebview.flowrl.com/) | 以 HTML/CSS/JS 建立原生視窗 |
| [PyInstaller](https://pyinstaller.org/) | 打包成單一執行檔 |
