# PDF 圖片轉換小工具 (pdf-image-converter)

[![GitHub License](https://img.shields.io/github/license/kaoshou/pdf-image-converter)](https://github.com/kaoshou/pdf-image-converter/blob/main/LICENSE)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![Platform Support](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey.svg)](https://github.com/kaoshou/pdf-image-converter)

一個基於 Python 和 CustomTkinter 打造的精美、現代化桌上型 **PDF 與圖片雙向轉換工具箱**。提供極致流暢的 UI/UX 互動體驗，無任何廣告、完全離線運行，100% 保證您的文件隱私安全。

---

## 🎨 介面視覺與 UX 亮點

* **Dashboard 側邊欄導覽**：採用現代控制台佈局，頁面切換流暢，功能結構清晰。
* **深/淺色主題開關**：底部集成一鍵切換開關，無縫適配「極客深色」與「商務淺色」模式。
* **實時拖曳排序 (Drag & Drop)**：清單項目支援滑鼠「即時隨動」拖曳重排，視覺引導極佳。
* **空白拖曳提示**：檔案清單空白時，提供顯眼的引導區域，支援將檔案直接拖放到引導 Label 快速載入。
* **右鍵捷徑操作**：所有清單項目皆支援右鍵選單，可「在檔案總管中顯示並定位」或「移除此項目」。

---

## 🚀 六大核心功能介紹

### 1. 📂 圖片/PDF ➔ PDF 合併
* 支援混合圖片與 PDF 文件進行批次合併。
* **PDF 自動展開**：可一鍵將多頁 PDF 拆解展開為獨立頁面，隨心混編。
* 支援設定頁面尺寸（原始大小、A4、A3 等）、方向（橫式/直式）與縮放模式（自動填滿/保持比例）。
* 整合圖片壓縮、黑白化 (B&W)、自動旋轉糾正以及 PDF 加密防護功能。

### 2. 🖼️ PDF ➔ 圖片
* 將 PDF 文件批次導出為常見的圖片格式 (PNG, JPEG, BMP, TIFF)。
* 可自訂導出解析度 (DPI)，滿足高解析度輸出需求。

### 3. ✂️ PDF 拆分與擷取
* 可自由指定頁碼範圍進行拆分（例如 `1-3, 5, 8-12`）。
* 提供拆分模式選擇：**「擷取指定頁面並合併」** 或 **「每頁單獨拆分成一個 PDF」**。

### 4. ⚡ PDF 壓縮
* 專門針對掃描版與大體積 PDF 進行極速瘦身。
* 支援多種壓縮品質與黑白化功能。
* **多核心加速**：使用線程池並行渲染技術，大幅提升灰階 PDF 的壓縮與處理速度。

### 5. 🔒 PDF 加密防護
* 為 PDF 一鍵加上高強度的開啟密碼與限制權限。
* 對於需要輸入密碼解鎖的檔案，會自動彈出高質感密碼驗證框。

### 6. ⚙️ 圖片壓縮與縮放
* 支援大批圖片的格式轉換 (JPG, PNG, WEBP) 與尺寸/品質壓縮。
* 支援圖片方向批次旋轉、移除 EXIF 隱私資訊與 PNG 色彩量化技術。
* **即時對比預覽**：右鍵點擊項目可彈出雙欄對比視窗，在背景虛擬運行壓縮並即時呈現原圖/壓縮圖的畫質與體積縮減對比 (如 `預估體積減少 90%`)。

---

## 💻 系統要求與安裝說明

本程式完全開源，且不需要安裝額外的資料庫或繁重的框架。

### 1. 克隆專案
```bash
git clone https://github.com/kaoshou/pdf-image-converter.git
cd pdf-image-converter
```

### 2. 安裝依賴套件
推薦在虛擬環境中運行本專案：
```bash
python -m venv venv
# 啟用虛擬環境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 安裝依賴套件
pip install -r requirements.txt
```

> **Note**: 本程式主要依賴 `customtkinter` (UI 庫)、`pymupdf` (PDF 處理庫)、`pillow` (圖片處理庫) 以及 `tkinterdnd2` (拖放庫)。

### 3. 啟動應用程式
```bash
python pdf_image_toolkit.py
```

---

## 📦 如何打包成單一免安裝執行檔 (.exe)

如果您想將其打包成一個可以直接在 Windows 上運行的單一 `.exe` 綠色軟體，請遵循以下步驟：

1. 安裝 PyInstaller：
   ```bash
   pip install pyinstaller
   ```
2. 執行打包指令（指令中已包含圖示及拖放 DLL 依賴項的複製）：
   ```bash
   pyinstaller --noconfirm --onedir --windowed --name "PDF圖片轉換小工具" --icon "assets/app_icon.ico" --add-data "assets;assets" --add-data "components;components" --add-data "utils;utils" pdf_image_toolkit.py
   ```
3. 打包完成後，可在 `dist/PDF圖片轉換小工具` 資料夾內找到可執行檔。

---

## 📜 第三方套件與開源聲明

本專案基於以下優秀開源專案與授權協議實作：

* [PyMuPDF (fitz)](https://github.com/pymupdf/PyMuPDF) - GNU AGPL v3.0 授權
* [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - MIT 授權
* [TkinterDnD2](https://github.com/pmgagne/tkinterdnd2) - MIT 授權

---

## ✉️ 聯絡開發者
* **開發者**：鄭郁翰 (Cheng, Yu-Han)
* **Email**：kaoshou@gmail.com
* **GitHub**：[kaoshou](https://github.com/kaoshou)
