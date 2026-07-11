import os
import re
import platform
from PIL import Image

# 取得目前作業系統的預設字型
def get_system_font():
    current_os = platform.system()
    if current_os == "Windows":
        return "Microsoft JhengHei"  # 微軟正黑體
    elif current_os == "Darwin":
        return "PingFang TC"        # 蘋果繁體蘋方
    elif current_os == "Linux":
        return "Noto Sans CJK TC"    # 思源黑體繁體
    else:
        return "Arial"

# 全域字型設定
SYSTEM_FONT = get_system_font()
# macOS 字型放大 4pt，Windows 與 Linux 放大 2pt
FONT_OFFSET = 4 if platform.system() == "Darwin" else 2

# 解析拖放的多個檔案路徑 (支援帶空格且有大括號的 Windows 路徑)
def parse_dropped_files(data_str):
    pattern = r'\{([^}]+)\}|(\S+)'
    matches = re.findall(pattern, data_str)
    files = []
    for m in matches:
        path = m[0] if m[0] else m[1]
        if path:
            files.append(os.path.normpath(path))
    return files

# 自動處理檔名重複 (若檔案已存在，命名為 _1, _2 等)
def unique_filename(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while True:
        new_p = f"{base}_{i}{ext}"
        if not os.path.exists(new_p):
            return new_p
        i += 1

# 解析使用者輸入的頁碼範圍字串 (如 "1-5, 7, 9-11")，回傳對應的 0-based 頁碼列表
def parse_range_string(range_str, max_pages):
    pages = []
    if not range_str.strip():
        return list(range(max_pages))
    
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            try:
                start, end = part.split('-')
                s = int(start.strip())
                e = int(end.strip())
                s = max(1, min(s, max_pages))
                e = max(1, min(e, max_pages))
                if s <= e:
                    pages.extend(range(s - 1, e))
                else:
                    pages.extend(range(e - 1, s))
            except:
                pass
        else:
            try:
                p = int(part)
                p = max(1, min(p, max_pages))
                pages.append(p - 1)
            except:
                pass
    return sorted(list(set(pages)))

# 確保應用程式 ICON 檔案存在。若缺少 ico 但有 png，則自動進行轉換
def ensure_app_icon():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    png_path = os.path.join(base_dir, "app_icon.png")
    ico_path = os.path.join(base_dir, "app_icon.ico")
    
    if not os.path.exists(ico_path) and os.path.exists(png_path):
        try:
            # 轉換為 .ico 檔，包入多種尺寸以供 Windows 系統顯示
            img = Image.open(png_path)
            img.save(ico_path, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
            return ico_path
        except Exception as e:
            print(f"ICON 轉換失敗: {e}")
            return None
    elif os.path.exists(ico_path):
        return ico_path
    return None

def finalize_and_save_pdf(doc, save_path, pdf_settings):
    """
    對 fitz.Document 物件統一進行浮水印套用、Metadata 設定，並根據加密參數儲存至指定路徑。
    """
    import fitz
    
    # 1. 智慧文字浮水印
    if pdf_settings.get("watermark") and pdf_settings.get("wm_text"):
        from utils.watermark import apply_watermark_to_page
        wm_text = pdf_settings["wm_text"]
        wm_opacity = pdf_settings.get("wm_opacity", 0.3)
        wm_angle = pdf_settings.get("wm_angle", 45)
        wm_tile = pdf_settings.get("wm_tile", False)
        wm_color = pdf_settings.get("wm_color", "灰色")
        wm_size = pdf_settings.get("wm_size", 48)
        
        for page in doc:
            apply_watermark_to_page(
                page,
                text=wm_text,
                opacity=wm_opacity,
                angle=wm_angle,
                tile=wm_tile,
                color_name=wm_color,
                size=wm_size
            )
            
    # 2. Metadata 設定
    meta = {
        "title": pdf_settings.get("meta_title", ""),
        "author": pdf_settings.get("meta_author", ""),
        "subject": pdf_settings.get("meta_subject", ""),
        "keywords": pdf_settings.get("meta_keywords", ""),
        "creator": "PDF & Image Toolkit",
        "producer": "PyMuPDF"
    }
    doc.set_metadata(meta)
    
    # 3. 加密防護與儲存
    enc = pdf_settings.get("encrypt", False)
    opw = pdf_settings.get("password", "")
    
    if enc and opw:
        perm = int(
            fitz.PDF_PERM_ACCESSIBILITY |
            fitz.PDF_PERM_PRINT |
            fitz.PDF_PERM_COPY |
            fitz.PDF_PERM_ANNOTATE
        )
        doc.save(save_path, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=opw, owner_pw=opw, permissions=perm, garbage=4, deflate=True)
    else:
        doc.save(save_path, garbage=4, deflate=True)
