from PIL import Image, ImageDraw
import customtkinter as ctk

def create_icon_image(icon_name, is_dark):
    """
    在 64x64 的透明畫布上動態繪製扁平化、現代線條風格的 UI 圖示。
    
    參數:
      icon_name (str): 圖示鍵值名稱，如 'logo', 'run', 'cancel' 等。
      is_dark (bool): 是否為深色模式，決定輸出的圖示顏色。
      
    回傳:
      PIL.Image: 具備透明背景的 RGBA 圖示影像。
    """
    # 建立 64x64 像素、完全透明背景的 RGBA 畫布
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 決定配色風格，以契合 CustomTkinter 的外觀模式
    if is_dark:
        # 深色模式下：主要線條為淺灰色，重點裝飾/箭頭為亮藍色
        color_main = (229, 231, 235, 255)  # 淺灰 #E5E7EB
        color_accent = (96, 165, 250, 255) # 亮藍 #60A5FA
    else:
        # 淺色模式下：主要線條為深灰色，重點裝飾/箭頭為深藍色
        color_main = (55, 65, 81, 255)     # 深灰 #374151
        color_accent = (37, 99, 235, 255)  # 深藍 #2563EB
        
    def draw_doc_shape(draw, x0, y0, x1, y1, color, fold=12):
        """
        輔助函式：繪製一個帶有右上摺角的文件檔案外框。
        """
        # 繪製文件外邊框 (右上角缺角)
        draw.line([
            (x0, y0),
            (x1 - fold, y0),
            (x1, y0 + fold),
            (x1, y1),
            (x0, y1),
            (x0, y0)
        ], fill=color, width=4, joint="round")
        # 繪製右上摺角內部的折痕線條
        draw.line([
            (x1 - fold, y0),
            (x1 - fold, y0 + fold),
            (x1, y0 + fold)
        ], fill=color, width=4, joint="round")

    def draw_image_shape(draw, x0, y0, x1, y1, color):
        """
        輔助函式：繪製一個相框，內部帶有山脈與太陽的圖片象徵。
        """
        # 繪製圓角相框外框
        draw.rounded_rectangle([x0, y0, x1, y1], radius=6, outline=color, width=4)
        # 繪製相框內的大山 (用三角形連線表示)
        draw.line([
            (x0 + 6, y1 - 4),
            (x0 + (x1 - x0) // 2 + 2, y0 + 12),
            (x1 - 6, y1 - 4),
            (x0 + 6, y1 - 4)
        ], fill=color, width=4, joint="round")
        # 繪製太陽 (右上角圓形)
        draw.ellipse([x1 - 18, y0 + 8, x1 - 10, y0 + 16], outline=color, width=3)

    # ================== 🎨 圖示樣式繪製分流 ==================

    if icon_name == "logo":
        # 螺絲起子與扳手交叉圖示 (代表 ToolKit 工具箱)
        # 1. 螺絲起子 (桿子為主色，手柄為主題裝飾色)
        draw.line([(16, 16), (48, 48)], fill=color_main, width=5)
        draw.line([(40, 40), (50, 50)], fill=color_accent, width=10, joint="round")
        
        # 2. 扳手 (全主題色)
        draw.line([(18, 46), (46, 18)], fill=color_accent, width=5)
        # 扳手左下端 (C形開口)
        draw.ellipse([(8, 40), (24, 56)], outline=color_accent, width=5)
        draw.line([(8, 48), (16, 56)], fill=(0, 0, 0, 0), width=6)
        # 扳手右上端 (C形開口)
        draw.ellipse([(40, 8), (56, 24)], outline=color_accent, width=5)
        draw.line([(48, 8), (56, 16)], fill=(0, 0, 0, 0), width=6)

    elif icon_name == "image_to_pdf":
        # 圖片 ➔ PDF 圖示 (左圖片、右文件、中間箭頭)
        draw_image_shape(draw, 6, 18, 30, 40, color_main)
        draw_doc_shape(draw, 34, 14, 58, 50, color_main, fold=10)
        # 繪製指向右邊的箭頭
        draw.line([(24, 46), (40, 46)], fill=color_accent, width=4)
        draw.line([(35, 41), (40, 46), (35, 51)], fill=color_accent, width=4, joint="round")
        # PDF 文件內部字串線條
        draw.line([(40, 26), (52, 26)], fill=color_accent, width=3)
        draw.line([(40, 34), (48, 34)], fill=color_accent, width=3)

    elif icon_name == "pdf_to_image":
        # PDF ➔ 圖片 圖示 (左文件、右圖片、中間箭頭)
        draw_doc_shape(draw, 6, 14, 30, 50, color_main, fold=10)
        draw_image_shape(draw, 34, 18, 58, 40, color_main)
        # 繪製指向右邊的箭頭
        draw.line([(24, 46), (40, 46)], fill=color_accent, width=4)
        draw.line([(35, 41), (40, 46), (35, 51)], fill=color_accent, width=4, joint="round")
        # PDF 文件內部字串線條
        draw.line([(12, 26), (24, 26)], fill=color_accent, width=3)
        draw.line([(12, 34), (20, 34)], fill=color_accent, width=3)

    elif icon_name == "pdf_split":
        # PDF 拆分圖示 (兩個錯位的文件，左上角帶有一把剪刀)
        draw_doc_shape(draw, 24, 10, 50, 44, color_accent, fold=10)
        draw_doc_shape(draw, 14, 20, 40, 54, color_main, fold=10)
        # 繪製小剪刀
        draw.ellipse([(6, 8), (14, 16)], outline=color_accent, width=3)
        draw.ellipse([(6, 18), (14, 26)], outline=color_accent, width=3)
        draw.line([(13, 13), (26, 21)], fill=color_accent, width=3)
        draw.line([(13, 21), (26, 13)], fill=color_accent, width=3)

    elif icon_name == "pdf_compress":
        # PDF 壓縮圖示 (文件在中央，上下各有指向文件內部的箭頭)
        draw_doc_shape(draw, 18, 12, 46, 52, color_main, fold=10)
        # 上方向下壓縮箭頭
        draw.line([(32, 2), (32, 10)], fill=color_accent, width=4)
        draw.line([(28, 6), (32, 10), (36, 6)], fill=color_accent, width=4, joint="round")
        # 下方向上壓縮箭頭
        draw.line([(32, 62), (32, 54)], fill=color_accent, width=4)
        draw.line([(28, 58), (32, 54), (36, 58)], fill=color_accent, width=4, joint="round")

    elif icon_name == "pdf_protect":
        # PDF 加密防護圖示 (文件外框 + 右下角精緻小鎖頭)
        draw_doc_shape(draw, 10, 12, 38, 52, color_main, fold=10)
        draw.line([(16, 26), (28, 26)], fill=color_main, width=3)
        draw.line([(16, 34), (24, 34)], fill=color_main, width=3)
        # 鎖頭本體
        draw.rounded_rectangle([34, 32, 56, 52], radius=4, fill=color_accent)
        # 鎖耳弧形與兩側插腳
        draw.arc([38, 20, 52, 34], start=180, end=360, fill=color_accent, width=4)
        draw.line([(38, 27), (38, 33)], fill=color_accent, width=4)
        draw.line([(52, 27), (52, 33)], fill=color_accent, width=4)

    elif icon_name == "image_compress":
        # 圖片壓縮圖示 (大相框重疊小相框，代表比例縮放)
        draw_image_shape(draw, 6, 12, 38, 44, color_main)
        draw_image_shape(draw, 30, 28, 58, 52, color_accent)

    elif icon_name == "about":
        # 關於本程式圖示 (經典圓圈與 i 字符號組合)
        draw.ellipse([(10, 10), (54, 54)], outline=color_main, width=4)
        # 資訊 i 的頭部圓點
        draw.ellipse([(30, 20), (34, 24)], fill=color_accent)
        # 資訊 i 的身體與底座
        draw.line([(32, 28), (32, 44)], fill=color_accent, width=4)
        draw.line([(28, 44), (36, 44)], fill=color_accent, width=4)

    elif icon_name == "import":
        # 拖曳匯入圖示 (大向下箭頭 + 底置托盤，提示使用者拖入此處)
        # 繪製底置托盤
        draw.line([(16, 44), (16, 52), (48, 52), (48, 44)], fill=color_main, width=4, joint="round")
        # 繪製指向托盤的向下大箭頭
        draw.line([(32, 12), (32, 40)], fill=color_accent, width=5)
        draw.line([(24, 32), (32, 40), (40, 32)], fill=color_accent, width=5, joint="round")

    elif icon_name == "settings":
        # 設定齒輪圖示 (中心圓圈 + 8個輻射突出齒輪齒，高度幾何對稱)
        # 內圓心與外齒輪框
        draw.ellipse([(24, 24), (40, 40)], outline=color_accent, width=4)
        draw.ellipse([(18, 18), (46, 46)], outline=color_accent, width=4)
        # 8個方位的齒輪齒
        # 90度方位
        draw.line([(32, 10), (32, 18)], fill=color_accent, width=5, joint="round")
        draw.line([(32, 46), (32, 54)], fill=color_accent, width=5, joint="round")
        draw.line([(10, 32), (18, 32)], fill=color_accent, width=5, joint="round")
        draw.line([(46, 32), (54, 32)], fill=color_accent, width=5, joint="round")
        # 45度方位
        draw.line([(16, 16), (22, 22)], fill=color_accent, width=5, joint="round")
        draw.line([(42, 42), (48, 48)], fill=color_accent, width=5, joint="round")
        draw.line([(16, 48), (22, 42)], fill=color_accent, width=5, joint="round")
        draw.line([(48, 16), (42, 22)], fill=color_accent, width=5, joint="round")

    elif icon_name == "run":
        # 開始執行圖示 (向右側的精緻播放三角實心按鈕，不分深淺色模式均使用純白色，以與按鈕文字完全融合)
        draw.polygon([
            (22, 16),
            (48, 32),
            (22, 48)
        ], fill=(255, 255, 255, 255))

    elif icon_name == "cancel":
        # 終止作業圖示 (不分深淺色模式均使用純白色線條，以與按鈕上的白色文字完全融合)
        # 外圓圈
        draw.ellipse([(16, 16), (48, 48)], outline=(255, 255, 255, 255), width=5)
        # 對角禁制斜線
        draw.line([(24, 24), (40, 40)], fill=(255, 255, 255, 255), width=5)

    elif icon_name == "plus":
        # 新增檔案加號圖示
        draw.line([(16, 32), (48, 32)], fill=color_accent, width=5, joint="round")
        draw.line([(32, 16), (32, 48)], fill=color_accent, width=5, joint="round")

    elif icon_name == "minus":
        # 移除檔案減號圖示
        draw.line([(16, 32), (48, 32)], fill=color_accent, width=5, joint="round")

    elif icon_name == "trash":
        # 垃圾桶圖示 (清空專用)
        # 蓋子與頂端提手
        draw.line([(14, 20), (50, 20)], fill=color_accent, width=4, joint="round")
        draw.line([(26, 20), (26, 14), (38, 14), (38, 20)], fill=color_accent, width=4, joint="round")
        # 垃圾桶本體框
        draw.line([(20, 20), (24, 54), (40, 54), (44, 20)], fill=color_accent, width=4, joint="round")
        # 桶身內部裝飾縱線
        draw.line([(28, 28), (28, 46)], fill=color_accent, width=3)
        draw.line([(36, 28), (36, 46)], fill=color_accent, width=3)

    elif icon_name == "expand":
        # 展開 PDF 圖示 (資料夾開啟示意)
        # 後方資料夾擋板 (主色)
        draw.line([(10, 48), (10, 22), (24, 22), (30, 28), (54, 28), (54, 48), (10, 48)], fill=color_main, width=4, joint="round")
        # 前方掀開的擋板 (點綴主題色，形成立體開口效果)
        draw.line([(10, 48), (16, 34), (58, 34), (54, 48)], fill=color_accent, width=4, joint="round")

    elif icon_name == "lock":
        # 安全防護鎖頭圖示 (大鎖頭，用於對話框標題)
        # 鎖身本體
        draw.rounded_rectangle([18, 28, 46, 52], radius=4, outline=color_accent, width=4)
        # 鎖耳
        draw.arc([22, 12, 42, 32], start=180, end=360, fill=color_accent, width=4)
        draw.line([(22, 22), (22, 28)], fill=color_accent, width=4)
        draw.line([(42, 22), (42, 28)], fill=color_accent, width=4)

    elif icon_name == "success":
        # 任務成功打勾圖示 (圓圈 + 內部Checkmark)
        # 外圓圈
        draw.ellipse([(12, 12), (52, 52)], outline=color_accent, width=5)
        # 勾勾
        draw.line([(22, 32), (30, 40), (44, 24)], fill=color_accent, width=5, joint="round")

    return img

def get_icon(icon_name, size=(20, 20)):
    """
    取得專屬於 CustomTkinter 的 CTkImage，能配合深/淺外觀模式自動切換顏色。
    
    參數:
      icon_name (str): 圖示鍵值名稱，如 'logo', 'settings', 'run' 等。
      size (tuple): 圖示最終在介面上渲染的尺寸 (width, height)，預設為 (20, 20)。
      
    回傳:
      ctk.CTkImage: 具備深淺主題切換相容性的 CTkImage 物件。
    """
    light_img = create_icon_image(icon_name, is_dark=False)
    dark_img = create_icon_image(icon_name, is_dark=True)
    return ctk.CTkImage(light_image=light_img, dark_image=dark_img, size=size)
