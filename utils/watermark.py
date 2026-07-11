import fitz

# 為指定 PDF 頁面添加文字浮水印
def apply_watermark_to_page(page, text, opacity, angle, is_tile, size, color_name):
    """
    對單一 PDF 頁面套用文字浮水印
    :param page: fitz.Page 物件
    :param text: 浮水印文字
    :param opacity: 透明度 (0.1 ~ 1.0)
    :param angle: 旋轉角度
    :param is_tile: 是否啟用平鋪 (滿版) 模式
    :param size: 字型大小
    :param color_name: 顏色名稱 ("灰色", "紅色", "藍色", "綠色", "黑色")
    """
    color_map = {
        "灰色": (0.5, 0.5, 0.5),
        "紅色": (0.9, 0.1, 0.1),
        "藍色": (0.1, 0.1, 0.9),
        "綠色": (0.1, 0.6, 0.1),
        "黑色": (0.0, 0.0, 0.0)
    }
    color = color_map.get(color_name, (0.5, 0.5, 0.5))
    
    try:
        # 使用對繁體中文支援最佳的內建 china-t 字型
        if is_tile:
            rect = page.rect
            w = int(rect.width)
            h = int(rect.height)
            
            # 根據字型大小動態決定 x 與 y 的網格間距
            x_step = max(150, size * 4)
            y_step = max(150, size * 3)
            
            # 使用雙重網格迴圈平鋪浮水印
            for x in range(30, w + 100, x_step):
                for y in range(50, h + 100, y_step):
                    page.insert_text((x, y), text, fontname="china-t", fontsize=size, 
                                     color=color, rotate=angle, fill_opacity=opacity, align=1)
        else:
            # 單一中心模式
            center = page.rect.center
            page.insert_text(center, text, fontname="china-t", fontsize=size, 
                             color=color, rotate=angle, fill_opacity=opacity, align=1)
    except Exception as e:
        print(f"浮水印添加失敗: {e}")
