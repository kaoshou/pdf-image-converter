import tkinter as tk
import customtkinter as ctk
from utils.helpers import SYSTEM_FONT, FONT_OFFSET

class PDFFeaturesFrame(ctk.CTkFrame):
    """
    可重用的 PDF 高級功能設定面板元件 (包含加密防護、元資料、文字浮水印)
    適配於 Tab 1、Tab 3 與 Tab 4，落實物件導向重用原則。
    """
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        # 初始化 UI 元件
        self._build_ui()
        
    def _build_ui(self):
        # ================== 1. PDF 加密防護 ==================
        self.encrypt_var = tk.BooleanVar(value=False)
        self.check_encrypt = ctk.CTkCheckBox(
            self, text="PDF 加密防護", 
            variable=self.encrypt_var, 
            command=self._toggle_encrypt, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_encrypt.pack(anchor="w", pady=4)
        
        # 密碼欄位 (預設有明顯的提醒 placeholder 文字)
        self.entry_pw = ctk.CTkEntry(
            self, 
            placeholder_text="設定開啟密碼（加密保護，可選）", 
            show="●", 
            height=30, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.entry_pw.pack(fill=tk.X, pady=(2, 8))
        
        # ================== 2. PDF 資訊 / 元資料 (Metadata) ==================
        lbl_meta = ctk.CTkLabel(self, text="PDF 資訊 (Metadata):", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_meta.pack(anchor="w", pady=(4, 1))
        
        self.meta_title = ctk.CTkEntry(self, placeholder_text="標題 (Title)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.meta_title.pack(fill=tk.X, pady=2)
        
        self.meta_author = ctk.CTkEntry(self, placeholder_text="作者 (Author)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.meta_author.pack(fill=tk.X, pady=2)
        
        self.meta_subject = ctk.CTkEntry(self, placeholder_text="主題 (Subject)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.meta_subject.pack(fill=tk.X, pady=2)
        
        self.meta_keywords = ctk.CTkEntry(self, placeholder_text="關鍵字 (Keywords)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.meta_keywords.pack(fill=tk.X, pady=(2, 8))
        
        # ================== 3. PDF 智慧文字浮水印 ==================
        self.watermark_var = tk.BooleanVar(value=False)
        self.check_watermark = ctk.CTkCheckBox(
            self, text="啟用文字浮水印", 
            variable=self.watermark_var,
            command=self._toggle_watermark,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_watermark.pack(anchor="w", pady=4)
        
        self.entry_wm_text = ctk.CTkEntry(
            self, 
            placeholder_text="輸入浮水印文字 (例如: 機密檔案)", 
            height=30, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.entry_wm_text.pack(fill=tk.X, pady=2)
        
        # 平鋪勾選框
        self.wm_tile_var = tk.BooleanVar(value=False)
        self.check_wm_tile = ctk.CTkCheckBox(
            self, text="平鋪防偽浮水印 (滿版)", 
            variable=self.wm_tile_var, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_wm_tile.pack(anchor="w", pady=4)

        # 顏色選擇
        self.wm_color_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.wm_color_frame.pack(fill=tk.X, pady=(2, 4))
        lbl_wm_color = ctk.CTkLabel(self.wm_color_frame, text="浮水印顏色:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_wm_color.pack(side=tk.LEFT)
        self.combo_wm_color = ctk.CTkOptionMenu(
            self.wm_color_frame, values=["灰色", "紅色", "藍色", "綠色", "黑色"], 
            height=30, width=120
        )
        self.combo_wm_color.pack(side=tk.RIGHT)
        self.combo_wm_color.set("灰色")

        # 字型大小選擇
        self.wm_size_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.wm_size_frame.pack(fill=tk.X, pady=(2, 4))
        lbl_wm_size = ctk.CTkLabel(self.wm_size_frame, text="字型大小:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_wm_size.pack(side=tk.LEFT)
        self.combo_wm_size = ctk.CTkOptionMenu(
            self.wm_size_frame, values=["24", "32", "40", "48", "64", "80"], 
            height=30, width=120
        )
        self.combo_wm_size.pack(side=tk.RIGHT)
        self.combo_wm_size.set("48")

        # 透明度 Slider
        self.wm_opacity_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.wm_opacity_frame.pack(fill=tk.X, pady=(2, 4))
        
        self.lbl_wm_opacity = ctk.CTkLabel(self.wm_opacity_frame, text="透明度: 0.3", font=(SYSTEM_FONT, 9 + FONT_OFFSET))
        self.lbl_wm_opacity.pack(side=tk.LEFT)
        
        self.slider_wm_opacity = ctk.CTkSlider(
            self.wm_opacity_frame, from_=0.1, to=1.0, 
            number_of_steps=18, height=16,
            command=self._update_wm_opacity_lbl
        )
        self.slider_wm_opacity.set(0.3)
        self.slider_wm_opacity.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        # 旋轉角度 OptionMenu
        self.wm_angle_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.wm_angle_frame.pack(fill=tk.X, pady=(2, 8))
        
        lbl_wm_angle = ctk.CTkLabel(self.wm_angle_frame, text="旋轉角度:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_wm_angle.pack(side=tk.LEFT)
        
        self.combo_wm_angle = ctk.CTkOptionMenu(
            self.wm_angle_frame, values=["0", "15", "30", "45", "60", "90"], 
            height=30, width=120
        )
        self.combo_wm_angle.pack(side=tk.RIGHT)
        self.combo_wm_angle.set("45")
        
        # 初始化控制元件啟用狀態
        self._toggle_encrypt()
        self._toggle_watermark()

    def _toggle_encrypt(self):
        """加密核取方塊連動"""
        if self.encrypt_var.get():
            self.entry_pw.configure(state="normal")
        else:
            self.entry_pw.delete(0, tk.END)
            self.entry_pw.configure(state="disabled")

    def _toggle_watermark(self):
        """浮水印核取方塊連動"""
        state = "normal" if self.watermark_var.get() else "disabled"
        slider_state = tk.NORMAL if self.watermark_var.get() else tk.DISABLED
        
        self.entry_wm_text.configure(state=state)
        self.slider_wm_opacity.configure(state=slider_state if slider_state == tk.NORMAL else "disabled")
        self.combo_wm_angle.configure(state=state)
        self.check_wm_tile.configure(state=state)
        self.combo_wm_color.configure(state=state)
        self.combo_wm_size.configure(state=state)
        
        self.lbl_wm_opacity.configure(text_color=("black", "white") if self.watermark_var.get() else "#9CA3AF")

    def _update_wm_opacity_lbl(self, val):
        """更新透明度數值標籤"""
        self.lbl_wm_opacity.configure(text=f"透明度: {val:.2f}")

    def get_settings(self) -> dict:
        """
        獲取目前面板的所有 PDF 功能設定參數
        """
        try:
            wm_angle = int(self.combo_wm_angle.get())
        except:
            wm_angle = 45
            
        try:
            wm_size = int(self.combo_wm_size.get())
        except:
            wm_size = 48
            
        return {
            "encrypt": self.encrypt_var.get(),
            "password": self.entry_pw.get().strip(),
            "meta_title": self.meta_title.get().strip(),
            "meta_author": self.meta_author.get().strip(),
            "meta_subject": self.meta_subject.get().strip(),
            "meta_keywords": self.meta_keywords.get().strip(),
            "watermark": self.watermark_var.get(),
            "wm_text": self.entry_wm_text.get().strip(),
            "wm_tile": self.wm_tile_var.get(),
            "wm_color": self.combo_wm_color.get(),
            "wm_size": wm_size,
            "wm_opacity": self.slider_wm_opacity.get(),
            "wm_angle": wm_angle
        }

    def configure_state(self, state: str):
        """
        一鍵啟用/禁用面板內的所有控制項 (狀態傳入 'normal' 或 'disabled')
        """
        mode = "normal" if state == "normal" else "disabled"
        self.check_encrypt.configure(state=mode)
        self.meta_title.configure(state=mode)
        self.meta_author.configure(state=mode)
        self.meta_subject.configure(state=mode)
        self.meta_keywords.configure(state=mode)
        self.check_watermark.configure(state=mode)
        
        if mode == "normal":
            self._toggle_encrypt()
            self._toggle_watermark()
        else:
            self.entry_pw.configure(state="disabled")
            self.entry_wm_text.configure(state="disabled")
            self.slider_wm_opacity.configure(state=tk.DISABLED)
            self.combo_wm_angle.configure(state="disabled")
            self.check_wm_tile.configure(state="disabled")
            self.combo_wm_color.configure(state="disabled")
            self.combo_wm_size.configure(state="disabled")
