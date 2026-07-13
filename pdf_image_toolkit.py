import os
import platform
import queue
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from tkinterdnd2 import DND_FILES

# 引入 utils 輔助函數與字型設定
from utils.helpers import (
    SYSTEM_FONT,
    FONT_OFFSET,
    ensure_app_icon,
)
from utils.icons import get_icon

# 引入對話框組件
from components.dialogs import (
    ModernPasswordDialog,
    ModernAboutDialog,
    ModernSuccessDialog,
)

# 引入功能分頁組件
from components.tab_image_to_pdf import TabImageToPDF
from components.tab_pdf_to_image import TabPDFToImage
from components.tab_pdf_split import TabPDFSplit
from components.tab_pdf_compress import TabPDFCompress
from components.tab_pdf_protect import TabPDFProtect
from components.tab_image_compress import TabImageCompress

# 系統 DPI 感知與拖放元件環境設定
try:
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ================== 📁 拖放與 CustomTkinter 結合 ==================
from tkinterdnd2 import TkinterDnD

class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    """繼承 CTk 與 DnDWrapper 以啟用拖放支援"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

# ================== 🚀 主入口程式類別 ==================
class PDFImageToolkit(CTkDnD):
    def __init__(self):
        super().__init__()
        
        self.title("PDF 圖片轉換小工具")
        
        # 智慧型視窗初始尺寸計算與置中
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        base_width = 1180
        base_height = 720
        height_factor = 0.70 if platform.system() == "Darwin" else 0.75
        
        target_width = min(base_width, int(screen_width * 0.9))
        target_height = min(base_height, int(screen_height * height_factor))
        
        x = (screen_width // 2) - (target_width // 2)
        y = (screen_height // 2) - (target_height // 2)
        if platform.system() == "Darwin":
            y = max(40, y - 20)
            
        self.geometry(f"{target_width}x{target_height}+{x}+{y}")
        self.minsize(1020, 640)
        
        # 自動等比放大設定，確保在高解析度螢幕上文字與畫面舒適
        ctk.set_widget_scaling(1.15)
        ctk.set_window_scaling(1.15)
        
        # 初始外觀設定
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        # 套用與設計專案 ICON
        icon_path = ensure_app_icon()
        if icon_path:
            try:
                if platform.system() == "Windows":
                    self.iconbitmap(icon_path)
                else:
                    # 跨平台 (macOS/Linux) 載入 PNG 比較相容
                    png_path = os.path.join(os.path.dirname(icon_path), "app_icon.png")
                    if os.path.exists(png_path):
                        self.iconphoto(True, tk.PhotoImage(file=png_path))
            except Exception as e:
                print(f"無法套用圖示: {e}")
        
        # 建立多執行緒 Queue
        self.queue = queue.Queue()
        
        # 初始化 UI
        self._build_main_ui()
        self._setup_treeview_styles()
        
        # 啟動 Queue 輪詢處理器
        self.after(100, self._process_queue)
        
    def _build_main_ui(self):
        # 左側側邊欄
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=("#F3F4F6", "#0F172A"))
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)
        
        # 側邊欄 LOGO 區
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.pack(fill=tk.X, padx=15, pady=(20, 30))
        
        # 建立內部容器，以讓 Logo 圖示與主標題能水平排列
        logo_title_frame = ctk.CTkFrame(logo_frame, fg_color="transparent")
        logo_title_frame.pack(anchor="w")
        
        # 建立獨立的 Logo 圖示與標題文字元件，徹底消除字串前怪異的空格，並更名為「PDF 圖片工具箱」
        logo_icon = get_icon("logo", size=(24, 24))
        logo_img_lbl = ctk.CTkLabel(logo_title_frame, text="", image=logo_icon)
        logo_img_lbl.pack(side=tk.LEFT)
        
        logo_txt_lbl = ctk.CTkLabel(logo_title_frame, text="PDF 圖片工具箱", 
                                     font=(SYSTEM_FONT, 15 + FONT_OFFSET, "bold"),
                                     text_color=("#2563EB", "#3B82F6"))
        logo_txt_lbl.pack(side=tk.LEFT, padx=(8, 0))
        
        # 版本資訊，排在標題下方
        version_lbl = ctk.CTkLabel(logo_frame, text="版本 v2.0.1", 
                                   font=(SYSTEM_FONT, 8 + FONT_OFFSET),
                                   text_color=("#9CA3AF", "#6B7280"))
        version_lbl.pack(anchor="w", pady=(2, 0))
        
        # 右側工作區
        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        self.pages = []
        self.nav_buttons = []
        
        # 導覽項目定義
        nav_items = [
            ("圖片/PDF ➔ PDF", TabImageToPDF, "image_to_pdf"),
            ("PDF ➔ 圖片", TabPDFToImage, "pdf_to_image"),
            ("PDF 拆分與擷取", TabPDFSplit, "pdf_split"),
            ("PDF 壓縮", TabPDFCompress, "pdf_compress"),
            ("PDF 加密防護", TabPDFProtect, "pdf_protect"),
            ("圖片壓縮與縮放", TabImageCompress, "image_compress"),
        ]
        
        for idx, (label, cls, icon_key) in enumerate(nav_items):
            page = cls(self.main_area, self)
            self.pages.append(page)
            
            icon_img = get_icon(icon_key, size=(20, 20))
            btn = ctk.CTkButton(
                self.sidebar, 
                text=f"  {label}",
                image=icon_img,
                compound="left",
                height=40,
                anchor="w",
                fg_color="transparent",
                text_color=("#374151", "#D1D5DB"),
                hover_color=("#E5E7EB", "#1E293B"),
                font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                command=lambda i=idx: self.select_page(i)
            )
            btn.pack(fill=tk.X, padx=12, pady=4)
            self.nav_buttons.append(btn)
            
        # 側邊欄底部控制區
        bottom_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=20)
        
        # 深/淺色主題 Switch
        self.theme_switch_var = tk.BooleanVar(value=True if ctk.get_appearance_mode() == "Dark" else False)
        self.theme_switch = ctk.CTkSwitch(
            bottom_frame, text="深色模式", 
            variable=self.theme_switch_var,
            command=self._toggle_theme,
            font=(SYSTEM_FONT, 9 + FONT_OFFSET)
        )
        self.theme_switch.pack(anchor="w", pady=(0, 15))
        
        # 關於按鈕
        about_icon = get_icon("about", size=(16, 16))
        about_btn = ctk.CTkButton(
            bottom_frame, text="  關於本程式", 
            image=about_icon,
            compound="left",
            height=32,
            fg_color=("#E5E7EB", "#1E293B"),
            text_color=("#374151", "#E5E7EB"),
            hover_color=("#D1D5DB", "#334155"),
            font=(SYSTEM_FONT, 9 + FONT_OFFSET, "bold"),
            command=self._show_about
        )
        about_btn.pack(fill=tk.X)
        
        # 設定別名對接 Queue 引用
        self.tab_image_to_pdf = self.pages[0]
        self.tab_pdf_to_image = self.pages[1]
        self.tab_pdf_split = self.pages[2]
        self.tab_pdf_compress = self.pages[3]
        self.tab_pdf_protect = self.pages[4]
        self.tab_image_compress = self.pages[5]
        
        # 預設載入首頁
        self.select_page(0)

    def select_page(self, idx):
        for page in self.pages:
            page.pack_forget()
        self.pages[idx].pack(fill=tk.BOTH, expand=True)
        
        for i, btn in enumerate(self.nav_buttons):
            if i == idx:
                btn.configure(
                    fg_color=("#3B82F6", "#2563EB"),
                    text_color="white",
                    hover_color=("#2563EB", "#1D4ED8")
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color=("#374151", "#D1D5DB"),
                    hover_color=("#E5E7EB", "#1E293B")
                )

    def _toggle_theme(self):
        if self.theme_switch_var.get():
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")
        self.after(100, self._setup_treeview_styles)

    def _show_about(self):
        ModernAboutDialog(self)

    # ================== 📥 QUEUE 訊息中央輪詢處理器 ==================
    def _process_queue(self):
        try:
            while True:
                kind, data = self.queue.get_nowait()
                
                # --- TAB 1 (圖片 ➔ PDF) 處理 ---
                if kind == "t1_status":
                    msg, val = data
                    self.tab_image_to_pdf.lbl_status.configure(text=msg, text_color=("#2563EB", "#3B82F6"))
                    self.tab_image_to_pdf.progress.set(val)
                elif kind == "t1_done":
                    save_path = data
                    self.tab_image_to_pdf.progress.set(1.0)
                    self.tab_image_to_pdf.lbl_status.configure(text="✨ PDF 檔案合併轉換成功！", text_color=("#059669", "#10B981"))
                    self.tab_image_to_pdf.is_converting = False
                    self.tab_image_to_pdf._toggle_ui_state("normal")
                    
                    dialog = ModernSuccessDialog(self, "任務成功完成", f"PDF 文件已成功產生並儲存！\n\n檔案名稱：{os.path.basename(save_path)}", file_path=save_path)
                    self.wait_window(dialog)
                elif kind == "t1_error":
                    err_msg = data
                    self.tab_image_to_pdf.lbl_status.configure(text=f"❌ 錯誤：{err_msg}", text_color=("#DC2626", "#EF4444"))
                    self.tab_image_to_pdf.is_converting = False
                    self.tab_image_to_pdf._toggle_ui_state("normal")
                    messagebox.showerror("錯誤", f"轉換失敗：\n{err_msg}")
                    
                # --- TAB 2 (PDF ➔ 圖片) 處理 ---
                elif kind == "t2_log":
                    self.tab_pdf_to_image.update_log(data)
                elif kind == "t2_progress":
                    self.tab_pdf_to_image.progress.set(data)
                elif kind == "t2_done":
                    out_dir = data
                    self.tab_pdf_to_image.progress.set(1.0)
                    self.tab_pdf_to_image.update_log("✨ 恭喜！所有 PDF 轉圖片作業已全部完成。")
                    self.tab_pdf_to_image.btn_cancel.pack_forget()
                    self.tab_pdf_to_image.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_to_image._toggle_ui_state("normal")
                    self.tab_pdf_to_image.is_converting = False
                    
                    dialog = ModernSuccessDialog(self, "任務成功完成", "所有 PDF 轉圖片作業已全部完成！", folder_path=out_dir)
                    self.wait_window(dialog)
                elif kind == "t2_cancelled":
                    self.tab_pdf_to_image.update_log("⚠️ 轉檔作業已被手動取消。")
                    self.tab_pdf_to_image.btn_cancel.pack_forget()
                    self.tab_pdf_to_image.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_to_image.btn_cancel.configure(state="normal")
                    self.tab_pdf_to_image._toggle_ui_state("normal")
                    self.tab_pdf_to_image.is_converting = False
                    messagebox.showinfo("取消", "作業已取消。")
                elif kind == "t2_error":
                    err_msg = data
                    self.tab_pdf_to_image.update_log(f"❌ 錯誤：{err_msg}")
                    self.tab_pdf_to_image.btn_cancel.pack_forget()
                    self.tab_pdf_to_image.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_to_image._toggle_ui_state("normal")
                    self.tab_pdf_to_image.is_converting = False
                    messagebox.showerror("錯誤", f"發生錯誤：\n{err_msg}")
                    
                # --- TAB 3 (PDF 拆分與擷取) 處理 ---
                elif kind == "t3_log":
                    self.tab_pdf_split.update_log(data)
                elif kind == "t3_progress":
                    self.tab_pdf_split.progress.set(data)
                elif kind == "t3_done":
                    out_dir = data
                    self.tab_pdf_split.progress.set(1.0)
                    self.tab_pdf_split.update_log("✨ 恭喜！PDF 拆分與擷取作業已全部完成。")
                    self.tab_pdf_split.btn_cancel.pack_forget()
                    self.tab_pdf_split.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_split._toggle_ui_state("normal")
                    self.tab_pdf_split.is_converting = False
                    
                    dialog = ModernSuccessDialog(self, "任務成功完成", "PDF 拆分與擷取作業已全部完成！", folder_path=out_dir)
                    self.wait_window(dialog)
                elif kind == "t3_cancelled":
                    self.tab_pdf_split.update_log("⚠️ 拆分作業已被手動取消。")
                    self.tab_pdf_split.btn_cancel.pack_forget()
                    self.tab_pdf_split.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_split.btn_cancel.configure(state="normal")
                    self.tab_pdf_split._toggle_ui_state("normal")
                    self.tab_pdf_split.is_converting = False
                    messagebox.showinfo("取消", "作業已取消。")
                elif kind == "t3_error":
                    err_msg = data
                    self.tab_pdf_split.update_log(f"❌ 錯誤：{err_msg}")
                    self.tab_pdf_split.btn_cancel.pack_forget()
                    self.tab_pdf_split.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_split._toggle_ui_state("normal")
                    self.tab_pdf_split.is_converting = False
                    messagebox.showerror("錯誤", f"發生錯誤：\n{err_msg}")
                    
                # --- TAB 4 (PDF 一鍵瘦身) 處理 ---
                elif kind == "t4_log":
                    self.tab_pdf_compress.update_log(data)
                elif kind == "t4_progress":
                    self.tab_pdf_compress.progress.set(data)
                elif kind == "t4_done":
                    save_path = data
                    self.tab_pdf_compress.progress.set(1.0)
                    self.tab_pdf_compress.update_log("✨ 恭喜！所有 PDF 壓縮優化作業已全部完成。")
                    self.tab_pdf_compress.btn_cancel.pack_forget()
                    self.tab_pdf_compress.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_compress._toggle_ui_state("normal")
                    self.tab_pdf_compress.is_converting = False
                    
                    dialog = ModernSuccessDialog(self, "任務成功完成", f"所有 PDF 壓縮優化與處理作業已全部完成！\n\n處理後檔案：{os.path.basename(save_path)}", file_path=save_path)
                    self.wait_window(dialog)
                elif kind == "t4_cancelled":
                    self.tab_pdf_compress.update_log("⚠️ 壓縮作業已被手動取消。")
                    self.tab_pdf_compress.btn_cancel.pack_forget()
                    self.tab_pdf_compress.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_compress.btn_cancel.configure(state="normal")
                    self.tab_pdf_compress._toggle_ui_state("normal")
                    self.tab_pdf_compress.is_converting = False
                    messagebox.showinfo("取消", "作業已取消。")
                elif kind == "t4_error":
                    err_msg = data
                    self.tab_pdf_compress.update_log(f"❌ 錯誤：{err_msg}")
                    self.tab_pdf_compress.btn_cancel.pack_forget()
                    self.tab_pdf_compress.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_compress._toggle_ui_state("normal")
                    self.tab_pdf_compress.is_converting = False
                    messagebox.showerror("錯誤", f"發生錯誤：\n{err_msg}")
                    
                # --- TAB 5 (PDF 加密防護) 處理 ---
                elif kind == "t5_log":
                    self.tab_pdf_protect.update_log(data)
                elif kind == "t5_progress":
                    self.tab_pdf_protect.progress.set(data)
                elif kind == "t5_done":
                    out_path = data
                    self.tab_pdf_protect.progress.set(1.0)
                    self.tab_pdf_protect.update_log("✨ 恭喜！所有 PDF 安全防護處理作業已全部完成。")
                    self.tab_pdf_protect.btn_cancel.pack_forget()
                    self.tab_pdf_protect.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_protect._toggle_ui_state("normal")
                    self.tab_pdf_protect.is_converting = False
                    
                    if out_path:
                        dialog = ModernSuccessDialog(
                            self, "任務成功完成", 
                            f"所有 PDF 安全防護與限制處理已完成！\n\n儲存目錄/檔案：{os.path.basename(out_path)}", 
                            file_path=out_path
                        )
                        self.wait_window(dialog)
                    else:
                        messagebox.showinfo("完成", "處理完成，但未產生任何檔案。")
                elif kind == "t5_cancelled":
                    self.tab_pdf_protect.update_log("⚠️ 防護處理作業已被手動取消。")
                    self.tab_pdf_protect.btn_cancel.pack_forget()
                    self.tab_pdf_protect.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_protect.btn_cancel.configure(state="normal")
                    self.tab_pdf_protect._toggle_ui_state("normal")
                    self.tab_pdf_protect.is_converting = False
                    messagebox.showinfo("取消", "作業已取消。")
                elif kind == "t5_error":
                    err_msg = data
                    self.tab_pdf_protect.update_log(f"❌ 錯誤：{err_msg}")
                    self.tab_pdf_protect.btn_cancel.pack_forget()
                    self.tab_pdf_protect.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_pdf_protect._toggle_ui_state("normal")
                    self.tab_pdf_protect.is_converting = False
                    messagebox.showerror("錯誤", f"發生錯誤：\n{err_msg}")
                    
                # --- TAB 6 (圖片壓縮) 處理 ---
                elif kind == "t6_log":
                    self.tab_image_compress.update_log(data)
                elif kind == "t6_progress":
                    self.tab_image_compress.progress.set(data)
                elif kind == "t6_done":
                    out_path = data
                    self.tab_image_compress.progress.set(1.0)
                    self.tab_image_compress.update_log("✨ 恭喜！所有圖片壓縮處理作業已全部完成。")
                    self.tab_image_compress.btn_cancel.pack_forget()
                    self.tab_image_compress.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_image_compress._toggle_ui_state("normal")
                    self.tab_image_compress.is_converting = False
                    
                    if out_path:
                        dialog = ModernSuccessDialog(
                            self, "任務成功完成", 
                            f"所有圖片壓縮與縮放處理作業已全部完成！\n\n儲存目錄/檔案：{os.path.basename(out_path)}", 
                            file_path=out_path
                        )
                        self.wait_window(dialog)
                    else:
                        messagebox.showinfo("完成", "處理完成，但未產生 any 檔案。")
                elif kind == "t6_cancelled":
                    self.tab_image_compress.update_log("⚠️ 壓縮處理作業已被手動取消。")
                    self.tab_image_compress.btn_cancel.pack_forget()
                    self.tab_image_compress.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_image_compress.btn_cancel.configure(state="normal")
                    self.tab_image_compress._toggle_ui_state("normal")
                    self.tab_image_compress.is_converting = False
                    messagebox.showinfo("取消", "作業已取消。")
                elif kind == "t6_error":
                    err_msg = data
                    self.tab_image_compress.update_log(f"❌ 錯誤：{err_msg}")
                    self.tab_image_compress.btn_cancel.pack_forget()
                    self.tab_image_compress.btn_run.pack(fill=tk.BOTH, expand=True)
                    self.tab_image_compress._toggle_ui_state("normal")
                    self.tab_image_compress.is_converting = False
                    messagebox.showerror("錯誤", f"發生錯誤：\n{err_msg}")

                # --- 共用密碼解鎖請求 ---
                elif kind == "ask_pw":
                    path, evt, res = data
                    dialog = ModernPasswordDialog(self, os.path.basename(path))
                    self.wait_window(dialog)
                    res["pw"] = dialog.password
                    evt.set()
                elif kind == "error_msg":
                    messagebox.showerror("錯誤", data)
                    
                self.queue.task_done()
        except queue.Empty:
            pass
        finally:
            self.after(100, self._process_queue)

    def _setup_treeview_styles(self):
        """為 Treeview 元件設置美觀的配色樣式，適配深色與淺色模式"""
        style = ttk.Style(self)
        
        # 依據 CustomTkinter 的外觀模式來決定配色
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            bg_color = "#1E293B"
            fg_color = "#F1F5F9"
            selected_bg = "#2563EB"
            selected_fg = "#FFFFFF"
            border_color = "#334155"
            header_bg = "#0F172A"
            header_fg = "#FFFFFF"
        else:
            bg_color = "#FFFFFF"
            fg_color = "#0F172A"
            selected_bg = "#EFF6FF"
            selected_fg = "#1E40AF"
            border_color = "#E2E8F0"
            header_bg = "#F8FAFC"
            header_fg = "#1E293B"
            
        for tag in ["T1", "T2", "T3", "T4", "T5", "T6"]:
            style.theme_use("clam")
            style.configure(f"{tag}.Treeview",
                            background=bg_color,
                            foreground=fg_color,
                            rowheight=54 if tag == "T1" else 36,
                            fieldbackground=bg_color,
                            bordercolor=border_color,
                            borderwidth=1,
                            font=(SYSTEM_FONT, 10 + FONT_OFFSET))
            
            style.map(f"{tag}.Treeview",
                      background=[('selected', selected_bg)],
                      foreground=[('selected', selected_fg)])
            
            style.configure(f"{tag}.Treeview.Heading",
                            background=header_bg,
                            foreground=header_fg,
                            font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                            borderwidth=0)

# ================== 🏁 啟動主程式 ==================
if __name__ == "__main__":
    app = PDFImageToolkit()
    app.mainloop()
