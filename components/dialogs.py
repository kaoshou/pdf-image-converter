import os
import platform
import webbrowser
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from utils.helpers import SYSTEM_FONT, FONT_OFFSET
from utils.icons import get_icon
from PIL import Image, ImageTk, ImageOps
import io
import threading

# ================== 🔐 密碼解鎖對話框 (CustomTkinter 風格) ==================
class ModernPasswordDialog(ctk.CTkToplevel):
    """用於解鎖加密 PDF 的密碼輸入彈窗"""
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.title("安全性驗證")
        self.geometry("450x230")
        self.resizable(False, False)
        self.configure(fg_color=("#F3F4F6", "#111827"))  # 淺色/深色底色
        self.transient(parent)
        self.grab_set()

        self.password = None

        # 置中視窗
        try:
            self.update_idletasks()
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 225
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 115
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        # 容器排版
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 建立大鎖頭圖示，置於文字左側，移除原先的 Emoji
        lock_icon = get_icon("lock", size=(24, 24))
        title_lbl = ctk.CTkLabel(container, text="  檔案受密碼保護", 
                                 image=lock_icon,
                                 compound="left",
                                 font=(SYSTEM_FONT, 14 + FONT_OFFSET, "bold"),
                                 text_color=("#2563EB", "#3B82F6"))
        title_lbl.pack(anchor="w", pady=(0, 6))

        info_lbl = ctk.CTkLabel(container, text=f"檔案「{filename}」需要開啟密碼：",
                                font=(SYSTEM_FONT, 11 + FONT_OFFSET),
                                text_color=("#4B5563", "#D1D5DB"),
                                justify="left")
        info_lbl.pack(anchor="w", pady=(0, 12))

        self.entry = ctk.CTkEntry(container, show="●", 
                                  placeholder_text="請輸入密碼",
                                  font=(SYSTEM_FONT, 11 + FONT_OFFSET),
                                  height=36)
        self.entry.pack(fill=tk.X, pady=(0, 20))
        self.entry.focus_set()

        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill=tk.X)

        cancel_btn = ctk.CTkButton(btn_frame, text="略過此檔", 
                                   width=90, height=32,
                                   fg_color=("#E5E7EB", "#374151"),
                                   text_color=("#374151", "#F9FAFB"),
                                   hover_color=("#D1D5DB", "#4B5563"),
                                   font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                   command=self.on_cancel)
        cancel_btn.pack(side=tk.RIGHT, padx=(10, 0))

        confirm_btn = ctk.CTkButton(btn_frame, text="確認解鎖 🔓", 
                                    width=110, height=32,
                                    fg_color=("#2563EB", "#3B82F6"),
                                    text_color="white",
                                    hover_color=("#1D4ED8", "#2563EB"),
                                    font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                    command=self.on_submit)
        confirm_btn.pack(side=tk.RIGHT)

        self.bind('<Return>', lambda e: self.on_submit())
        self.bind('<Escape>', lambda e: self.on_cancel())

    def on_submit(self):
        self.password = self.entry.get()
        self.destroy()

    def on_cancel(self):
        self.password = None
        self.destroy()


# ================== ℹ️ 關於視窗 (CustomTkinter 風格) ==================
class ModernAboutDialog(ctk.CTkToplevel):
    """軟體關於資訊與開源聲明對話框"""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("關於本程式")
        self.geometry("560x500")
        self.resizable(False, False)
        self.transient(parent)
        self.configure(fg_color=("#F9FAFB", "#111827"))
        
        try:
            self.update_idletasks()
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 280
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 250
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        title_lbl = ctk.CTkLabel(container, text="PDF 圖片工具箱", 
                                 font=(SYSTEM_FONT, 16 + FONT_OFFSET, "bold"),
                                 text_color=("#1E3A8A", "#60A5FA"))
        title_lbl.pack(anchor="w", pady=(0, 2))
        
        ver_lbl = ctk.CTkLabel(container, text="版本: 2.0.1", 
                               font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                               text_color=("#6B7280", "#9CA3AF"))
        ver_lbl.pack(anchor="w", pady=(0, 15))

        # 開發者名片卡區 (淡背景框)
        card = ctk.CTkFrame(container, fg_color=("#F3F4F6", "#1F2937"), border_width=1, border_color=("#E5E7EB", "#374151"))
        card.pack(fill=tk.X, pady=(0, 15), padx=2)
        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill=tk.X, padx=15, pady=12)

        dev_lbl = ctk.CTkLabel(card_inner, text="開發者: 鄭郁翰 (Cheng, Yu-Han)", font=(SYSTEM_FONT, 11 + FONT_OFFSET, "bold"))
        dev_lbl.pack(anchor="w")
        
        email_lbl = ctk.CTkLabel(card_inner, text="Email: kaoshou@gmail.com", font=(SYSTEM_FONT, 11 + FONT_OFFSET))
        email_lbl.pack(anchor="w", pady=(2, 4))

        gh_frame = ctk.CTkFrame(card_inner, fg_color="transparent", height=24)
        gh_frame.pack(fill=tk.X)
        gh_lbl = ctk.CTkLabel(gh_frame, text="GitHub: ", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        gh_lbl.pack(side=tk.LEFT)
        gh_link = ctk.CTkLabel(gh_frame, text="https://github.com/kaoshou/pdf-image-converter", 
                               font=(SYSTEM_FONT, 10 + FONT_OFFSET, "underline"),
                               text_color=("#2563EB", "#3B82F6"), cursor="hand2")
        gh_link.pack(side=tk.LEFT)
        gh_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/kaoshou/pdf-image-converter"))

        divider = ctk.CTkFrame(container, fg_color=("#E5E7EB", "#374151"), height=1)
        divider.pack(fill=tk.X, pady=5)

        third_party_lbl = ctk.CTkLabel(container, text="第三方套件與開源聲明 (Open Source Disclosure):", 
                                       font=(SYSTEM_FONT, 11 + FONT_OFFSET, "bold"),
                                       text_color=("#374151", "#E5E7EB"))
        third_party_lbl.pack(anchor="w", pady=(10, 5))

        license_desc = (
            "本程式核心功能基於以下優秀開源專案實作：\n\n"
            "• PyMuPDF (fitz) - 採用 GNU AGPL v3.0 授權\n"
            "  用以處理所有 PDF 頁面渲染、合併、加密與平面化等核心操作。\n"
            "  專案網址：https://github.com/pymupdf/PyMuPDF\n\n"
            "• CustomTkinter - 採用 MIT 授權\n"
            "  用以打造現代化扁平設計、自適應且支援深淺色切換的 UI 元件。\n"
            "  專案網址：https://github.com/TomSchimansky/CustomTkinter\n\n"
            "• TkinterDnD2 - 採用 MIT 授權\n"
            "  用以支援作業系統檔案一鍵拖放 (Drag and Drop) 匯入之功能。\n"
            "  專案網址：https://github.com/pmgagne/tkinterdnd2\n\n"
            "免責聲明：本軟體完全依「現狀」提供，開發者對於因使用本程式所產生的任何資料損毀、遺失或商業損失概不負責。"
        )

        text_box = ctk.CTkTextbox(container, font=("Consolas", 9 + FONT_OFFSET), 
                                  fg_color=("#F3F4F6", "#1F2937"),
                                  text_color=("#374151", "#D1D5DB"),
                                  wrap=tk.WORD, height=140)
        text_box.insert(tk.END, license_desc)
        text_box.configure(state="disabled")
        text_box.pack(fill=tk.X, pady=(0, 15))

        close_btn = ctk.CTkButton(container, text="關閉視窗", 
                                  width=100, height=32,
                                  fg_color=("#E5E7EB", "#374151"),
                                  text_color=("#374151", "#F9FAFB"),
                                  hover_color=("#D1D5DB", "#4B5563"),
                                  font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                  command=self.destroy)
        close_btn.pack(pady=(10, 0))


# ================== 🎉 任務成功完成對話框 (CustomTkinter 風格) ==================
class ModernSuccessDialog(ctk.CTkToplevel):
    """高級成果提示視窗，提供立即開啟檔案、開啟資料夾與複製路徑功能"""
    def __init__(self, parent, title, message, file_path=None, folder_path=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("520x250")
        self.resizable(False, False)
        self.configure(fg_color=("#F3F4F6", "#111827"))
        self.transient(parent)
        self.grab_set()

        self.file_path = file_path
        self.folder_path = folder_path if folder_path else (os.path.dirname(file_path) if file_path else None)

        # 置中視窗
        try:
            self.update_idletasks()
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 260
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 125
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill=tk.BOTH, expand=True, padx=24, pady=24)

        # 建立成功勾勾圖示，置於文字左側，移除原先的 Emoji
        success_icon = get_icon("success", size=(24, 24))
        title_lbl = ctk.CTkLabel(container, text="  任務成功完成", 
                                 image=success_icon,
                                 compound="left",
                                 font=(SYSTEM_FONT, 15 + FONT_OFFSET, "bold"),
                                 text_color=("#10B981", "#34D399"))
        title_lbl.pack(anchor="w", pady=(0, 8))

        msg_lbl = ctk.CTkLabel(container, text=message,
                               font=(SYSTEM_FONT, 11 + FONT_OFFSET),
                               text_color=("#374151", "#E5E7EB"),
                               justify="left", wraplength=460)
        msg_lbl.pack(anchor="w", pady=(0, 20))

        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # 關閉按鈕
        close_btn = ctk.CTkButton(btn_frame, text="關閉", 
                                   width=90, height=32,
                                   fg_color=("#E5E7EB", "#374151"),
                                   text_color=("#374151", "#F9FAFB"),
                                   hover_color=("#D1D5DB", "#4B5563"),
                                   font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                   command=self.destroy)
        close_btn.pack(side=tk.RIGHT, padx=(10, 0))

        # 複製路徑按鈕
        if self.file_path:
            copy_btn = ctk.CTkButton(btn_frame, text="複製檔案路徑", 
                                      width=110, height=32,
                                      fg_color=("#F3F4F6", "#1F2937"),
                                      text_color=("#2563EB", "#3B82F6"),
                                      hover_color=("#E5E7EB", "#374151"),
                                      border_width=1, border_color=("#2563EB", "#3B82F6"),
                                      font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                      command=self.on_copy)
            copy_btn.pack(side=tk.RIGHT, padx=(10, 0))

        # 開啟資料夾按鈕
        if self.folder_path:
            open_folder_btn = ctk.CTkButton(btn_frame, text="開啟資料夾", 
                                             width=100, height=32,
                                             fg_color=("#F3F4F6", "#1F2937"),
                                             text_color=("#059669", "#10B981"),
                                             hover_color=("#E5E7EB", "#374151"),
                                             border_width=1, border_color=("#059669", "#10B981"),
                                             font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                             command=self.on_open_folder)
            open_folder_btn.pack(side=tk.RIGHT, padx=(10, 0))

        # 立即開啟檔案按鈕
        if self.file_path:
            open_file_btn = ctk.CTkButton(btn_frame, text="立即開啟", 
                                           width=100, height=32,
                                           fg_color=("#2563EB", "#3B82F6"),
                                           text_color="white",
                                           hover_color=("#1D4ED8", "#2563EB"),
                                           font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                           command=self.on_open_file)
            open_file_btn.pack(side=tk.RIGHT)

    def on_copy(self):
        self.clipboard_clear()
        self.clipboard_append(self.file_path)
        self.update()
        messagebox.showinfo("提示", "已成功複製檔案路徑到剪貼簿！")

    def on_open_folder(self):
        try:
            if platform.system() == "Windows":
                os.startfile(self.folder_path)
            else:
                webbrowser.open(f"file://{self.folder_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟資料夾：{e}")

    def on_open_file(self):
        try:
            if platform.system() == "Windows":
                os.startfile(self.file_path)
            else:
                webbrowser.open(f"file://{self.file_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟檔案：{e}")


# ================== 🖼️ 圖片壓縮雙欄對比預覽對話框 (CustomTkinter 風格) ==================
class ModernCompressPreviewDialog(ctk.CTkToplevel):
    """即時展示圖片壓縮前與壓縮後的雙欄對比預覽彈窗"""
    def __init__(self, parent, img_path, settings):
        super().__init__(parent)
        self.title("壓縮效果即時預覽")
        self.geometry("900x560")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.img_path = img_path
        self.settings = settings
        
        # 置中視窗
        try:
            self.update_idletasks()
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 450
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 280
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        # 頂部狀態/載入中提示
        # 頂部狀態/載入中提示 (已移除 Emoji)
        self.title_lbl = ctk.CTkLabel(self, text="正在模擬壓縮處理中，請稍後...", 
                                      font=(SYSTEM_FONT, 14 + FONT_OFFSET, "bold"),
                                      text_color=("#2563EB", "#3B82F6"))
        self.title_lbl.pack(pady=(15, 10))

        # 中間雙欄對比卡片
        self.comparison_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.comparison_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        # 左邊：原始圖片區
        self.left_frame = ctk.CTkFrame(self.comparison_frame, border_width=1, border_color=("#E5E7EB", "#374151"))
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 原始圖片載入中 (已移除 Emoji)
        self.left_img_lbl = ctk.CTkLabel(self.left_frame, text="載入中...", fg_color="black")
        self.left_img_lbl.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.left_info_lbl = ctk.CTkLabel(self.left_frame, text="原圖載入中...", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.left_info_lbl.pack(pady=8)
        
        # 右邊：壓縮預覽區
        self.right_frame = ctk.CTkFrame(self.comparison_frame, border_width=1, border_color=("#E5E7EB", "#374151"))
        self.right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # 壓縮預覽處理中 (已移除 Emoji)
        self.right_img_lbl = ctk.CTkLabel(self.right_frame, text="處理中...", fg_color="black")
        self.right_img_lbl.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.right_info_lbl = ctk.CTkLabel(self.right_frame, text="預估壓縮效果處理中...", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"), text_color=("#059669", "#10B981"))
        self.right_info_lbl.pack(pady=8)
        
        # 底部關閉按鈕
        self.close_btn = ctk.CTkButton(self, text="關閉預覽", width=120, height=36,
                                       fg_color=("#E5E7EB", "#374151"),
                                       text_color=("#374151", "#F9FAFB"),
                                       hover_color=("#D1D5DB", "#4B5563"),
                                       font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                       command=self.destroy)
        self.close_btn.pack(pady=(10, 15))
        
        # 保存圖片的 reference 防止 GC 回收
        self.orig_photo = None
        self.comp_photo = None
        
        # 啟動背景處理線程
        threading.Thread(target=self._process_preview, daemon=True).start()

    def _process_preview(self):
        import io
        try:
            orig_size = os.path.getsize(self.img_path)
            
            with Image.open(self.img_path) as img:
                # 1. 修正方向
                if self.settings["exif_transpose"]:
                    img = ImageOps.exif_transpose(img)
                
                orig_w, orig_h = img.size
                
                # 2. 縮放
                mode = self.settings["resize_mode"]
                new_w, new_h = orig_w, orig_h
                
                if mode == "設定百分比 (%)":
                    new_w = int(orig_w * self.settings["percent"] / 100.0)
                    new_h = int(orig_h * self.settings["percent"] / 100.0)
                elif mode == "設定寬度 (高度自適應)":
                    new_w = self.settings["width"]
                    new_h = int(orig_h * (new_w / orig_w))
                elif mode == "設定高度 (寬度自適應)":
                    new_h = self.settings["height"]
                    new_w = int(orig_w * (new_h / orig_h))
                elif mode == "設定固定寬高 (像素)":
                    new_w = self.settings["width"]
                    new_h = self.settings["height"]
                    
                new_w = max(1, new_w)
                new_h = max(1, new_h)
                
                if (new_w, new_h) != (orig_w, orig_h):
                    img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                else:
                    img_resized = img.copy()
                    
                # 3. 旋轉
                if self.settings["rotate"] and self.settings["rotate_angle"] != 0:
                    img_resized = img_resized.rotate(self.settings["rotate_angle"], expand=True)
                    
                # 4. 模擬品質與色彩量化
                fmt = self.settings["format"]
                orig_ext = os.path.splitext(self.img_path)[1].lower()
                if fmt == "保持原格式":
                    save_format = "PNG" if orig_ext == ".png" else ("WEBP" if orig_ext == ".webp" else "JPEG")
                else:
                    save_format = fmt
                    
                buffer = io.BytesIO()
                save_kwargs = {"optimize": True}
                if save_format == "PNG":
                    save_kwargs["compress_level"] = 9
                if save_format in ["JPEG", "WEBP"]:
                    save_kwargs["quality"] = self.settings["quality"]
                    
                if save_format == "JPEG" and img_resized.mode in ("RGBA", "P"):
                    img_to_save = img_resized.convert("RGB")
                elif save_format == "PNG" and self.settings["png_quantize"]:
                    img_to_save = img_resized.quantize(colors=256)
                else:
                    img_to_save = img_resized
                    
                # 儲存至記憶體 buffer
                img_to_save.save(buffer, save_format, **save_kwargs)
                compressed_bytes = buffer.getvalue()
                
                # 計算大小
                comp_size = len(compressed_bytes)
                saved_size = orig_size - comp_size
                saved_pct = (saved_size / orig_size) * 100 if orig_size > 0 else 0
                
                # 載入壓縮後的 PhotoImage 用以展示
                with Image.open(io.BytesIO(compressed_bytes)) as comp_img:
                    # 將原圖與壓縮圖縮小至適合顯示的大小 (400x320)
                    orig_show = ImageOps.contain(img, (400, 320))
                    comp_show = ImageOps.contain(comp_img, (400, 320))
                    
                    self.orig_photo = ImageTk.PhotoImage(orig_show)
                    self.comp_photo = ImageTk.PhotoImage(comp_show)
                    
                # 在主線程更新 UI
                self.after(0, self._update_ui, orig_w, orig_h, orig_size, comp_img.width, comp_img.height, comp_size, saved_pct)
                
        except Exception as e:
            self.after(0, self._show_error, str(e))

    def _update_ui(self, orig_w, orig_h, orig_size, comp_w, comp_h, comp_size, saved_pct):
        self.title_lbl.configure(text="✅ 預覽壓縮處理完成", text_color=("#059669", "#10B981"))
        
        # 更新原圖圖片與標籤
        self.left_img_lbl.configure(image=self.orig_photo, text="")
        sz_mb = orig_size / (1024*1024)
        orig_sz_str = f"{sz_mb:.2f} MB" if sz_mb >= 1.0 else f"{orig_size/1024:.1f} KB"
        self.left_info_lbl.configure(text=f"原圖尺寸: {orig_w}x{orig_h} | 大小: {orig_sz_str}")
        
        # 更新壓縮圖圖片與標籤
        self.right_img_lbl.configure(image=self.comp_photo, text="")
        csz_mb = comp_size / (1024*1024)
        comp_sz_str = f"{csz_mb:.2f} MB" if csz_mb >= 1.0 else f"{comp_size/1024:.1f} KB"
        
        if saved_pct > 0:
            pct_str = f" (體積減少 {saved_pct:.1f}%)"
        else:
            pct_str = " (體積增加)" if saved_pct < 0 else " (無變化)"
            
        self.right_info_lbl.configure(text=f"預估尺寸: {comp_w}x{comp_h} | 預估大小: {comp_sz_str}{pct_str}")

    def _show_error(self, err_msg):
        self.title_lbl.configure(text="❌ 模擬處理失敗", text_color=("#DC2626", "#EF4444"))
        self.left_img_lbl.configure(text="無法讀取原圖")
        self.right_img_lbl.configure(text=f"錯誤: {err_msg}")
