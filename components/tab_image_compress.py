import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageOps
from tkinterdnd2 import DND_FILES

from utils.helpers import SYSTEM_FONT, FONT_OFFSET, parse_dropped_files, unique_filename

class TabImageCompress(ctk.CTkFrame):
    """圖片批次壓縮、縮放、旋轉與 EXIF 修改的功能分頁"""
    def __init__(self, master, app):
        super().__init__(master, fg_color="transparent")
        self.app = app
        
        # 功能內部變數
        self.selected_files = []
        self.is_converting = False
        self.stop_event = threading.Event()
        
        # 初始化 UI
        self._build_ui()

    def _build_ui(self):
        # 左右分割配置
        t6_main = ctk.CTkFrame(self, fg_color="transparent")
        t6_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：圖片檔案清單
        left_pane = ctk.CTkFrame(t6_main, fg_color="transparent")
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        lbl_list = ctk.CTkLabel(left_pane, text="待處理圖片檔案清單:", font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        lbl_list.pack(anchor="w", pady=(0, 5))
        
        # Treeview 清單 (加強樣式，有虛擬滾動條與框線)
        tree_frame = ctk.CTkFrame(left_pane, border_width=1, border_color=("#E5E7EB", "#374151"))
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(
            tree_frame, 
            columns=("no", "name", "res", "size", "status"), 
            show="headings", 
            selectmode="extended",
            style="T6.Treeview"
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 設定表頭與寬度
        self.tree.heading("no", text="序號")
        self.tree.heading("name", text="檔案名稱")
        self.tree.heading("res", text="原始解析度")
        self.tree.heading("size", text="原始體積")
        self.tree.heading("status", text="狀態")
        
        self.tree.column("no", width=50, anchor="center")
        self.tree.column("name", width=250, anchor="w")
        self.tree.column("res", width=120, anchor="center")
        self.tree.column("size", width=90, anchor="center")
        self.tree.column("status", width=90, anchor="center")
        
        # 綁定拖放接收 (DND) 及 Delete 鍵快速刪除項目
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self._handle_drop)
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        
        # 空白清單引導 Label
        self.lbl_empty_tip = ctk.CTkLabel(
            self.tree, 
            text="📥 拖曳圖片檔案至此處，或點擊選擇圖片新增",
            font=(SYSTEM_FONT, 11 + FONT_OFFSET),
            text_color=("#9CA3AF", "#6B7280"),
            fg_color="transparent"
        )
        self.lbl_empty_tip.bind("<Button-1>", lambda e: self.tree.focus_set())
        
        # 註冊拖放給 Label
        self.lbl_empty_tip.drop_target_register(DND_FILES)
        self.lbl_empty_tip.dnd_bind("<<Drop>>", self._handle_drop)
        self.lbl_empty_tip.place(relx=0.5, rely=0.5, anchor="center")
        
        # 右鍵快顯選單
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="🔍 預覽壓縮與縮放效果", command=self._preview_compress_effect)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📁 在檔案總管中顯示", command=self._show_in_explorer)
        self.context_menu.add_command(label="❌ 移除此項目", command=self._remove_selected)
        
        self.tree.bind("<Button-3>", self._show_context_menu)
        
        # 下方檔案操作按鈕
        btn_frame = ctk.CTkFrame(left_pane, fg_color="transparent")
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.btn_add = ctk.CTkButton(
            btn_frame, text="➕ 新增圖片", 
            command=self._add_files, 
            width=100, height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.btn_add.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_remove = ctk.CTkButton(
            btn_frame, text="➖ 移除所選", 
            command=self._remove_selected, 
            width=100, height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.btn_remove.pack(side=tk.LEFT, padx=10)
        
        self.btn_clear = ctk.CTkButton(
            btn_frame, text="🧹 清空清單", 
            command=self._clear_all, 
            width=100, height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.btn_clear.pack(side=tk.LEFT, padx=10)
        
        # 右側：執行按鈕與參數設定
        right_pane_outer = ctk.CTkFrame(
            t6_main, width=340, 
            fg_color=("#FFFFFF", "#1F2937"),
            border_width=1, border_color=("#E5E7EB", "#374151")
        )
        right_pane_outer.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        right_pane_outer.pack_propagate(False)
        
        right_pane = ctk.CTkFrame(right_pane_outer, fg_color="transparent")
        right_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 標題
        t6_setting_lbl = ctk.CTkLabel(right_pane, text="⚙️ 執行與壓縮設定", font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        t6_setting_lbl.pack(anchor="w", pady=(0, 10))
        
        # 執行區域
        exec_frame = ctk.CTkFrame(right_pane, fg_color="transparent")
        exec_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.auto_open_var = tk.BooleanVar(value=False)
        self.check_open = ctk.CTkCheckBox(
            exec_frame, text="完成後開啟資料夾", 
            variable=self.auto_open_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_open.pack(anchor="w", pady=(0, 8))
        
        self.progress = ctk.CTkProgressBar(exec_frame)
        self.progress.set(0)
        self.progress.pack(fill=tk.X, pady=(0, 8))
        
        self.btn_run_frame = ctk.CTkFrame(exec_frame, fg_color="transparent", height=44)
        self.btn_run_frame.pack(fill=tk.X)
        self.btn_run_frame.pack_propagate(False)
        
        self.btn_run = ctk.CTkButton(
            self.btn_run_frame, text="🚀 開始壓縮", 
            fg_color=("#2563EB", "#3B82F6"),
            text_color="white",
            hover_color=("#1D4ED8", "#2563EB"),
            font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
            command=self._start_compress,
            height=44
        )
        self.btn_run.pack(fill=tk.BOTH, expand=True)
        
        self.btn_cancel = ctk.CTkButton(
            self.btn_run_frame, text="⛔ 終止作業", 
            fg_color=("#DC2626", "#EF4444"),
            text_color="white",
            hover_color=("#B91C1C", "#DC2626"),
            font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
            command=self._cancel_compress,
            height=44
        )
        
        ctk.CTkFrame(right_pane, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 滾動設定區
        t6_scroll_settings = ctk.CTkScrollableFrame(right_pane, fg_color="transparent", label_text="")
        t6_scroll_settings.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 1. 壓縮品質 Quality
        self.quality_frame = ctk.CTkFrame(t6_scroll_settings, fg_color="transparent")
        self.quality_frame.pack(fill=tk.X, pady=(2, 6))
        self.lbl_quality = ctk.CTkLabel(self.quality_frame, text="壓縮品質: 70%", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        self.lbl_quality.pack(side=tk.LEFT)
        self.slider_quality = ctk.CTkSlider(
            self.quality_frame, from_=10, to=95, 
            number_of_steps=85, height=16,
            command=self._update_quality_lbl
        )
        self.slider_quality.set(70)
        self.slider_quality.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        ctk.CTkFrame(t6_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=6)
        
        # 2. 縮放尺寸設定
        lbl_resize = ctk.CTkLabel(t6_scroll_settings, text="圖片縮放設定:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_resize.pack(anchor="w", pady=(2, 2))
        
        self.combo_resize_mode = ctk.CTkOptionMenu(
            t6_scroll_settings, 
            values=["不縮放 (保持原尺寸)", "設定百分比 (%)", "設定寬度 (高度自適應)", "設定高度 (寬度自適應)", "設定固定寬高 (像素)"],
            command=self._toggle_resize_fields,
            height=30
        )
        self.combo_resize_mode.pack(fill=tk.X, pady=2)
        self.combo_resize_mode.set("不縮放 (保持原尺寸)")
        
        # 縮放欄位區
        self.resize_fields_frame = ctk.CTkFrame(t6_scroll_settings, fg_color="transparent")
        
        # 百分比輸入
        self.entry_percent = ctk.CTkEntry(self.resize_fields_frame, placeholder_text="百分比 (例: 50)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_percent.pack(fill=tk.X, pady=1)
        self.entry_percent.insert(0, "50")
        
        # 寬度輸入
        self.entry_width = ctk.CTkEntry(self.resize_fields_frame, placeholder_text="寬度 px (例: 1280)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_width.pack(fill=tk.X, pady=1)
        
        # 高度輸入
        self.entry_height = ctk.CTkEntry(self.resize_fields_frame, placeholder_text="高度 px (例: 720)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_height.pack(fill=tk.X, pady=1)
        
        ctk.CTkFrame(t6_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=6)
        
        # 3. 旋轉與方向設定
        lbl_rot = ctk.CTkLabel(t6_scroll_settings, text="圖片方向與旋轉:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_rot.pack(anchor="w", pady=(2, 2))
        
        self.rotate_var = tk.BooleanVar(value=False)
        self.check_rotate = ctk.CTkCheckBox(
            t6_scroll_settings, text="啟用圖片旋轉", 
            variable=self.rotate_var,
            command=self._toggle_rotate_field,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_rotate.pack(anchor="w", pady=4)
        
        self.combo_rotate_angle = ctk.CTkOptionMenu(
            t6_scroll_settings, 
            values=["90", "180", "270", "45", "135", "315"],
            height=30
        )
        self.combo_rotate_angle.pack(fill=tk.X, pady=2)
        self.combo_rotate_angle.set("90")
        
        self.exif_transpose_var = tk.BooleanVar(value=True)
        self.check_exif_transpose = ctk.CTkCheckBox(
            t6_scroll_settings, text="自動修正手機相片方向 (Exif)", 
            variable=self.exif_transpose_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_exif_transpose.pack(anchor="w", pady=4)
        
        ctk.CTkFrame(t6_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=6)
        
        # 4. 格式與隱私設定
        lbl_fmt_sec = ctk.CTkLabel(t6_scroll_settings, text="輸出格式與隱私:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_fmt_sec.pack(anchor="w", pady=(2, 2))
        
        fmt_f = ctk.CTkFrame(t6_scroll_settings, fg_color="transparent")
        fmt_f.pack(fill=tk.X, pady=2)
        lbl_fmt = ctk.CTkLabel(fmt_f, text="輸出格式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_fmt.pack(side=tk.LEFT)
        self.combo_format = ctk.CTkOptionMenu(
            fmt_f, values=["保持原格式", "JPEG", "PNG", "WebP"], 
            height=30, width=120
        )
        self.combo_format.pack(side=tk.RIGHT)
        self.combo_format.set("保持原格式")
        
        self.erase_exif_var = tk.BooleanVar(value=True)
        self.check_erase_exif = ctk.CTkCheckBox(
            t6_scroll_settings, text="抹除 EXIF 隱私資訊 (如 GPS 等)", 
            variable=self.erase_exif_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_erase_exif.pack(anchor="w", pady=4)
        
        self.png_quantize_var = tk.BooleanVar(value=True)
        self.check_png_quantize = ctk.CTkCheckBox(
            t6_scroll_settings, text="PNG 啟用色彩量化 (大幅縮減 PNG 體積)", 
            variable=self.png_quantize_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_png_quantize.pack(anchor="w", pady=4)
        
        # 檔名後綴
        suffix_f = ctk.CTkFrame(t6_scroll_settings, fg_color="transparent")
        suffix_f.pack(fill=tk.X, pady=4)
        lbl_suffix = ctk.CTkLabel(suffix_f, text="輸出檔名後綴:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_suffix.pack(side=tk.LEFT)
        self.entry_suffix = ctk.CTkEntry(suffix_f, placeholder_text="例: _compressed", height=30, width=120, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_suffix.pack(side=tk.RIGHT)
        self.entry_suffix.insert(0, "_compressed")
        
        ctk.CTkFrame(t6_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=6)
        
        # 5. 批次 EXIF 寫入設定
        self.exif_write_var = tk.BooleanVar(value=False)
        self.check_exif_write = ctk.CTkCheckBox(
            t6_scroll_settings, text="啟用批次自訂 EXIF 資訊", 
            variable=self.exif_write_var,
            command=self._toggle_exif_fields,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold")
        )
        self.check_exif_write.pack(anchor="w", pady=4)
        
        self.exif_fields_frame = ctk.CTkFrame(t6_scroll_settings, fg_color="transparent")
        
        self.entry_exif_artist = ctk.CTkEntry(self.exif_fields_frame, placeholder_text="作者 (Artist)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_exif_artist.pack(fill=tk.X, pady=1)
        
        self.entry_exif_copyright = ctk.CTkEntry(self.exif_fields_frame, placeholder_text="版權聲明 (Copyright)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_exif_copyright.pack(fill=tk.X, pady=1)
        
        self.entry_exif_desc = ctk.CTkEntry(self.exif_fields_frame, placeholder_text="相片描述 (Description)", height=30, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_exif_desc.pack(fill=tk.X, pady=1)
        
        ctk.CTkFrame(t6_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=6)
        
        # 6. 輸出目錄模式
        lbl_out = ctk.CTkLabel(t6_scroll_settings, text="輸出目錄模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_out.pack(anchor="w", pady=(2, 1))
        
        self.out_mode_var = tk.StringVar(value="folder")
        out_radio_f = ctk.CTkFrame(t6_scroll_settings, fg_color="transparent")
        out_radio_f.pack(fill=tk.X, pady=(0, 10))
        
        self.out_radio1 = ctk.CTkRadioButton(
            out_radio_f, text="建立獨立子資料夾", 
            variable=self.out_mode_var, value="folder",
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.out_radio1.pack(side=tk.LEFT, padx=(0, 10))
        
        self.out_radio2 = ctk.CTkRadioButton(
            out_radio_f, text="同層輸出", 
            variable=self.out_mode_var, value="same",
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.out_radio2.pack(side=tk.LEFT)
        
        ctk.CTkFrame(right_pane, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # Log 輸出框
        self.show_log_var = tk.BooleanVar(value=False)
        self.check_show_log = ctk.CTkCheckBox(
            right_pane, text="顯示詳細作業日誌", 
            variable=self.show_log_var,
            command=self._toggle_log_view,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_show_log.pack(anchor="w", pady=(5, 5))
        
        self.log_text = ctk.CTkTextbox(
            right_pane, font=("Consolas", 9 + FONT_OFFSET), 
            fg_color=("#F9FAFB", "#1A1A1A"),
            text_color=("#374151", "#E5E7EB"),
            border_width=1, border_color=("#E5E7EB", "#374151"),
            height=100
        )
        
        self._toggle_resize_fields()
        self._toggle_rotate_field()
        self._toggle_exif_fields()

    def _update_quality_lbl(self, val):
        self.lbl_quality.configure(text=f"壓縮品質: {int(val)}%")

    def _toggle_resize_fields(self, val=None):
        mode = self.combo_resize_mode.get()
        
        # 先全部隱藏/清空
        self.entry_percent.pack_forget()
        self.entry_width.pack_forget()
        self.entry_height.pack_forget()
        
        if mode == "不縮放 (保持原尺寸)":
            self.resize_fields_frame.pack_forget()
        else:
            self.resize_fields_frame.pack(fill=tk.X, pady=4, after=self.combo_resize_mode)
            if mode == "設定百分比 (%)":
                self.entry_percent.pack(fill=tk.X, pady=1)
            elif mode == "設定寬度 (高度自適應)":
                self.entry_width.pack(fill=tk.X, pady=1)
            elif mode == "設定高度 (寬度自適應)":
                self.entry_height.pack(fill=tk.X, pady=1)
            elif mode == "設定固定寬高 (像素)":
                self.entry_width.pack(fill=tk.X, pady=1)
                self.entry_height.pack(fill=tk.X, pady=1)

    def _toggle_rotate_field(self):
        state = "normal" if self.rotate_var.get() else "disabled"
        self.combo_rotate_angle.configure(state=state)

    def _toggle_exif_fields(self):
        if self.exif_write_var.get():
            self.exif_fields_frame.pack(fill=tk.X, pady=2, after=self.check_exif_write)
            self.entry_exif_artist.configure(state="normal")
            self.entry_exif_copyright.configure(state="normal")
            self.entry_exif_desc.configure(state="normal")
        else:
            self.exif_fields_frame.pack_forget()

    def _log(self, msg):
        self.app.queue.put(("t6_log", msg))

    def update_log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def _toggle_log_view(self):
        if self.show_log_var.get():
            self.log_text.pack(fill=tk.X, expand=False, pady=(0, 5))
        else:
            self.log_text.pack_forget()

    def _add_files(self):
        if self.is_converting:
            return
        files = filedialog.askopenfilenames(
            title="選擇待處理圖片", 
            filetypes=[("圖片檔案", "*.jpg *.jpeg *.png *.bmp *.tiff *.webp")]
        )
        if files:
            self._process_incoming_files(files)

    def _handle_drop(self, event):
        if self.is_converting:
            return
        files = parse_dropped_files(event.data)
        self._process_incoming_files(files)

    def _process_incoming_files(self, files):
        added = False
        valid_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp')
        for f in files:
            if not f.lower().endswith(valid_exts):
                continue
            if f in self.selected_files:
                continue
            self.selected_files.append(f)
            added = True
            
        if added:
            self._update_tree_content()

    def _remove_selected(self):
        if self.is_converting:
            return
        selected = self.tree.selection()
        if not selected:
            return
        
        indices = []
        for item in selected:
            values = self.tree.item(item, "values")
            idx = int(values[0]) - 1
            indices.append(idx)
            
        indices.sort(reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.selected_files):
                self.selected_files.pop(idx)
                
        self._update_tree_content()

    def _show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _show_in_explorer(self):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        try:
            idx = int(values[0]) - 1
            path = self.selected_files[idx]
            import subprocess
            subprocess.run(f'explorer /select,"{os.path.normpath(path)}"')
        except Exception as e:
            pass

    def _preview_compress_effect(self):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        try:
            idx = int(values[0]) - 1
            img_path = self.selected_files[idx]
        except Exception:
            return
            
        def parse_val(text, default):
            v = text.strip()
            if not v:
                return default
            try:
                return int(v)
            except:
                return default

        settings = {
            "quality": int(self.slider_quality.get()),
            "resize_mode": self.combo_resize_mode.get(),
            "percent": parse_val(self.entry_percent.get(), 100),
            "width": parse_val(self.entry_width.get(), 0),
            "height": parse_val(self.entry_height.get(), 0),
            "rotate": self.rotate_var.get(),
            "rotate_angle": int(self.combo_rotate_angle.get()) if self.rotate_var.get() else 0,
            "exif_transpose": self.exif_transpose_var.get(),
            "format": self.combo_format.get(),
            "erase_exif": self.erase_exif_var.get(),
            "png_quantize": self.png_quantize_var.get(),
        }
        
        from components.dialogs import ModernCompressPreviewDialog
        ModernCompressPreviewDialog(self.winfo_toplevel(), img_path, settings)

    def _clear_all(self):
        if self.is_converting or not self.selected_files:
            return
        if messagebox.askyesno("確認", "是否確定清空所選的圖片檔案？"):
            self.selected_files.clear()
            self._update_tree_content()

    def _update_tree_content(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for idx, f in enumerate(self.selected_files):
            fname = os.path.basename(f)
            res_str = "讀取中..."
            size_str = "讀取中..."
            
            try:
                # 僅讀取 Header，避免解碼整張大圖以維持載入速度
                with Image.open(f) as img:
                    res_str = f"{img.width} x {img.height}"
                orig_size = os.path.getsize(f)
                size_str = f"{orig_size / (1024*1024):.2f} MB" if orig_size > 1024*1024 else f"{orig_size / 1024:.1f} KB"
            except:
                res_str = "無法讀取"
                size_str = "未知"
                
            self.tree.insert("", tk.END, values=(idx + 1, fname, res_str, size_str, "待處理"))
            
        if not self.selected_files:
            self.lbl_empty_tip.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.lbl_empty_tip.place_forget()

    def _start_compress(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "請先選擇圖片檔案。")
            return
            
        mode = self.combo_resize_mode.get()
        percent_str = self.entry_percent.get().strip()
        width_str = self.entry_width.get().strip()
        height_str = self.entry_height.get().strip()
        
        # 尺寸參數校驗
        pct = 100
        w = 0
        h = 0
        
        try:
            if mode == "設定百分比 (%)":
                pct = int(percent_str)
                if pct <= 0: raise ValueError()
            elif mode == "設定寬度 (高度自適應)":
                w = int(width_str)
                if w <= 0: raise ValueError()
            elif mode == "設定高度 (寬度自適應)":
                h = int(height_str)
                if h <= 0: raise ValueError()
            elif mode == "設定固定寬高 (像素)":
                w = int(width_str)
                h = int(height_str)
                if w <= 0 or h <= 0: raise ValueError()
        except:
            messagebox.showerror("參數錯誤", "縮放大小設定值必須為正整數！")
            return
            
        settings = {
            "quality": int(self.slider_quality.get()),
            "resize_mode": mode,
            "percent": pct,
            "width": w,
            "height": h,
            "rotate": self.rotate_var.get(),
            "rotate_angle": int(self.combo_rotate_angle.get()) if self.rotate_var.get() else 0,
            "exif_transpose": self.exif_transpose_var.get(),
            "format": self.combo_format.get(),
            "erase_exif": self.erase_exif_var.get(),
            "png_quantize": self.png_quantize_var.get(),
            "suffix": self.entry_suffix.get().strip(),
            "exif_write": self.exif_write_var.get(),
            "exif_artist": self.entry_exif_artist.get().strip(),
            "exif_copyright": self.entry_exif_copyright.get().strip(),
            "exif_desc": self.entry_exif_desc.get().strip(),
            "out_mode": self.out_mode_var.get(),
            "open": self.auto_open_var.get()
        }
        
        self.is_converting = True
        self.btn_run.pack_forget()
        self.btn_cancel.pack(fill=tk.BOTH, expand=True)
        self._toggle_ui_state("disabled")
        
        self.progress.set(0)
        self.log_text.delete("1.0", tk.END)
        self.update_log("===============================")
        self.update_log("🚀 圖片批次壓縮與處理作業開始...")
        self.stop_event.clear()
        
        threading.Thread(target=self._compress_worker, args=(settings,), daemon=True).start()

    def _cancel_compress(self):
        if messagebox.askyesno("取消", "確定要終止目前的圖片壓縮作業？"):
            self.stop_event.set()
            self.btn_cancel.configure(state="disabled")
            self._log("🛑 正在停止作業，請稍後...")

    def _toggle_ui_state(self, state):
        mode = "normal" if state == "normal" else "disabled"
        self.btn_add.configure(state=mode)
        self.btn_remove.configure(state=mode)
        self.btn_clear.configure(state=mode)
        self.combo_resize_mode.configure(state=mode)
        self.check_rotate.configure(state=mode)
        self.check_exif_transpose.configure(state=mode)
        self.combo_format.configure(state=mode)
        self.check_erase_exif.configure(state=mode)
        self.check_png_quantize.configure(state=mode)
        self.entry_suffix.configure(state=mode)
        self.check_exif_write.configure(state=mode)
        self.check_open.configure(state=mode)
        self.out_radio1.configure(state=mode)
        self.out_radio2.configure(state=mode)
        self.check_show_log.configure(state=mode)
        
        if mode == "normal":
            self._toggle_resize_fields()
            self._toggle_rotate_field()
            self._toggle_exif_fields()
        else:
            self.entry_percent.configure(state="disabled")
            self.entry_width.configure(state="disabled")
            self.entry_height.configure(state="disabled")
            self.combo_rotate_angle.configure(state="disabled")
            self.entry_exif_artist.configure(state="disabled")
            self.entry_exif_copyright.configure(state="disabled")
            self.entry_exif_desc.configure(state="disabled")

    def _compress_worker(self, settings):
        """圖片壓縮背景 Worker (多執行緒並行版)"""
        import concurrent.futures
        
        try:
            total_tasks = len(self.selected_files)
            final_out_dir = None
            processed_count = 0
            count_lock = threading.Lock()
            
            # 決定並行數量，一般 4 個執行緒能獲得極佳效能提升
            max_workers = min(4, os.cpu_count() or 1)
            self._log(f"💻 啟動多核心並行處理，並行執行緒數：{max_workers}")
            
            def process_single_image(task_info):
                nonlocal processed_count, final_out_dir
                idx, img_path = task_info
                
                if self.stop_event.is_set():
                    return
                    
                base_name = os.path.splitext(os.path.basename(img_path))[0]
                orig_ext = os.path.splitext(img_path)[1].lower()
                out_dir = os.path.dirname(img_path)
                
                # 計算輸出目錄
                if settings["out_mode"] == "folder":
                    out_target_dir = os.path.join(out_dir, "compressed_images")
                    os.makedirs(out_target_dir, exist_ok=True)
                else:
                    out_target_dir = out_dir
                    
                final_out_dir = out_target_dir
                
                # 計算輸出格式與副檔名
                fmt = settings["format"]
                if fmt == "保持原格式":
                    ext = orig_ext
                    save_format = "PNG" if orig_ext == ".png" else ("WEBP" if orig_ext == ".webp" else "JPEG")
                else:
                    ext = f".{fmt.lower()}"
                    save_format = fmt
                    
                out_fname = f"{base_name}{settings['suffix']}{ext}"
                out_path = os.path.join(out_target_dir, out_fname)
                out_path = unique_filename(out_path)
                
                self._log(f"⚡ 正在處理：{os.path.basename(img_path)}")
                
                try:
                    orig_size = os.path.getsize(img_path)
                    
                    with Image.open(img_path) as img:
                        # 1. 手機拍攝方向修正
                        if settings["exif_transpose"]:
                            img = ImageOps.exif_transpose(img)
                            
                        # 2. 獲取原 EXIF 物件
                        exif = img.getexif()
                        
                        # 3. 尺寸縮放計算
                        mode = settings["resize_mode"]
                        orig_w, orig_h = img.size
                        new_w, new_h = orig_w, orig_h
                        
                        if mode == "設定百分比 (%)":
                            new_w = int(orig_w * settings["percent"] / 100.0)
                            new_h = int(orig_h * settings["percent"] / 100.0)
                        elif mode == "設定寬度 (高度自適應)":
                            new_w = settings["width"]
                            new_h = int(orig_h * (new_w / orig_w))
                        elif mode == "設定高度 (寬度自適應)":
                            new_h = settings["height"]
                            new_w = int(orig_w * (new_h / orig_h))
                        elif mode == "設定固定寬高 (像素)":
                            new_w = settings["width"]
                            new_h = settings["height"]
                            
                        new_w = max(1, new_w)
                        new_h = max(1, new_h)
                        
                        # 執行縮放
                        if (new_w, new_h) != (orig_w, orig_h):
                            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                            self._log(f"  ➜ {os.path.basename(img_path)}：調整 {orig_w}x{orig_h} ➜ {new_w}x{new_h}")
                            
                        # 4. 圖片旋轉
                        if settings["rotate"] and settings["rotate_angle"] != 0:
                            img = img.rotate(settings["rotate_angle"], expand=True)
                            self._log(f"  ➜ {os.path.basename(img_path)}：旋轉 {settings['rotate_angle']}°")
                            
                        # 5. EXIF 抹除與寫入處理
                        if settings["erase_exif"]:
                            exif.clear()
                            
                        if settings["exif_write"]:
                            if settings["exif_artist"]:
                                exif[315] = settings["exif_artist"]
                            if settings["exif_copyright"]:
                                exif[33432] = settings["exif_copyright"]
                            if settings["exif_desc"]:
                                exif[270] = settings["exif_desc"]
                            exif[305] = "PDF & Image Toolkit"
                            
                        # 6. 壓縮儲存
                        save_kwargs = {"optimize": True}
                        if save_format == "PNG":
                            save_kwargs["compress_level"] = 9
                        if save_format in ["JPEG", "WEBP"]:
                            save_kwargs["quality"] = settings["quality"]
                        if len(exif) > 0:
                            try:
                                save_kwargs["exif"] = exif.tobytes()
                            except Exception as exif_err:
                                self._log(f"  ⚠️ {os.path.basename(img_path)}：EXIF 寫入失敗 (格式不支援)：{str(exif_err)}")
                            
                        if save_format == "JPEG" and img.mode in ("RGBA", "P"):
                            img_to_save = img.convert("RGB")
                        elif save_format == "PNG" and settings["png_quantize"]:
                            img_to_save = img.quantize(colors=256)
                        else:
                            img_to_save = img
                            
                        img_to_save.save(out_path, save_format, **save_kwargs)
                        
                    new_size = os.path.getsize(out_path)
                    saved = orig_size - new_size
                    pct = (saved / orig_size) * 100 if orig_size > 0 else 0
                    
                    if saved > 0:
                        self._log(f"  ➜ {os.path.basename(img_path)}：體積減少 {pct:.1f}% ({saved / (1024*1024):.2f} MB)")
                    else:
                        self._log(f"  ➜ {os.path.basename(img_path)}：處理完成 (體積無變化)")
                        
                except Exception as e:
                    self._log(f"❌ 圖片「{os.path.basename(img_path)}」處理失敗：{str(e)}")
                    
                with count_lock:
                    processed_count += 1
                    self.app.queue.put(("t6_progress", processed_count / total_tasks))
            
            # 提交並等待執行緒池任務
            tasks = list(enumerate(self.selected_files))
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(process_single_image, t) for t in tasks]
                concurrent.futures.wait(futures)
                
            if self.stop_event.is_set():
                raise InterruptedError()
                
            self.app.queue.put(("t6_done", final_out_dir))
            
        except InterruptedError:
            self.app.queue.put(("t6_cancelled", None))
        except Exception as e:
            self.app.queue.put(("t6_error", str(e)))
