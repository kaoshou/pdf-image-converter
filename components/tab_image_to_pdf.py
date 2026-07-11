import os
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import fitz  # PyMuPDF
from tkinterdnd2 import DND_FILES

from utils.helpers import SYSTEM_FONT, FONT_OFFSET, parse_dropped_files, unique_filename, finalize_and_save_pdf
from components.dialogs import ModernPasswordDialog
from components.pdf_features import PDFFeaturesFrame

# 標準 PDF 頁面尺寸定義 (Points: 1 inch = 72 points)
PAGE_SIZES = {
    "原始大小": None,
    "A3 (297 x 420 mm)": (841.89, 1190.55),
    "A4 (210 x 297 mm)": (595.27, 841.89),
    "A5 (148 x 210 mm)": (419.53, 595.27),
    "A6 (105 x 148 mm)": (297.64, 419.53),
    "B4 (250 x 353 mm)": (708.66, 1000.63),
    "B5 (176 x 250 mm)": (498.90, 708.66),
    "Letter (8.5 x 11\")": (612.0, 792.0),
    "Legal (8.5 x 14\")": (612.0, 1008.0),
    "Tabloid (11 x 17\")": (792.0, 1224.0),
    "4 x 6 吋 (相片)": (288.0, 432.0),
    "5 x 7 吋 (相片)": (360.0, 504.0),
}

class TabImageToPDF(ctk.CTkFrame):
    """圖片與 PDF 合併轉換為新 PDF 的功能分頁"""
    def __init__(self, master, app):
        super().__init__(master, fg_color="transparent")
        self.app = app
        
        # 功能內部變數
        self.file_list = []        # 儲存的項目: {'path': f, 'page': None/idx, 'page_count': n}
        self.passwords = {}        # 儲存載入 PDF 的解鎖密碼
        self.thumbnails = {}       # 快取縮圖: {f"{path}_{page}": PhotoImage}
        self.doc_handles = {}      # 快取 fitz.Document 物件以提升效能
        self.is_converting = False
        
        self.thumbnail_size = 50   # 預設縮圖大小為 50 像素
        self.last_col0_width = 70  # 預設第一欄 (#0) 寬度為 70 像素
        self._debounce_id = None   # 用於防震處理
        
        # 啟動異步縮圖載入執行緒
        self.thumb_queue = queue.Queue()
        self.thumb_thread_running = True
        self.thumb_worker = threading.Thread(target=self._thumbnail_worker, daemon=True)
        self.thumb_worker.start()
        
        # 初始化 UI 介面
        self._build_ui()

    def _build_ui(self):
        # 左右分割配置
        t1_main = ctk.CTkFrame(self, fg_color="transparent")
        t1_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：檔案清單與按鈕
        left_pane = ctk.CTkFrame(t1_main, fg_color="transparent")
        
        # 頂部控制列
        t1_ctrl = ctk.CTkFrame(left_pane, fg_color="transparent", height=40)
        t1_ctrl.pack(fill=tk.X, pady=(0, 8))
        
        self.btn_add = ctk.CTkButton(t1_ctrl, text=" ＋ 選擇檔案... ", 
                                         fg_color=("#2563EB", "#3B82F6"),
                                         text_color="white",
                                         hover_color=("#1D4ED8", "#2563EB"),
                                         font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                         command=self._add_files,
                                         height=32)
        self.btn_add.pack(side=tk.LEFT)
        

        
        self.lbl_count = ctk.CTkLabel(t1_ctrl, text="已選擇: 0 個項目",
                                         font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                         text_color=("#2563EB", "#3B82F6"))
        self.lbl_count.pack(side=tk.RIGHT)
        
        # 列表框
        list_container = ctk.CTkFrame(left_pane, fg_color="transparent")
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # Treeview 與其捲軸
        tree_frame = ctk.CTkFrame(list_container, fg_color="transparent")
        tree_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("Index", "Type", "Name"), show='headings', selectmode='extended', style="T1.Treeview")
        self.tree.heading("Index", text="順序/頁碼")
        self.tree.heading("Type", text="類型")
        self.tree.heading("Name", text="檔案名稱 / 頁面")
        
        self.tree.column("#0", width=70, anchor="center", stretch=False)  # 縮圖欄
        self.tree.column("Index", width=90, anchor="center", stretch=False)
        self.tree.column("Type", width=90, anchor="center", stretch=False)
        self.tree.column("Name", width=350, anchor="w")
        self.tree.configure(show="tree headings")
        self.tree.heading("#0", text="預覽")
        
        t1_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=t1_scroll.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        t1_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview 拖放註冊
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self._handle_drop)
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Configure>", self._on_tree_configure)
        
        # 拖曳排序事件綁定
        self._dragged_item = None
        self._is_dragging = False
        self.tree.bind("<ButtonPress-1>", self._on_drag_start, add="+")
        self.tree.bind("<B1-Motion>", self._on_drag_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_drag_drop, add="+")
        
        # 空白清單引導 Label
        self.lbl_empty_tip = ctk.CTkLabel(
            self.tree, 
            text="📥 拖曳 PDF 或圖片檔案至此處，或點擊選擇檔案新增",
            font=(SYSTEM_FONT, 11 + FONT_OFFSET),
            text_color=("#9CA3AF", "#6B7280"),
            fg_color="transparent"
        )
        self.lbl_empty_tip.bind("<Button-1>", lambda e: self.tree.focus_set())
        
        # 註冊拖放給 Label，保證拖曳到 Label 上也有效
        self.lbl_empty_tip.drop_target_register(DND_FILES)
        self.lbl_empty_tip.dnd_bind("<<Drop>>", self._handle_drop)
        self.lbl_empty_tip.place(relx=0.5, rely=0.5, anchor="center")
        
        # 右鍵快顯選單
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="📁 在檔案總管中顯示", command=self._show_in_explorer)
        self.context_menu.add_command(label="❌ 移除此項目", command=self._remove_selected)
        
        self.tree.bind("<Button-3>", self._show_context_menu)
        
        # 橫向控制按鈕列 (兩行自適應設計)
        ctrl_bar1 = ctk.CTkFrame(list_container, fg_color="transparent", height=32)
        ctrl_bar1.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))
        
        ctrl_bar2 = ctk.CTkFrame(list_container, fg_color="transparent", height=32)
        ctrl_bar2.pack(side=tk.TOP, fill=tk.X, pady=(6, 0))
        
        btn_opt = {"height": 30, "font": (SYSTEM_FONT, 9 + FONT_OFFSET)}
        
        # 第一行：清單操作與排序
        self.btn_expand = ctk.CTkButton(ctrl_bar1, text="📂 展開 PDF", 
                                            fg_color=("#EFF6FF", "#1E293B"),
                                            text_color=("#2563EB", "#3B82F6"),
                                            hover_color=("#DBEAFE", "#334155"),
                                            command=self._expand_selected_pdf, **btn_opt)
        self.btn_expand.pack(side=tk.LEFT, padx=(0, 6), expand=True, fill=tk.X)
        
        self.btn_up = ctk.CTkButton(ctrl_bar1, text="▲ 上移", 
                                        fg_color=("#F3F4F6", "#374151"),
                                        text_color=("#374151", "#F9FAFB"),
                                        hover_color=("#E5E7EB", "#4B5563"),
                                        command=self._move_up, **btn_opt)
        self.btn_up.pack(side=tk.LEFT, padx=(0, 6), expand=True, fill=tk.X)
        
        self.btn_down = ctk.CTkButton(ctrl_bar1, text="▼ 下移", 
                                          fg_color=("#F3F4F6", "#374151"),
                                          text_color=("#374151", "#F9FAFB"),
                                          hover_color=("#E5E7EB", "#4B5563"),
                                          command=self._move_down, **btn_opt)
        self.btn_down.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        # 第二行：自動排序與移除清空
        self.btn_sort_asc = ctk.CTkButton(ctrl_bar2, text="A-Z 排序", 
                                              fg_color=("#F3F4F6", "#374151"),
                                              text_color=("#374151", "#F9FAFB"),
                                              hover_color=("#E5E7EB", "#4B5563"),
                                              command=lambda: self._sort_files(False), **btn_opt)
        self.btn_sort_asc.pack(side=tk.LEFT, padx=(0, 6), expand=True, fill=tk.X)
        
        self.btn_sort_desc = ctk.CTkButton(ctrl_bar2, text="Z-A 排序", 
                                               fg_color=("#F3F4F6", "#374151"),
                                               text_color=("#374151", "#F9FAFB"),
                                               hover_color=("#E5E7EB", "#4B5563"),
                                               command=lambda: self._sort_files(True), **btn_opt)
        self.btn_sort_desc.pack(side=tk.LEFT, padx=(0, 6), expand=True, fill=tk.X)
        
        self.btn_remove = ctk.CTkButton(ctrl_bar2, text="✕ 移除選取", 
                                            fg_color=("#FEF2F2", "#450A0A"),
                                            text_color=("#DC2626", "#F87171"),
                                            hover_color=("#FEE2E2", "#7F1D1D"),
                                            command=self._remove_selected, **btn_opt)
        self.btn_remove.pack(side=tk.LEFT, padx=(0, 6), expand=True, fill=tk.X)
        
        self.btn_clear = ctk.CTkButton(ctrl_bar2, text="🗑 全部清空", 
                                           fg_color=("#FEF2F2", "#450A0A"),
                                           text_color=("#DC2626", "#F87171"),
                                           hover_color=("#FEE2E2", "#7F1D1D"),
                                           command=self._clear_all, **btn_opt)
        self.btn_clear.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # 右側：執行按鈕與參數設定
        right_pane_outer = ctk.CTkFrame(t1_main, width=340, fg_color=("#FFFFFF", "#1F2937"), 
                                        border_width=1, border_color=("#E5E7EB", "#374151"))
        right_pane_outer.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        right_pane_outer.pack_propagate(False)
        
        # 讓左側面板在右側面板 pack 之後才 pack，以防其 expand=True 擠壓右側空間
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        right_pane = ctk.CTkFrame(right_pane_outer, fg_color="transparent")
        right_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 標題
        t1_setting_lbl = ctk.CTkLabel(right_pane, text="⚙️ 執行與參數設定", 
                                      font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        t1_setting_lbl.pack(anchor="w", pady=(0, 8))
        
        # 執行區域 (置頂放置)
        exec_frame = ctk.CTkFrame(right_pane, fg_color="transparent")
        exec_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.auto_open_var = tk.BooleanVar(value=False)
        self.check_open = ctk.CTkCheckBox(exec_frame, text="完成後開啟資料夾", 
                                             variable=self.auto_open_var,
                                             font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_open.pack(anchor="w", pady=(0, 8))
        
        self.progress = ctk.CTkProgressBar(exec_frame)
        self.progress.set(0)
        self.progress.pack(fill=tk.X, pady=(0, 5))
        
        self.lbl_status = ctk.CTkLabel(exec_frame, text="等待作業中...", 
                                           font=(SYSTEM_FONT, 9 + FONT_OFFSET),
                                           text_color="#9CA3AF")
        self.lbl_status.pack(anchor="w", pady=(0, 8))
        
        self.btn_run = ctk.CTkButton(exec_frame, text="🚀 開始產生 PDF", 
                                         fg_color=("#059669", "#10B981"),
                                         text_color="white",
                                         hover_color=("#047857", "#059669"),
                                         font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
                                         height=44,
                                         command=self._start_conversion)
        self.btn_run.pack(fill=tk.X)
        
        # 分割線
        ctk.CTkFrame(right_pane, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 滾動設定區
        scroll_settings = ctk.CTkScrollableFrame(right_pane, fg_color="transparent", label_text="")
        scroll_settings.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 1. 頁面尺寸
        lbl_size = ctk.CTkLabel(scroll_settings, text="頁面尺寸:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_size.pack(anchor="w", pady=(5, 1))
        self.combo_size = ctk.CTkOptionMenu(scroll_settings, values=list(PAGE_SIZES.keys()), height=30)
        self.combo_size.pack(fill=tk.X, pady=(0, 8))
        self.combo_size.set("原始大小")
        
        # 2. 方向
        lbl_orient = ctk.CTkLabel(scroll_settings, text="方向:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_orient.pack(anchor="w", pady=(5, 1))
        self.combo_orient = ctk.CTkOptionMenu(scroll_settings, values=["直式", "橫式"], height=30)
        self.combo_orient.pack(fill=tk.X, pady=(0, 8))
        self.combo_orient.set("直式")
        
        # 3. 圖片縮放
        lbl_scale = ctk.CTkLabel(scroll_settings, text="圖片縮放模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_scale.pack(anchor="w", pady=(5, 1))
        self.combo_scale = ctk.CTkOptionMenu(scroll_settings, values=["自動填滿", "保持原尺寸"], height=30)
        self.combo_scale.pack(fill=tk.X, pady=(0, 8))
        self.combo_scale.set("自動填滿")
        
        ctk.CTkFrame(scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 4. 圖片壓縮
        self.compress_var = tk.BooleanVar(value=False)
        self.check_compress = ctk.CTkCheckBox(scroll_settings, text="啟用圖片壓縮", 
                                                 variable=self.compress_var,
                                                 command=self._toggle_compress,
                                                 font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_compress.pack(anchor="w", pady=4)
        
        self.slider_frame = ctk.CTkFrame(scroll_settings, fg_color="transparent")
        self.slider_frame.pack(fill=tk.X, pady=(2, 8))
        self.lbl_quality = ctk.CTkLabel(self.slider_frame, text="品質: 80%", font=(SYSTEM_FONT, 9 + FONT_OFFSET))
        self.lbl_quality.pack(side=tk.LEFT)
        self.slider_quality = ctk.CTkSlider(self.slider_frame, from_=10, to=100, 
                                               number_of_steps=90, height=16,
                                               command=self._update_quality_lbl)
        self.slider_quality.set(80)
        self.slider_quality.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        # 5. 特殊選項
        self.grayscale_var = tk.BooleanVar(value=False)
        self.check_grayscale = ctk.CTkCheckBox(scroll_settings, text="黑白模式", variable=self.grayscale_var, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_grayscale.pack(anchor="w", pady=4)
        
        self.auto_rotate_var = tk.BooleanVar(value=False)
        self.check_auto_rotate = ctk.CTkCheckBox(scroll_settings, text="自動旋轉圖片頁面", variable=self.auto_rotate_var, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_auto_rotate.pack(anchor="w", pady=4)
        
        self.flatten_var = tk.BooleanVar(value=False)
        self.check_flatten = ctk.CTkCheckBox(scroll_settings, text="PDF 平面化", variable=self.flatten_var, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_flatten.pack(anchor="w", pady=4)
        
        ctk.CTkFrame(scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 6. PDF 高級功能設定 (加密、元資料、浮水印) - 重用元件
        self.pdf_features = PDFFeaturesFrame(scroll_settings)
        self.pdf_features.pack(fill=tk.X, pady=(5, 10))
        
        self._toggle_compress()

    def _toggle_compress(self):
        state = tk.NORMAL if self.compress_var.get() else tk.DISABLED
        self.slider_quality.configure(state=state if state == tk.NORMAL else "disabled")
        self.lbl_quality.configure(text_color=("black", "white") if self.compress_var.get() else "#9CA3AF")

    def _update_quality_lbl(self, val):
        self.lbl_quality.configure(text=f"品質: {int(val)}%")

    def _get_pdf_doc(self, path):
        if path in self.doc_handles:
            return self.doc_handles[path]
        try:
            doc = fitz.open(path)
            if doc.is_encrypted:
                doc.authenticate(self.passwords.get(path, ""))
            self.doc_handles[path] = doc
            return doc
        except:
            return None
    def _on_tree_configure(self, event):
        """當 Treeview 被重新繪製或寬度調整時，監聽並動態更新預覽圖大小"""
        try:
            w = self.tree.column("#0", "width")
        except Exception:
            return
            
        # 如果欄寬變更超過 5 像素，防震處理
        if abs(w - self.last_col0_width) > 5:
            self.last_col0_width = w
            if self._debounce_id is not None:
                self.after_cancel(self._debounce_id)
            self._debounce_id = self.after(350, self._rebuild_thumbnails_with_new_size, w)

    def _rebuild_thumbnails_with_new_size(self, new_width):
        """根據新的寬度，動態調整 rowheight、行高，並清空快取以重新渲染高畫質圖片"""
        self._debounce_id = None
        
        # 行高比欄寬小 8 像素，限制在 40 到 150 像素之間
        target_row_height = max(40, min(new_width - 8, 150))
        self.thumbnail_size = target_row_height - 6
        
        # 1. 更新 Treeview rowheight 樣式
        style = ttk.Style()
        style.configure("T1.Treeview", rowheight=target_row_height)
        
        # 2. 清空圖片快取以強迫重新生成
        self.thumbnails.clear()
        
        # 3. 重新將現有清單中的項目塞入 queue
        for item in self.tree.get_children():
            try:
                idx = self.tree.index(item)
                file_item = self.file_list[idx]
                path = file_item['path']
                p_idx = file_item['page'] if file_item['page'] is not None else 0
                
                self.tree.item(item, image="")
                self.thumb_queue.put((item, path, p_idx))
            except Exception:
                pass

    def _thumbnail_worker(self):
        """背景線程：異步渲染並載入清單項目的縮圖"""
        while self.thumb_thread_running:
            try:
                item_id, path, page_idx = self.thumb_queue.get(timeout=1)
                cache_key = f"{path}_{page_idx}"
                if cache_key not in self.thumbnails:
                    try:
                        target_sz = getattr(self, "thumbnail_size", 50)
                        if path.lower().endswith('.pdf'):
                            with fitz.open(path) as doc:
                                if doc.is_encrypted:
                                    doc.authenticate(self.passwords.get(path, ""))
                                page = doc[page_idx]
                                rect = page.rect
                                zoom = min(target_sz/rect.width, target_sz/rect.height)
                                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                                img_data = pix.tobytes("png")
                        else:
                            with fitz.open(path) as doc:
                                page = doc[0]
                                rect = page.rect
                                zoom = min(target_sz/rect.width, target_sz/rect.height)
                                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                                img_data = pix.tobytes("png")
                                
                        self.after(0, self._update_item_thumbnail, item_id, cache_key, img_data)
                    except:
                        pass
                else:
                    self.after(0, lambda: self._apply_cached_thumbnail(item_id, cache_key))
                self.thumb_queue.task_done()
            except queue.Empty:
                continue

    def _update_item_thumbnail(self, item_id, cache_key, img_data):
        if not self.tree.exists(item_id):
            return
        photo = tk.PhotoImage(data=img_data)
        self.thumbnails[cache_key] = photo
        self.tree.item(item_id, image=photo)

    def _apply_cached_thumbnail(self, item_id, cache_key):
        if self.tree.exists(item_id) and cache_key in self.thumbnails:
            self.tree.item(item_id, image=self.thumbnails[cache_key])

    def _add_files(self):
        if self.is_converting:
            return
        files = filedialog.askopenfilenames(
            title="選擇檔案", 
            filetypes=[("支援格式", "*.jpg *.jpeg *.png *.pdf *.bmp *.tiff")]
        )
        if files:
            self._process_incoming_files(files)

    def _handle_drop(self, event):
        if self.is_converting:
            return
        files = parse_dropped_files(event.data)
        self._process_incoming_files(files)

    def _process_incoming_files(self, files):
        valid_exts = ('.jpg', '.jpeg', '.png', '.pdf', '.bmp', '.tiff')
        added = False
        
        for f in files:
            if not f.lower().endswith(valid_exts):
                continue
            exists = any(item['path'] == f and item['page'] is None for item in self.file_list)
            if exists:
                continue
                
            count = 1
            if f.lower().endswith('.pdf'):
                doc = self._get_pdf_doc(f)
                if doc:
                    if doc.is_encrypted and not self.passwords.get(f):
                        correct = False
                        while not correct:
                            dialog = ModernPasswordDialog(self.app, os.path.basename(f))
                            self.app.wait_window(dialog)
                            if dialog.password is None:
                                break
                            if doc.authenticate(dialog.password):
                                self.passwords[f] = dialog.password
                                correct = True
                            else:
                                messagebox.showerror("錯誤", "密碼不正確")
                        if not correct:
                            continue
                    count = len(doc)
            self.file_list.append({'path': f, 'page': None, 'page_count': count})
            added = True
            
        if added:
            self._update_tree_content()

    def _update_tree_content(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for idx, item in enumerate(self.file_list):
            path = item['path']
            fname = os.path.basename(path)
            
            if path.lower().endswith('.pdf'):
                t_str = "PDF 文件"
                if item['page'] is not None:
                    idx_str = f"{idx+1} (頁 {item['page']+1})"
                    name_str = f"{fname} - 第 {item['page']+1} 頁"
                else:
                    idx_str = f"{idx+1}"
                    name_str = f"{fname} (共 {item['page_count']} 頁, 未展開)"
            else:
                t_str = "圖片"
                idx_str = f"{idx+1}"
                name_str = fname
                
            item_id = self.tree.insert("", tk.END, values=(idx_str, t_str, name_str))
            
            # 排入縮圖載入佇列
            p_idx = item['page'] if item['page'] is not None else 0
            self.thumb_queue.put((item_id, path, p_idx))
            
        self.lbl_count.configure(text=f"已選擇: {len(self.file_list)} 個項目")
        
        if not self.file_list:
            self.lbl_empty_tip.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.lbl_empty_tip.place_forget()

    def _expand_selected_pdf(self):
        if self.is_converting:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "請先選擇列表中未展開的 PDF 項目。")
            return
            
        new_list = []
        sel_indices = [self.tree.index(i) for i in sel]
        
        for idx, item in enumerate(self.file_list):
            if idx in sel_indices and item['path'].lower().endswith('.pdf') and item['page'] is None:
                doc = self._get_pdf_doc(item['path'])
                if doc:
                    for p in range(len(doc)):
                        new_list.append({'path': item['path'], 'page': p, 'page_count': 1})
            else:
                new_list.append(item)
                
        self.file_list = new_list
        self._update_tree_content()

    def _move_up(self):
        if self.is_converting:
            return
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted([self.tree.index(i) for i in sel])
        if idxs[0] == 0:
            return
            
        for idx in idxs:
            self.file_list[idx], self.file_list[idx-1] = self.file_list[idx-1], self.file_list[idx]
            
        self._update_tree_content()
        # 保持原本的選取狀態
        children = self.tree.get_children()
        for idx in idxs:
            self.tree.selection_add(children[idx-1])

    def _move_down(self):
        if self.is_converting:
            return
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
        if idxs[0] == len(self.file_list) - 1:
            return
            
        for idx in idxs:
            self.file_list[idx], self.file_list[idx+1] = self.file_list[idx+1], self.file_list[idx]
            
        self._update_tree_content()
        # 保持選取
        children = self.tree.get_children()
        for idx in idxs:
            self.tree.selection_add(children[idx+1])

    def _sort_files(self, rev):
        if self.is_converting:
            return
        self.file_list.sort(key=lambda x: (
            os.path.basename(x['path']).lower(), 
            x['page'] if x['page'] is not None else -1
        ), reverse=rev)
        self._update_tree_content()

    def _remove_selected(self):
        if self.is_converting:
            return
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
        
        for idx in idxs:
            item = self.file_list.pop(idx)
            if not any(it['path'] == item['path'] for it in self.file_list):
                if item['path'] in self.doc_handles:
                    try:
                        self.doc_handles[item['path']].close()
                    except:
                        pass
                    del self.doc_handles[item['path']]
                self.passwords.pop(item['path'], None)
                keys = [k for k in self.thumbnails if k.startswith(item['path'])]
                for k in keys:
                    del self.thumbnails[k]
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
            idx = int(values[0].split()[0]) - 1
            path = self.file_list[idx]['path']
            import subprocess
            subprocess.run(f'explorer /select,"{os.path.normpath(path)}"')
        except Exception as e:
            pass

    def _clear_all(self):
        if self.is_converting or not self.file_list:
            return
        if messagebox.askyesno("確認", "是否確定清空所選的項目？"):
            for p in self.doc_handles.values():
                try:
                    p.close()
                except:
                    pass
            self.doc_handles.clear()
            self.file_list.clear()
            self.passwords.clear()
            self.thumbnails.clear()
            self._update_tree_content()

    def _on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        all_ids = self.tree.get_children()
        try:
            idx = all_ids.index(item_id)
            item = self.file_list[idx]
            self._show_preview(item)
        except Exception:
            pass

    def _on_drag_start(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell" and region != "tree":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self._dragged_item = item
        self._is_dragging = True

    def _on_drag_motion(self, event):
        if not hasattr(self, "_is_dragging") or not self._is_dragging:
            return
        if not self._dragged_item:
            return
            
        target_item = self.tree.identify_row(event.y)
        if not target_item or target_item == self._dragged_item:
            return
            
        children = self.tree.get_children()
        try:
            drag_idx = children.index(self._dragged_item)
            target_idx = children.index(target_item)
            
            # 即時對調數據層
            item_data = self.file_list.pop(drag_idx)
            self.file_list.insert(target_idx, item_data)
            
            # 即時重新渲染列表
            self._update_tree_content()
            
            # 重新定位被拖曳項目的識別 ID
            new_children = self.tree.get_children()
            self._dragged_item = new_children[target_idx]
            
            # 將該項目保持高亮選中與焦點
            self.tree.selection_set(self._dragged_item)
            self.tree.focus(self._dragged_item)
        except Exception as e:
            pass

    def _on_drag_drop(self, event):
        self._dragged_item = None
        self._is_dragging = False

    def _show_preview(self, item):
        path = item['path']
        page_idx = item['page'] if item['page'] is not None else 0
        
        preview_win = ctk.CTkToplevel(self)
        preview_win.title(f"高清預覽：{os.path.basename(path)}")
        preview_win.configure(fg_color=("#1A1A1A", "#1A1A1A"))
        preview_win.transient(self.app)
        preview_win.grab_set()

        screen_h = self.winfo_screenheight()
        max_h = int(screen_h * 0.8)
        
        try:
            with fitz.open(path) as doc:
                if doc.is_encrypted:
                    doc.authenticate(self.passwords.get(path, ""))
                page = doc[page_idx]
                rect = page.rect
                zoom = max_h / rect.height
                zoom = min(zoom, 2.0)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img_data = pix.tobytes("png")
                
                photo = tk.PhotoImage(data=img_data)
                preview_win.photo = photo
                
                lbl = tk.Label(preview_win, image=photo, cursor="hand2", bd=0, bg="#1A1A1A")
                lbl.pack(padx=10, pady=10)
                
                lbl.bind("<Button-1>", lambda e: preview_win.destroy())
                preview_win.bind("<Key>", lambda e: preview_win.destroy())
                
                preview_win.update_idletasks()
                w, h = preview_win.winfo_width(), preview_win.winfo_height()
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                preview_win.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
        except Exception as e:
            messagebox.showerror("預覽失敗", f"無法渲染預覽圖：\n{str(e)}")
            preview_win.destroy()

    def _start_conversion(self):
        if not self.file_list:
            messagebox.showwarning("提示", "清單中尚無檔案，請先加入圖片或 PDF。")
            return
            
        pdf_settings = self.pdf_features.get_settings()
        if pdf_settings["encrypt"] and not pdf_settings["password"]:
            messagebox.showwarning("警告", "您啟用了加密，但尚未設定開啟密碼。")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 檔案", "*.pdf")])
        if not save_path:
            return
            
        settings = {
            "save_path": save_path,
            "compress": self.compress_var.get(),
            "quality": int(self.slider_quality.get()),
            "grayscale": self.grayscale_var.get(),
            "auto_rotate": self.auto_rotate_var.get(),
            "scale_mode": self.combo_scale.get(),
            "flatten": self.flatten_var.get(),
            "page_size": self.combo_size.get(),
            "orient": self.combo_orient.get(),
            "pdf_settings": pdf_settings
        }
        
        self.is_converting = True
        self._toggle_ui_state("disabled")
        self.progress.set(0)
        self.lbl_status.configure(text="準備開始轉換...", text_color=("#2563EB", "#3B82F6"))
        
        # 啟動 Worker 執行緒
        threading.Thread(target=self._conversion_worker, args=(settings,), daemon=True).start()

    def _toggle_ui_state(self, state):
        mode = "normal" if state == "normal" else "disabled"
        self.btn_run.configure(state=mode)
        self.btn_add.configure(state=mode)
        self.btn_expand.configure(state=mode)
        self.btn_up.configure(state=mode)
        self.btn_down.configure(state=mode)
        self.btn_sort_asc.configure(state=mode)
        self.btn_sort_desc.configure(state=mode)
        self.btn_remove.configure(state=mode)
        self.btn_clear.configure(state=mode)
        self.combo_size.configure(state=mode)
        self.combo_orient.configure(state=mode)
        self.combo_scale.configure(state=mode)
        self.check_compress.configure(state=mode)
        self.check_grayscale.configure(state=mode)
        self.check_auto_rotate.configure(state=mode)
        self.check_flatten.configure(state=mode)
        
        # 連動 PDF 高級設定元件
        self.pdf_features.configure_state(state)
        
        if mode == "normal":
            self._toggle_compress()
        else:
            self.slider_quality.configure(state="disabled")

    def _conversion_worker(self, settings):
        """核心背景轉換 Worker"""
        save_path = settings["save_path"]
        pdf_settings = settings["pdf_settings"]
        doc = fitz.open()
        total_pages = sum(item.get('page_count', 1) for item in self.file_list)
        processed_pages = 0
        
        # 讀取介面參數
        c = settings["compress"]
        q = settings["quality"]
        gs = settings["grayscale"]
        ar = settings["auto_rotate"]
        sm = settings["scale_mode"]
        flatten = settings["flatten"]
        
        base_size = PAGE_SIZES.get(settings["page_size"])
        target_orient = settings["orient"]
        
        HIGH_RES_DPI = 300 / 72.0  # 300 DPI 高清渲染因子
        
        try:
            for item in self.file_list:
                path = item['path']
                
                # --- A. 處理圖片檔案 ---
                if not path.lower().endswith('.pdf'):
                    processed_pages += 1
                    self.app.queue.put(("t1_status", (f"正在處理圖片：第 {processed_pages}/{total_pages} 頁", processed_pages / total_pages)))
                    
                    pix = fitz.Pixmap(path)
                    
                    if gs:
                        pix = fitz.Pixmap(fitz.csGRAY, pix)
                    
                    if pix.alpha:
                        new_pix = fitz.Pixmap(fitz.csRGB if not gs else fitz.csGRAY, pix.width, pix.height, 0)
                        new_pix.clear_with(255)
                        new_pix.copy(pix, pix.irect)
                        pix = new_pix
                        
                    if c:
                        img_data = pix.tobytes("jpg", jpg_quality=q)
                    else:
                        img_data = pix.tobytes("png")
                        
                    if base_size:
                        tw, th = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                        if ar and ((pix.width > pix.height) != (tw > th)):
                            tw, th = th, tw
                        page = doc.new_page(width=tw, height=th)
                        rect = page.rect if sm == "自動填滿" else fitz.Rect(0, 0, pix.width, pix.height)
                        page.insert_image(rect, stream=img_data, keep_proportion=True)
                    else:
                        page = doc.new_page(width=pix.width, height=pix.height)
                        page.insert_image(page.rect, stream=img_data)
                    
                    if wm_enabled and wm_text:
                        apply_watermark_to_page(page, wm_text, wm_opacity, wm_angle, wm_tile, wm_size, wm_color)
                        
                    pix = None
                    
                # --- B. 處理 PDF 檔案 ---
                else:
                    with fitz.open(path) as sub:
                        if sub.is_encrypted:
                            sub.authenticate(self.passwords.get(path, ""))
                            
                        from_p = item['page'] if item['page'] is not None else 0
                        to_p = item['page'] if item['page'] is not None else len(sub) - 1
                        
                        for p_no in range(from_p, to_p + 1):
                            processed_pages += 1
                            self.app.queue.put(("t1_status", (f"正在處理 PDF 頁面：第 {processed_pages}/{total_pages} 頁", processed_pages / total_pages)))
                            
                            sp = sub[p_no]
                            
                            if flatten or gs or base_size:
                                pix = sp.get_pixmap(matrix=fitz.Matrix(HIGH_RES_DPI, HIGH_RES_DPI))
                                if gs:
                                    pix = fitz.Pixmap(fitz.csGRAY, pix)
                                if pix.alpha:
                                    new_pix = fitz.Pixmap(fitz.csRGB if not gs else fitz.csGRAY, pix.width, pix.height, 0)
                                    new_pix.clear_with(255)
                                    new_pix.copy(pix, pix.irect)
                                    pix = new_pix
                                    
                                img_data = pix.tobytes("jpg", jpg_quality=q if c else 95)
                                
                                if base_size:
                                    tw, th = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                                    if ar and ((sp.rect.width > sp.rect.height) != (tw > th)):
                                        tw, th = th, tw
                                    page = doc.new_page(width=tw, height=th)
                                    rect = page.rect if sm == "自動填滿" else sp.rect
                                    page.insert_image(rect, stream=img_data, keep_proportion=True)
                                else:
                                    page = doc.new_page(width=sp.rect.width, height=sp.rect.height)
                                    page.insert_image(page.rect, stream=img_data)
                                    
                                pix = None
                            else:
                                # 純頁面合併
                                if base_size:
                                    tw, th = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                                    lw, lh = (th, tw) if ar and ((sp.rect.width > sp.rect.height) != (tw > th)) else (tw, th)
                                    page = doc.new_page(width=lw, height=lh)
                                    rect = page.rect if sm == "自動填滿" else sp.rect
                                    page.show_pdf_page(rect, sub, sp.number)
                                else:
                                    doc.insert_pdf(sub, from_page=p_no, to_page=p_no)
                                    
            finalize_and_save_pdf(doc, save_path, pdf_settings)
            doc.close()
            
            self.app.queue.put(("t1_done", save_path))
        except Exception as e:
            self.app.queue.put(("t1_error", str(e)))

    def destroy(self):
        """重寫銷毀方法以停止背景載入線程"""
        self.thumb_thread_running = False
        super().destroy()
