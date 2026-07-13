import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import fitz
from tkinterdnd2 import DND_FILES

from utils.helpers import SYSTEM_FONT, FONT_OFFSET, parse_dropped_files, unique_filename, finalize_and_save_pdf
from utils.icons import get_icon
from components.pdf_features import PDFFeaturesFrame

class TabPDFCompress(ctk.CTkFrame):
    """PDF 壓縮功能分頁"""
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
        t4_main = ctk.CTkFrame(self, fg_color="transparent")
        t4_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：PDF 檔案清單
        left_pane = ctk.CTkFrame(t4_main, fg_color="transparent")
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        t4_ctrl = ctk.CTkFrame(left_pane, fg_color="transparent", height=40)
        t4_ctrl.pack(fill=tk.X, pady=(0, 8))
        
        self.btn_add = ctk.CTkButton(t4_ctrl, text=" ＋ 選擇 PDF 檔案... ", 
                                         fg_color=("#2563EB", "#3B82F6"),
                                         text_color="white",
                                         hover_color=("#1D4ED8", "#2563EB"),
                                         font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                         command=self._select_pdfs,
                                         height=32)
        self.btn_add.pack(side=tk.LEFT)
        
        t4_tips = ctk.CTkLabel(t4_ctrl, text="支援拖放多個 PDF 檔案至清單中", 
                               font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                               text_color=("#6B7280", "#9CA3AF"))
        t4_tips.pack(side=tk.LEFT, padx=12)
        
        self.lbl_count = ctk.CTkLabel(t4_ctrl, text="已選擇: 0 個檔案",
                                         font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                         text_color=("#2563EB", "#3B82F6"))
        self.lbl_count.pack(side=tk.RIGHT)
        
        # Treeview 與其捲軸
        tree_container = ctk.CTkFrame(left_pane, fg_color="transparent")
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_container, columns=("Index", "Name", "Size", "Status"), show='headings', selectmode='extended', style="T4.Treeview")
        self.tree.heading("Index", text="序號")
        self.tree.heading("Name", text="檔案路徑")
        self.tree.heading("Size", text="原始大小")
        self.tree.heading("Status", text="狀態")
        
        self.tree.column("Index", width=60, anchor="center", stretch=False)
        self.tree.column("Name", width=420, anchor="w")
        self.tree.column("Size", width=100, anchor="center", stretch=False)
        self.tree.column("Status", width=100, anchor="center", stretch=False)
        
        t4_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=t4_scroll.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        t4_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self._handle_drop)
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        
        # 空白清單引導 Label (使用大匯入圖示，置於文字上方，移除 Emoji)
        import_icon = get_icon("import", size=(32, 32))
        self.lbl_empty_tip = ctk.CTkLabel(
            self.tree, 
            text="\n拖曳 PDF 檔案至此處，或點擊選擇檔案新增",
            image=import_icon,
            compound="top",
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
        self.context_menu.add_command(label="📁 在檔案總管中顯示", command=self._show_in_explorer)
        self.context_menu.add_command(label="❌ 移除此項目", command=self._remove_selected)
        
        self.tree.bind("<Button-3>", self._show_context_menu)
        
        # 列表操作按鈕
        t4_list_ctrl = ctk.CTkFrame(left_pane, fg_color="transparent", height=40)
        t4_list_ctrl.pack(fill=tk.X, pady=(8, 0))
        
        self.btn_remove = ctk.CTkButton(t4_list_ctrl, text="移除項目", 
                                            fg_color=("#FEF2F2", "#450A0A"),
                                            text_color=("#DC2626", "#F87171"),
                                            hover_color=("#FEE2E2", "#7F1D1D"),
                                            font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                            command=self._remove_selected,
                                            width=100, height=32)
        self.btn_remove.pack(side=tk.LEFT)
        
        self.btn_clear = ctk.CTkButton(t4_list_ctrl, text="全部清空", 
                                           fg_color=("#FEF2F2", "#450A0A"),
                                           text_color=("#DC2626", "#F87171"),
                                           hover_color=("#FEE2E2", "#7F1D1D"),
                                           font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                           command=self._clear_all,
                                           width=100, height=32)
        self.btn_clear.pack(side=tk.LEFT, padx=10)
        
        # 右側：執行按鈕與參數設定
        right_pane_outer = ctk.CTkFrame(t4_main, width=340, fg_color=("#FFFFFF", "#1F2937"),
                                        border_width=1, border_color=("#E5E7EB", "#374151"))
        right_pane_outer.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        right_pane_outer.pack_propagate(False)
        
        right_pane = ctk.CTkFrame(right_pane_outer, fg_color="transparent")
        right_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 標題 (使用齒輪圖示，移除 Emoji)
        settings_icon = get_icon("settings", size=(16, 16))
        t4_setting_lbl = ctk.CTkLabel(right_pane, text="  執行與壓縮設定", 
                                       image=settings_icon,
                                       compound="left",
                                       font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        t4_setting_lbl.pack(anchor="w", pady=(0, 10))
        
        # 執行區域
        exec_frame = ctk.CTkFrame(right_pane, fg_color="transparent")
        exec_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.auto_open_var = tk.BooleanVar(value=False)
        self.check_open = ctk.CTkCheckBox(exec_frame, text="完成後開啟資料夾", 
                                             variable=self.auto_open_var,
                                             font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_open.pack(anchor="w", pady=(0, 8))
        
        self.progress = ctk.CTkProgressBar(exec_frame)
        self.progress.set(0)
        self.progress.pack(fill=tk.X, pady=(0, 8))
        
        self.btn_run_frame = ctk.CTkFrame(exec_frame, fg_color="transparent", height=44)
        self.btn_run_frame.pack(fill=tk.X)
        self.btn_run_frame.pack_propagate(False)
        
        # 開始壓縮按鈕 (使用播放圖示，移除 Emoji)
        run_icon = get_icon("run", size=(16, 16))
        self.btn_run = ctk.CTkButton(self.btn_run_frame, text="  開始壓縮", 
                                         image=run_icon,
                                         compound="left",
                                         fg_color=("#2563EB", "#3B82F6"),
                                         text_color="white",
                                         hover_color=("#1D4ED8", "#2563EB"),
                                         font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
                                         command=self._start_compress,
                                         height=44)
        self.btn_run.pack(fill=tk.BOTH, expand=True)
        
        # 終止作業按鈕 (使用終止圖示，移除 Emoji)
        cancel_icon = get_icon("cancel", size=(16, 16))
        self.btn_cancel = ctk.CTkButton(self.btn_run_frame, text="  終止作業", 
                                             image=cancel_icon,
                                             compound="left",
                                             fg_color=("#DC2626", "#EF4444"),
                                             text_color="white",
                                             hover_color=("#B91C1C", "#DC2626"),
                                             font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
                                             command=self._cancel_compress,
                                             height=44)
                                            
        ctk.CTkFrame(right_pane, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 滾動設定區
        t4_scroll_settings = ctk.CTkScrollableFrame(right_pane, fg_color="transparent", label_text="")
        t4_scroll_settings.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        lbl_mode = ctk.CTkLabel(t4_scroll_settings, text="PDF 壓縮模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_mode.pack(anchor="w", pady=(4, 1))
        
        self.compress_mode_var = tk.StringVar(value="lossy")
        self.radio1 = ctk.CTkRadioButton(t4_scroll_settings, text="圖片降階重壓縮 (效果顯著)", 
                                            variable=self.compress_mode_var, value="lossy",
                                            command=self._toggle_mode,
                                            font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.radio1.pack(anchor="w", pady=4)
        
        self.radio2 = ctk.CTkRadioButton(t4_scroll_settings, text="無損垃圾清掃優化", 
                                            variable=self.compress_mode_var, value="lossless",
                                            command=self._toggle_mode,
                                            font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.radio2.pack(anchor="w", pady=4)
        
        # 黑白/灰階 PDF 勾選框
        self.grayscale_var = tk.BooleanVar(value=False)
        self.check_grayscale = ctk.CTkCheckBox(t4_scroll_settings, text="轉換為黑白/灰階 PDF (適合省墨列印)", 
                                                 variable=self.grayscale_var,
                                                 font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.check_grayscale.pack(anchor="w", pady=4)
        
        self.quality_frame = ctk.CTkFrame(t4_scroll_settings, fg_color="transparent")
        self.quality_frame.pack(fill=tk.X, pady=(4, 8))
        
        self.lbl_quality = ctk.CTkLabel(self.quality_frame, text="壓縮品質: 60%", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.lbl_quality.pack(side=tk.LEFT)
        
        self.slider_quality = ctk.CTkSlider(self.quality_frame, from_=10, to=90, 
                                               number_of_steps=80, height=16,
                                               command=self._update_quality_lbl)
        self.slider_quality.set(60)
        self.slider_quality.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        ctk.CTkFrame(t4_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 4. PDF 高級功能設定 (加密、元資料、浮水印) - 重用元件
        self.pdf_features = PDFFeaturesFrame(t4_scroll_settings)
        self.pdf_features.pack(fill=tk.X, pady=(5, 10))
        
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
        
        self.log_text = ctk.CTkTextbox(right_pane, font=("Consolas", 9 + FONT_OFFSET), 
                                          fg_color=("#F9FAFB", "#1A1A1A"),
                                          text_color=("#374151", "#E5E7EB"),
                                          border_width=1, border_color=("#E5E7EB", "#374151"),
                                          height=100)
        
        self._toggle_mode()

    def _toggle_mode(self):
        if self.compress_mode_var.get() == "lossy":
            self.slider_quality.configure(state="normal")
            self.lbl_quality.configure(text_color=("black", "white"))
        else:
            self.slider_quality.configure(state="disabled")
            self.lbl_quality.configure(text_color="#9CA3AF")

    def _update_quality_lbl(self, val):
        self.lbl_quality.configure(text=f"壓縮品質: {int(val)}%")

    def _select_pdfs(self):
        if self.is_converting:
            return
        files = filedialog.askopenfilenames(title="選擇 PDF", filetypes=[("PDF", "*.pdf")])
        if files:
            self._process_incoming_files(files)

    def _handle_drop(self, event):
        if self.is_converting:
            return
        files = parse_dropped_files(event.data)
        self._process_incoming_files(files)

    def _process_incoming_files(self, files):
        added = False
        for f in files:
            if not f.lower().endswith('.pdf'):
                continue
            if f in self.selected_files:
                continue
            self.selected_files.append(f)
            added = True
        if added:
            self._update_tree_content()

    def _update_tree_content(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, f in enumerate(self.selected_files):
            size_text = "未知"
            status_text = "等待中"
            try:
                sz = os.path.getsize(f)
                if sz >= 1024 * 1024:
                    size_text = f"{sz / (1024*1024):.2f} MB"
                else:
                    size_text = f"{sz / 1024:.1f} KB"
            except:
                pass
            self.tree.insert("", tk.END, values=(idx + 1, f, size_text, status_text))
        self.lbl_count.configure(text=f"已選擇: {len(self.selected_files)} 個檔案")
        
        if not self.selected_files:
            self.lbl_empty_tip.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.lbl_empty_tip.place_forget()

    def _remove_selected(self):
        if self.is_converting:
            return
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
        for idx in idxs:
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

    def _clear_all(self):
        if self.is_converting or not self.selected_files:
            return
        if messagebox.askyesno("確認", "是否確定清空所選的 PDF 檔案？"):
            self.selected_files.clear()
            self._update_tree_content()

    def _log(self, msg):
        self.app.queue.put(("t4_log", msg))

    def update_log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def _toggle_log_view(self):
        if self.show_log_var.get():
            self.log_text.pack(fill=tk.X, expand=False, pady=(0, 5))
        else:
            self.log_text.pack_forget()

    def _toggle_ui_state(self, state):
        mode = "normal" if state == "normal" else "disabled"
        self.btn_add.configure(state=mode)
        self.btn_remove.configure(state=mode)
        self.btn_clear.configure(state=mode)
        self.radio1.configure(state=mode)
        self.radio2.configure(state=mode)
        self.check_open.configure(state=mode)
        self.check_grayscale.configure(state=mode)
        self.check_show_log.configure(state=mode)
        
        # 連動 PDF 高級設定元件
        self.pdf_features.configure_state(state)
        
        if mode == "normal":
            self._toggle_mode()
        else:
            self.slider_quality.configure(state="disabled")

    def _start_compress(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "請先選擇 PDF 檔案。")
            return
            
        pdf_settings = self.pdf_features.get_settings()
        if pdf_settings["encrypt"] and not pdf_settings["password"]:
            messagebox.showwarning("警告", "您啟用了加密，但尚未設定開啟密碼。")
            return
        
        mode = self.compress_mode_var.get()
        q = int(self.slider_quality.get())
        
        settings = {
            "mode": mode,
            "quality": q,
            "open": self.auto_open_var.get(),
            "grayscale": self.grayscale_var.get(),
            "pdf_settings": pdf_settings
        }
        
        self.is_converting = True
        self.btn_run.pack_forget()
        self.btn_cancel.pack(fill=tk.BOTH, expand=True)
        self._toggle_ui_state("disabled")
        
        self.progress.set(0)
        self.log_text.delete("1.0", tk.END)
        self.update_log("===============================")
        self.update_log("🚀 PDF 壓縮優化作業開始...")
        self.stop_event.clear()
        
        threading.Thread(target=self._compress_worker, args=(settings,), daemon=True).start()

    def _cancel_compress(self):
        if messagebox.askyesno("取消", "確定要終止目前的壓縮作業？"):
            self.stop_event.set()
            self.btn_cancel.configure(state="disabled")
            self._log("🛑 正在停止作業，請稍後...")

    def _compress_worker(self, settings):
        """PDF 壓縮與灰階背景 Worker"""
        try:
            tasks = []
            
            for f in self.selected_files:
                if self.stop_event.is_set():
                    raise InterruptedError()
                    
                pw = None
                try:
                    with fitz.open(f) as test_doc:
                        if test_doc.is_encrypted:
                            correct = False
                            while not correct:
                                evt = threading.Event()
                                res = {}
                                self.app.queue.put(("ask_pw", (f, evt, res)))
                                evt.wait()
                                pw = res.get("pw")
                                if pw is None:
                                    break
                                if test_doc.authenticate(pw):
                                    correct = True
                                else:
                                    self._log(f"❌ 檔案「{os.path.basename(f)}」密碼解鎖失敗。")
                                    self.app.queue.put(("error_msg", "密碼不正確，略過此檔案。"))
                            if not correct:
                                continue
                except Exception as e:
                    self._log(f"❌ 無法讀取檔案：{os.path.basename(f)} ({str(e)})")
                    continue
                
                tasks.append({"path": f, "pw": pw})
            
            if not tasks:
                self.app.queue.put(("t4_error", "清單中無有效檔案可供壓縮！"))
                return
                
            total_tasks = len(tasks)
            self._log(f"📊 分析完成：共 {total_tasks} 個 PDF 檔案待優化")
            
            out_path = ""
            
            for idx, task in enumerate(tasks):
                if self.stop_event.is_set():
                    raise InterruptedError()
                    
                pdf_path = task["path"]
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                out_dir = os.path.dirname(pdf_path)
                
                out_path = os.path.join(out_dir, f"{base_name}_compressed.pdf")
                out_path = unique_filename(out_path)
                
                self._log(f"📂 正在處理：{base_name}")
                
                try:
                    doc = fitz.open(pdf_path)
                    if doc.is_encrypted:
                        doc.authenticate(task["pw"])
                        
                    orig_size = os.path.getsize(pdf_path)
                    
                    if settings["grayscale"]:
                        self._log(f"  ➜ 正在執行 PDF 一鍵轉灰階 (多執行緒並行渲染)...")
                        
                        import concurrent.futures
                        max_workers = min(4, os.cpu_count() or 1)
                        self._log(f"  ➜ 啟動多執行緒加速渲染 (並行數: {max_workers})...")
                        
                        def render_single_page(page_num):
                            if self.stop_event.is_set():
                                return None
                            try:
                                with fitz.open(pdf_path) as doc_t:
                                    if doc_t.is_encrypted:
                                        doc_t.authenticate(task["pw"])
                                    page = doc_t[page_num]
                                    HIGH_RES_DPI = 300 / 72.0
                                    pix = page.get_pixmap(matrix=fitz.Matrix(HIGH_RES_DPI, HIGH_RES_DPI))
                                    pix_gs = fitz.Pixmap(fitz.csGRAY, pix)
                                    
                                    if pix_gs.alpha:
                                        new_pix = fitz.Pixmap(fitz.csGRAY, pix_gs.width, pix_gs.height, 0)
                                        new_pix.clear_with(255)
                                        new_pix.copy(pix_gs, pix_gs.irect)
                                        pix_gs = new_pix
                                        
                                    img_data = pix_gs.tobytes("jpg", jpg_quality=settings["quality"])
                                    return page_num, img_data, page.rect.width, page.rect.height
                            except Exception as pe:
                                self._log(f"  ⚠️ 第 {page_num+1} 頁渲染失敗：{str(pe)}")
                                return None
                        
                        page_nums = list(range(len(doc)))
                        
                        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                            results = list(executor.map(render_single_page, page_nums))
                            
                        if self.stop_event.is_set():
                            raise InterruptedError()
                            
                        out_doc = fitz.open()
                        for res in results:
                            if res is None:
                                raise RuntimeError("有頁面渲染失敗，無法產生完整的 PDF 檔案。")
                            p_num, img_data, w, h = res
                            new_page = out_doc.new_page(width=w, height=h)
                            new_page.insert_image(new_page.rect, stream=img_data)
                            
                        finalize_and_save_pdf(out_doc, out_path, settings["pdf_settings"])
                        out_doc.close()
                        doc.close()
                        
                    elif settings["mode"] == "lossy":
                        self._log(f"  ➜ 正在掃描並壓縮內置影像...")
                        img_count = 0
                        processed_xrefs = set()
                        
                        for page_num in range(len(doc)):
                            if self.stop_event.is_set():
                                raise InterruptedError()
                                
                            page = doc[page_num]
                            img_list = page.get_images()
                            
                            for img_info in img_list:
                                xref = img_info[0]
                                if xref in processed_xrefs:
                                    continue
                                
                                try:
                                    base_img = doc.extract_image(xref)
                                    if base_img:
                                        img_bytes = base_img["image"]
                                        from PIL import Image
                                        import io
                                        pil_img = Image.open(io.BytesIO(img_bytes))
                                        
                                        out_io = io.BytesIO()
                                        pil_img.convert("RGB").save(out_io, "JPEG", quality=settings["quality"])
                                        new_data = out_io.getvalue()
                                        
                                        page.replace_image(xref, stream=new_data)
                                        processed_xrefs.add(xref)
                                        img_count += 1
                                except Exception as img_err:
                                    pass
                                import time
                                time.sleep(0.015)  # 釋放 GIL 與 UI 更新喘息時間，防止介面凍結
                                    
                        self._log(f"  ➜ 成功重壓 {img_count} 張圖片資源")
                        finalize_and_save_pdf(doc, out_path, settings["pdf_settings"])
                        doc.close()
                        
                    else:
                        finalize_and_save_pdf(doc, out_path, settings["pdf_settings"])
                        doc.close()
                    
                    new_size = os.path.getsize(out_path)
                    saved = orig_size - new_size
                    pct = (saved / orig_size) * 100 if orig_size > 0 else 0
                    
                    if saved > 0:
                        self._log(f"  ➜ 壓縮完成！體積減少了 {pct:.1f}% ({saved / (1024*1024):.2f} MB)")
                        self._log(f"  ➜ 儲存為：{os.path.basename(out_path)}")
                    else:
                        self._log(f"  ➜ 整理完成 (此檔案已無優化空間)。儲存至：{os.path.basename(out_path)}")
                        
                except Exception as e:
                    self._log(f"❌ 檔案處理出錯：{base_name} ({str(e)})")
                    
                self.progress.set((idx + 1) / total_tasks)
                
            self.app.queue.put(("t4_done", out_path))
            
        except InterruptedError:
            self.app.queue.put(("t4_cancelled", None))
        except Exception as e:
            self.app.queue.put(("t4_error", str(e)))
