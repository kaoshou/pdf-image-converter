import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import fitz
from tkinterdnd2 import DND_FILES

from utils.helpers import SYSTEM_FONT, FONT_OFFSET, parse_dropped_files, unique_filename, parse_range_string, finalize_and_save_pdf
from components.pdf_features import PDFFeaturesFrame

class TabPDFSplit(ctk.CTkFrame):
    """PDF 拆分與擷取的功能分頁"""
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
        t3_main = ctk.CTkFrame(self, fg_color="transparent")
        t3_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：PDF 檔案清單
        left_pane = ctk.CTkFrame(t3_main, fg_color="transparent")
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        t3_ctrl = ctk.CTkFrame(left_pane, fg_color="transparent", height=40)
        t3_ctrl.pack(fill=tk.X, pady=(0, 8))
        
        self.btn_add = ctk.CTkButton(t3_ctrl, text=" ＋ 選擇 PDF 檔案... ", 
                                         fg_color=("#2563EB", "#3B82F6"),
                                         text_color="white",
                                         hover_color=("#1D4ED8", "#2563EB"),
                                         font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                         command=self._select_pdfs,
                                         height=32)
        self.btn_add.pack(side=tk.LEFT)
        
        t3_tips = ctk.CTkLabel(t3_ctrl, text="支援拖放多個 PDF 檔案至清單中", 
                               font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                               text_color=("#6B7280", "#9CA3AF"))
        t3_tips.pack(side=tk.LEFT, padx=12)
        
        self.lbl_count = ctk.CTkLabel(t3_ctrl, text="已選擇: 0 個檔案",
                                         font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"),
                                         text_color=("#2563EB", "#3B82F6"))
        self.lbl_count.pack(side=tk.RIGHT)
        
        # Treeview 與其捲軸
        tree_container = ctk.CTkFrame(left_pane, fg_color="transparent")
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_container, columns=("Index", "Name", "Pages", "Status"), show='headings', selectmode='extended', style="T3.Treeview")
        self.tree.heading("Index", text="序號")
        self.tree.heading("Name", text="檔案路徑")
        self.tree.heading("Pages", text="總頁數")
        self.tree.heading("Status", text="狀態")
        
        self.tree.column("Index", width=60, anchor="center", stretch=False)
        self.tree.column("Name", width=420, anchor="w")
        self.tree.column("Pages", width=80, anchor="center", stretch=False)
        self.tree.column("Status", width=120, anchor="center", stretch=False)
        
        t3_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=t3_scroll.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        t3_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self._handle_drop)
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        
        # 空白清單引導 Label
        self.lbl_empty_tip = ctk.CTkLabel(
            self.tree, 
            text="📥 拖曳 PDF 檔案至此處，或點擊選擇檔案新增",
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
        t3_list_ctrl = ctk.CTkFrame(left_pane, fg_color="transparent", height=40)
        t3_list_ctrl.pack(fill=tk.X, pady=(8, 0))
        
        self.btn_remove = ctk.CTkButton(t3_list_ctrl, text="移除項目", 
                                            fg_color=("#FEF2F2", "#450A0A"),
                                            text_color=("#DC2626", "#F87171"),
                                            hover_color=("#FEE2E2", "#7F1D1D"),
                                            font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                            command=self._remove_selected,
                                            width=100, height=32)
        self.btn_remove.pack(side=tk.LEFT)
        
        self.btn_clear = ctk.CTkButton(t3_list_ctrl, text="全部清空", 
                                           fg_color=("#FEF2F2", "#450A0A"),
                                           text_color=("#DC2626", "#F87171"),
                                           hover_color=("#FEE2E2", "#7F1D1D"),
                                           font=(SYSTEM_FONT, 10 + FONT_OFFSET),
                                           command=self._clear_all,
                                           width=100, height=32)
        self.btn_clear.pack(side=tk.LEFT, padx=10)
        
        # 右側：執行按鈕與參數設定
        right_pane_outer = ctk.CTkFrame(t3_main, width=340, fg_color=("#FFFFFF", "#1F2937"),
                                        border_width=1, border_color=("#E5E7EB", "#374151"))
        right_pane_outer.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        right_pane_outer.pack_propagate(False)
        
        right_pane = ctk.CTkFrame(right_pane_outer, fg_color="transparent")
        right_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 標題
        t3_setting_lbl = ctk.CTkLabel(right_pane, text="⚙️ 執行與拆分設定", 
                                       font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        t3_setting_lbl.pack(anchor="w", pady=(0, 10))
        
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
        
        self.btn_run = ctk.CTkButton(self.btn_run_frame, text="🚀 開始拆分", 
                                         fg_color=("#2563EB", "#3B82F6"),
                                         text_color="white",
                                         hover_color=("#1D4ED8", "#2563EB"),
                                         font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
                                         command=self._start_split,
                                         height=44)
        self.btn_run.pack(fill=tk.BOTH, expand=True)
        
        self.btn_cancel = ctk.CTkButton(self.btn_run_frame, text="⛔ 終止作業", 
                                            fg_color=("#DC2626", "#EF4444"),
                                            text_color="white",
                                            hover_color=("#B91C1C", "#DC2626"),
                                            font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
                                            command=self._cancel_split,
                                            height=44)
                                            
        ctk.CTkFrame(right_pane, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 滾動設定區
        t3_scroll_settings = ctk.CTkScrollableFrame(right_pane, fg_color="transparent", label_text="")
        t3_scroll_settings.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        lbl_mode = ctk.CTkLabel(t3_scroll_settings, text="拆分擷取模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_mode.pack(anchor="w", pady=(4, 1))
        
        self.split_mode_var = tk.StringVar(value="each")
        self.radio1 = ctk.CTkRadioButton(t3_scroll_settings, text="每頁單獨拆分一個檔", 
                                            variable=self.split_mode_var, value="each",
                                            command=self._toggle_mode,
                                            font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.radio1.pack(anchor="w", pady=4)
        
        self.radio2 = ctk.CTkRadioButton(t3_scroll_settings, text="擷取指定頁碼合併", 
                                            variable=self.split_mode_var, value="extract",
                                            command=self._toggle_mode,
                                            font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.radio2.pack(anchor="w", pady=4)
        
        self.range_frame = ctk.CTkFrame(t3_scroll_settings, fg_color="transparent")
        self.range_frame.pack(fill=tk.X, pady=(4, 8))
        
        lbl_range = ctk.CTkLabel(self.range_frame, text="指定頁碼範圍:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_range.pack(anchor="w", pady=(0, 2))
        
        self.entry_range = ctk.CTkEntry(self.range_frame, placeholder_text="例: 1-5, 8, 11-14", height=32, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_range.pack(fill=tk.X)

        # 3. 命名模板
        lbl_naming = ctk.CTkLabel(t3_scroll_settings, text="輸出檔案命名模板:", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_naming.pack(anchor="w", pady=(4, 1))
        self.entry_naming = ctk.CTkEntry(t3_scroll_settings, placeholder_text="預設: {pdf_name}_page_{page_03d}", height=32, font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.entry_naming.pack(fill=tk.X, pady=(0, 8))
        self.entry_naming.insert(0, "{pdf_name}_page_{page_03d}")
        
        # 4. 輸出目錄模式
        lbl_out = ctk.CTkLabel(t3_scroll_settings, text="輸出目錄模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_out.pack(anchor="w", pady=(4, 1))
        
        self.out_mode_var = tk.StringVar(value="folder")
        out_radio_f = ctk.CTkFrame(t3_scroll_settings, fg_color="transparent")
        out_radio_f.pack(fill=tk.X, pady=(0, 10))
        
        self.out_radio1 = ctk.CTkRadioButton(out_radio_f, text="建立獨立子資料夾", 
                                                variable=self.out_mode_var, value="folder",
                                                font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.out_radio1.pack(side=tk.LEFT, padx=(0, 10))
        
        self.out_radio2 = ctk.CTkRadioButton(out_radio_f, text="同層輸出", 
                                                variable=self.out_mode_var, value="same",
                                                font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        self.out_radio2.pack(side=tk.LEFT)
        
        ctk.CTkFrame(t3_scroll_settings, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 5. PDF 高級功能設定 (加密、元資料、浮水印) - 重用元件
        self.pdf_features = PDFFeaturesFrame(t3_scroll_settings)
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
        if self.split_mode_var.get() == "extract":
            self.entry_range.configure(state="normal")
            if hasattr(self, "entry_naming") and self.entry_naming.get().strip() == "{pdf_name}_page_{page_03d}":
                self.entry_naming.delete(0, tk.END)
                self.entry_naming.insert(0, "{pdf_name}_extracted")
        else:
            self.entry_range.delete(0, tk.END)
            self.entry_range.configure(state="disabled")
            if hasattr(self, "entry_naming") and self.entry_naming.get().strip() == "{pdf_name}_extracted":
                self.entry_naming.delete(0, tk.END)
                self.entry_naming.insert(0, "{pdf_name}_page_{page_03d}")

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
            fname = os.path.basename(f)
            pages_text = "讀取中..."
            status_text = "等待中"
            
            try:
                with fitz.open(f) as doc:
                    if doc.is_encrypted:
                        pages_text = "🔒 已加密"
                        status_text = "需要密碼"
                    else:
                        pages_text = str(len(doc))
            except:
                pages_text = "錯誤"
                status_text = "檔案損毀或無效"
                
            self.tree.insert("", tk.END, values=(idx + 1, f, pages_text, status_text))
            
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
        self.app.queue.put(("t3_log", msg))

    def update_log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def _toggle_log_view(self):
        if self.show_log_var.get():
            self.log_text.pack(fill=tk.X, expand=False, pady=(0, 5))
        else:
            self.log_text.pack_forget()

    def _start_split(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "請先選擇 PDF 檔案。")
            return
            
        pdf_settings = self.pdf_features.get_settings()
        if pdf_settings["encrypt"] and not pdf_settings["password"]:
            messagebox.showwarning("警告", "您啟用了加密，但尚未設定開啟密碼。")
            return
            
        mode = self.split_mode_var.get()
        range_str = self.entry_range.get()
        
        settings = {
            "mode": mode,
            "range": range_str,
            "open": self.auto_open_var.get(),
            "naming": self.entry_naming.get(),
            "out_mode": self.out_mode_var.get(),
            "pdf_settings": pdf_settings
        }
        
        self.is_converting = True
        self.btn_run.pack_forget()
        self.btn_cancel.pack(fill=tk.BOTH, expand=True)
        self._toggle_ui_state("disabled")
        
        self.progress.set(0)
        self.log_text.delete("1.0", tk.END)
        self.update_log("===============================")
        self.update_log("🚀 PDF 拆分與擷取作業開始...")
        self.stop_event.clear()
        
        threading.Thread(target=self._split_worker, args=(settings,), daemon=True).start()

    def _cancel_split(self):
        if messagebox.askyesno("取消", "確定要終止目前的拆分作業？"):
            self.stop_event.set()
            self.btn_cancel.configure(state="disabled")
            self._log("🛑 正在停止作業，請稍後...")

    def _toggle_ui_state(self, state):
        mode = "normal" if state == "normal" else "disabled"
        self.btn_add.configure(state=mode)
        self.btn_remove.configure(state=mode)
        self.btn_clear.configure(state=mode)
        self.radio1.configure(state=mode)
        self.radio2.configure(state=mode)
        self.check_open.configure(state=mode)
        self.entry_naming.configure(state=mode)
        self.out_radio1.configure(state=mode)
        self.out_radio2.configure(state=mode)
        self.check_show_log.configure(state=mode)
        
        # 連動 PDF 高級設定元件
        self.pdf_features.configure_state(state)
        
        if mode == "normal":
            self._toggle_mode()
        else:
            self.entry_range.configure(state="disabled")

    def _split_worker(self, settings):
        """PDF 拆分背景 Worker"""
        try:
            tasks = []
            
            # 第一階段：解析與密碼解鎖
            for f in self.selected_files:
                if self.stop_event.is_set():
                    raise InterruptedError()
                    
                pw = None
                pages_count = 0
                
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
                                    pages_count = len(test_doc)
                                    correct = True
                                else:
                                    self._log(f"❌ 檔案「{os.path.basename(f)}」密碼解鎖失敗。")
                                    self.app.queue.put(("error_msg", "密碼不正確，略過此檔案。"))
                            if not correct:
                                continue
                        else:
                            pages_count = len(test_doc)
                except Exception as e:
                    self._log(f"❌ 無法讀取檔案：{os.path.basename(f)} ({str(e)})")
                    continue
                
                tasks.append({"path": f, "pages_count": pages_count, "pw": pw})
            
            if not tasks:
                self.app.queue.put(("t3_error", "清單中無有效檔案可供拆分！"))
                return
                
            total_tasks = len(tasks)
            self._log(f"📊 分析完成：共 {total_tasks} 個 PDF 檔案待處理")
            
            final_out_dir = None
            
            # 第二階段：執行拆分與存檔
            for idx, task in enumerate(tasks):
                if self.stop_event.is_set():
                    raise InterruptedError()
                    
                pdf_path = task["path"]
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                out_dir = os.path.dirname(pdf_path)
                
                self._log(f"📂 正在處理 PDF：{base_name}")
                
                try:
                    with fitz.open(pdf_path) as src_doc:
                        if src_doc.is_encrypted:
                            src_doc.authenticate(task["pw"])
                            
                        max_pages = len(src_doc)
                        
                        tpl = settings.get("naming", "").strip()
                        if not tpl:
                            tpl = "{pdf_name}_page_{page_03d}" if settings["mode"] == "each" else "{pdf_name}_extracted"
                            
                        # 根據輸出目錄模式決定目標資料夾
                        if settings.get("out_mode", "folder") == "folder":
                            sub_dir_name = base_name + "_split" if settings["mode"] == "each" else base_name + "_extracted"
                            out_target_dir = os.path.join(out_dir, sub_dir_name)
                            os.makedirs(out_target_dir, exist_ok=True)
                        else:
                            out_target_dir = out_dir
                            
                        final_out_dir = out_target_dir

                        if settings["mode"] == "each":
                            for p_no in range(max_pages):
                                if self.stop_event.is_set():
                                    raise InterruptedError()
                                    
                                out_fname = tpl.replace("{pdf_name}", base_name)
                                out_fname = out_fname.replace("{page_03d}", f"{p_no+1:03d}")
                                out_fname = out_fname.replace("{page}", f"{p_no+1}")
                                if not out_fname.lower().endswith(".pdf"):
                                    out_fname = f"{out_fname}.pdf"
                                
                                out_path = os.path.join(out_target_dir, out_fname)
                                out_path = unique_filename(out_path)
                                
                                new_doc = fitz.open()
                                new_doc.insert_pdf(src_doc, from_page=p_no, to_page=p_no)
                                finalize_and_save_pdf(new_doc, out_path, settings["pdf_settings"])
                                new_doc.close()
                                import time
                                time.sleep(0.015)  # 釋放 GIL 與 UI 更新喘息時間，防止介面凍結
                                
                            self._log(f"  ➜ 成功拆分 {max_pages} 個單頁 PDF 儲存於 {os.path.basename(out_target_dir)} 目錄")
                            
                        else:
                            select_pages = parse_range_string(settings["range"], max_pages)
                            if not select_pages:
                                self._log(f"  ⚠️ 未匹配到有效頁面，跳過該檔案")
                                continue
                                
                            out_fname = tpl.replace("{pdf_name}", base_name)
                            out_fname = out_fname.replace("{page_03d}", "extracted")
                            out_fname = out_fname.replace("{page}", "extracted")
                            if not out_fname.lower().endswith(".pdf"):
                                out_fname = f"{out_fname}.pdf"
                                
                            out_path = os.path.join(out_target_dir, out_fname)
                            out_path = unique_filename(out_path)
                            
                            new_doc = fitz.open()
                            new_doc.insert_pdf(src_doc, select=select_pages)
                            finalize_and_save_pdf(new_doc, out_path, settings["pdf_settings"])
                            new_doc.close()
                            
                            self._log(f"  ➜ 成功擷取 {len(select_pages)} 個頁面，合併儲存至：{os.path.basename(out_path)}")
                            
                except Exception as e:
                    self._log(f"❌ 檔案處理出錯：{base_name} ({str(e)})")
                    
                self.app.queue.put(("t3_progress", (idx + 1) / total_tasks))
                
            self.app.queue.put(("t3_done", final_out_dir))
            
        except InterruptedError:
            self.app.queue.put(("t3_cancelled", None))
        except Exception as e:
            self.app.queue.put(("t3_error", str(e)))
