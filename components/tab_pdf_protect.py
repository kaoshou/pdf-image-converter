import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import fitz
from tkinterdnd2 import DND_FILES

from utils.helpers import SYSTEM_FONT, FONT_OFFSET, parse_dropped_files, unique_filename
from utils.icons import get_icon

class TabPDFProtect(ctk.CTkFrame):
    """PDF 加密、解除加密與權限限制的功能分頁"""
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
        t5_main = ctk.CTkFrame(self, fg_color="transparent")
        t5_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：PDF 檔案清單
        left_pane = ctk.CTkFrame(t5_main, fg_color="transparent")
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        lbl_list = ctk.CTkLabel(left_pane, text="待處理 PDF 檔案清單:", font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        lbl_list.pack(anchor="w", pady=(0, 5))
        
        # Treeview 清單 (加強樣式，有虛擬滾動條與框線)
        tree_frame = ctk.CTkFrame(left_pane, border_width=1, border_color=("#E5E7EB", "#374151"))
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(
            tree_frame, 
            columns=("no", "name", "pages", "status"), 
            show="headings", 
            selectmode="extended",
            style="T5.Treeview"
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 設定表頭與寬度
        self.tree.heading("no", text="序號")
        self.tree.heading("name", text="檔案名稱")
        self.tree.heading("pages", text="頁數")
        self.tree.heading("status", text="防護狀態")
        
        self.tree.column("no", width=60, anchor="center")
        self.tree.column("name", width=300, anchor="w")
        self.tree.column("pages", width=80, anchor="center")
        self.tree.column("status", width=120, anchor="center")
        
        # 綁定拖放接收 (DND)
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self._handle_drop)
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
        
        # 下方檔案操作按鈕
        btn_frame = ctk.CTkFrame(left_pane, fg_color="transparent")
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 新增檔案按鈕 (使用 plus 圖示)
        plus_icon = get_icon("plus", size=(14, 14))
        self.btn_add = ctk.CTkButton(
            btn_frame, text="  新增檔案", 
            image=plus_icon,
            compound="left",
            command=self._add_files, 
            width=100, height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.btn_add.pack(side=tk.LEFT, padx=(0, 10))
        
        # 移除所選按鈕 (使用 minus 圖示)
        minus_icon = get_icon("minus", size=(14, 14))
        self.btn_remove = ctk.CTkButton(
            btn_frame, text="  移除所選", 
            image=minus_icon,
            compound="left",
            command=self._remove_selected, 
            width=100, height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.btn_remove.pack(side=tk.LEFT, padx=10)
        
        # 清空清單按鈕 (使用 trash 圖示)
        trash_icon = get_icon("trash", size=(14, 14))
        self.btn_clear = ctk.CTkButton(
            btn_frame, text="  清空清單", 
            image=trash_icon,
            compound="left",
            command=self._clear_all, 
            width=100, height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.btn_clear.pack(side=tk.LEFT, padx=10)
        
        # 右側：執行按鈕與參數設定
        right_pane_outer = ctk.CTkFrame(
            t5_main, width=340, 
            fg_color=("#FFFFFF", "#1F2937"),
            border_width=1, border_color=("#E5E7EB", "#374151")
        )
        right_pane_outer.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        right_pane_outer.pack_propagate(False)
        
        right_pane = ctk.CTkFrame(right_pane_outer, fg_color="transparent")
        right_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 標題 (使用齒輪圖示，移除 Emoji)
        settings_icon = get_icon("settings", size=(16, 16))
        t5_setting_lbl = ctk.CTkLabel(right_pane, text="  執行與安全設定", 
                                       image=settings_icon,
                                       compound="left",
                                       font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"))
        t5_setting_lbl.pack(anchor="w", pady=(0, 10))
        
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
        
        # 開始處理按鈕 (使用播放圖示，移除 Emoji)
        run_icon = get_icon("run", size=(16, 16))
        self.btn_run = ctk.CTkButton(
            self.btn_run_frame, text="  開始處理", 
            image=run_icon,
            compound="left",
            fg_color=("#2563EB", "#3B82F6"),
            text_color="white",
            hover_color=("#1D4ED8", "#2563EB"),
            font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
            command=self._start_protect,
            height=44
        )
        self.btn_run.pack(fill=tk.BOTH, expand=True)
        
        # 終止作業按鈕 (使用終止圖示，移除 Emoji)
        cancel_icon = get_icon("cancel", size=(16, 16))
        self.btn_cancel = ctk.CTkButton(
            self.btn_run_frame, text="  終止作業", 
            image=cancel_icon,
            compound="left",
            fg_color=("#DC2626", "#EF4444"),
            text_color="white",
            hover_color=("#B91C1C", "#DC2626"),
            font=(SYSTEM_FONT, 12 + FONT_OFFSET, "bold"),
            command=self._cancel_protect,
            height=44
        )
        
        ctk.CTkFrame(right_pane, height=1, fg_color=("#E5E7EB", "#374151")).pack(fill=tk.X, pady=8)
        
        # 滾動設定區
        t5_scroll_settings = ctk.CTkScrollableFrame(right_pane, fg_color="transparent", label_text="")
        t5_scroll_settings.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 1. 模式選擇
        lbl_mode = ctk.CTkLabel(t5_scroll_settings, text="防護處理模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_mode.pack(anchor="w", pady=(4, 1))
        
        # 模式單選鈕：設定加密防護 (已去除 Emoji)
        self.protect_mode_var = tk.StringVar(value="encrypt")
        self.radio_enc = ctk.CTkRadioButton(
            t5_scroll_settings, text="設定加密與權限防護", 
            variable=self.protect_mode_var, value="encrypt",
            command=self._toggle_mode,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.radio_enc.pack(anchor="w", pady=4)
        
        # 模式單選鈕：解除加密防護 (已去除 Emoji)
        self.radio_dec = ctk.CTkRadioButton(
            t5_scroll_settings, text="解除密碼防護 (無密碼)", 
            variable=self.protect_mode_var, value="decrypt",
            command=self._toggle_mode,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.radio_dec.pack(anchor="w", pady=4)
        
        # 2. 密碼參數面板
        self.params_frame = ctk.CTkFrame(t5_scroll_settings, fg_color="transparent")
        self.params_frame.pack(fill=tk.X, pady=5)
        
        # 2.1 開啟密碼 (User Password)
        lbl_upw = ctk.CTkLabel(self.params_frame, text="設定開啟密碼 (User PW):", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_upw.pack(anchor="w", pady=(2, 1))
        self.entry_user_pw = ctk.CTkEntry(
            self.params_frame, 
            placeholder_text="開啟此 PDF 所需的密碼 (必填)", 
            show="●", height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.entry_user_pw.pack(fill=tk.X, pady=(0, 8))
        
        # 2.2 權限密碼 (Owner Password)
        lbl_opw = ctk.CTkLabel(self.params_frame, text="設定編輯權限密碼 (Owner PW):", font=(SYSTEM_FONT, 10 + FONT_OFFSET))
        lbl_opw.pack(anchor="w", pady=(2, 1))
        self.entry_owner_pw = ctk.CTkEntry(
            self.params_frame, 
            placeholder_text="若留空，預設同上", 
            show="●", height=32, 
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.entry_owner_pw.pack(fill=tk.X, pady=(0, 8))
        
        # 2.3 細緻權限限制選項
        lbl_perms = ctk.CTkLabel(self.params_frame, text="細緻安全限制選項:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_perms.pack(anchor="w", pady=(4, 2))
        
        self.restrict_print_var = tk.BooleanVar(value=True)
        self.check_restrict_print = ctk.CTkCheckBox(
            self.params_frame, text="限制列印 (禁止列印)", 
            variable=self.restrict_print_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_restrict_print.pack(anchor="w", pady=3)
        
        self.restrict_copy_var = tk.BooleanVar(value=True)
        self.check_restrict_copy = ctk.CTkCheckBox(
            self.params_frame, text="限制內容複製 (禁止複製文字與影像)", 
            variable=self.restrict_copy_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_restrict_copy.pack(anchor="w", pady=3)
        
        self.restrict_edit_var = tk.BooleanVar(value=True)
        self.check_restrict_edit = ctk.CTkCheckBox(
            self.params_frame, text="限制編輯 (禁止新增註解/旋轉頁面/填表)", 
            variable=self.restrict_edit_var,
            font=(SYSTEM_FONT, 10 + FONT_OFFSET)
        )
        self.check_restrict_edit.pack(anchor="w", pady=3)
        
        # 3. 輸出目錄模式
        lbl_out = ctk.CTkLabel(t5_scroll_settings, text="輸出目錄模式:", font=(SYSTEM_FONT, 10 + FONT_OFFSET, "bold"))
        lbl_out.pack(anchor="w", pady=(8, 1))
        
        self.out_mode_var = tk.StringVar(value="folder")
        out_radio_f = ctk.CTkFrame(t5_scroll_settings, fg_color="transparent")
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
        
        self._toggle_mode()

    def _toggle_mode(self):
        """根據選擇的模式啟用/禁用設定參數"""
        mode = self.protect_mode_var.get()
        if mode == "encrypt":
            self.entry_user_pw.configure(state="normal")
            self.entry_owner_pw.configure(state="normal")
            self.check_restrict_print.configure(state="normal")
            self.check_restrict_copy.configure(state="normal")
            self.check_restrict_edit.configure(state="normal")
        else:
            self.entry_user_pw.delete(0, tk.END)
            self.entry_user_pw.configure(state="disabled")
            self.entry_owner_pw.delete(0, tk.END)
            self.entry_owner_pw.configure(state="disabled")
            self.check_restrict_print.configure(state="disabled")
            self.check_restrict_copy.configure(state="disabled")
            self.check_restrict_edit.configure(state="disabled")

    def _add_files(self):
        if self.is_converting:
            return
        files = filedialog.askopenfilenames(
            title="選擇待處理 PDF 檔案", 
            filetypes=[("PDF 檔案", "*.pdf")]
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
        for f in files:
            if not f.lower().endswith('.pdf'):
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

    def _clear_all(self):
        if self.is_converting or not self.selected_files:
            return
        if messagebox.askyesno("確認", "是否確定清空所選的 PDF 檔案？"):
            self.selected_files.clear()
            self._update_tree_content()

    def _update_tree_content(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for idx, f in enumerate(self.selected_files):
            fname = os.path.basename(f)
            pages_str = "讀取中..."
            status_str = "讀取中..."
            
            try:
                with fitz.open(f) as doc:
                    pages_str = f"{len(doc)} 頁"
                    status_str = "🔒 已加密" if doc.is_encrypted else "🔓 無防護"
            except:
                pages_str = "無法讀取"
                status_str = "未知"
                
            self.tree.insert("", tk.END, values=(idx + 1, fname, pages_str, status_str))
            
        if not self.selected_files:
            self.lbl_empty_tip.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.lbl_empty_tip.place_forget()

    def _log(self, msg):
        self.app.queue.put(("t5_log", msg))

    def update_log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def _toggle_log_view(self):
        if self.show_log_var.get():
            self.log_text.pack(fill=tk.X, expand=False, pady=(0, 5))
        else:
            self.log_text.pack_forget()

    def _start_protect(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "請先選擇 PDF 檔案。")
            return
            
        mode = self.protect_mode_var.get()
        upw = self.entry_user_pw.get().strip()
        opw = self.entry_owner_pw.get().strip()
        
        if mode == "encrypt" and not upw:
            messagebox.showwarning("警告", "您選擇了設定加密，但開啟密碼尚未填寫。")
            return
            
        settings = {
            "mode": mode,
            "user_pw": upw,
            "owner_pw": opw if opw else upw,
            "restrict_print": self.restrict_print_var.get(),
            "restrict_copy": self.restrict_copy_var.get(),
            "restrict_edit": self.restrict_edit_var.get(),
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
        self.update_log("🚀 PDF 加密防護/限制處理開始...")
        self.stop_event.clear()
        
        threading.Thread(target=self._protect_worker, args=(settings,), daemon=True).start()

    def _cancel_protect(self):
        if messagebox.askyesno("取消", "確定要終止目前的防護處理作業？"):
            self.stop_event.set()
            self.btn_cancel.configure(state="disabled")
            self._log("🛑 正在停止作業，請稍後...")

    def _toggle_ui_state(self, state):
        mode = "normal" if state == "normal" else "disabled"
        self.btn_add.configure(state=mode)
        self.btn_remove.configure(state=mode)
        self.btn_clear.configure(state=mode)
        self.radio_enc.configure(state=mode)
        self.radio_dec.configure(state=mode)
        self.check_open.configure(state=mode)
        self.out_radio1.configure(state=mode)
        self.out_radio2.configure(state=mode)
        self.check_show_log.configure(state=mode)
        
        if mode == "normal":
            self._toggle_mode()
        else:
            self.entry_user_pw.configure(state="disabled")
            self.entry_owner_pw.configure(state="disabled")
            self.check_restrict_print.configure(state="disabled")
            self.check_restrict_copy.configure(state="disabled")
            self.check_restrict_edit.configure(state="disabled")

    def _protect_worker(self, settings):
        """PDF 加密/解密背景 Worker"""
        try:
            tasks = []
            
            # 第一階段：解析與密碼解鎖
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
                                    self._log(f"❌ 檔案「{os.path.basename(f)}」密碼認證失敗")
                                    self.app.queue.put(("error_msg", "密碼不正確，略過此檔案。"))
                                    
                            if not correct:
                                continue
                except Exception as e:
                    self._log(f"❌ 無法讀取檔案：{os.path.basename(f)} ({str(e)})")
                    continue
                    
                tasks.append({"path": f, "pw": pw})
                
            if not tasks:
                self._log("⚠️ 無有效檔案可處理。")
                self.app.queue.put(("t5_done", None))
                return
                
            # 第二階段：執行保護/解密處理
            total_tasks = len(tasks)
            final_out = None
            
            for idx, task in enumerate(tasks):
                if self.stop_event.is_set():
                    raise InterruptedError()
                    
                pdf_path = task["path"]
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                out_dir = os.path.dirname(pdf_path)
                
                # 計算輸出路徑
                if settings["out_mode"] == "folder":
                    out_target_dir = os.path.join(out_dir, f"{base_name}_protected")
                    os.makedirs(out_target_dir, exist_ok=True)
                else:
                    out_target_dir = out_dir
                    
                final_out = out_target_dir
                
                # 設定新檔名
                if settings["mode"] == "encrypt":
                    out_fname = f"{base_name}_encrypted.pdf"
                else:
                    out_fname = f"{base_name}_decrypted.pdf"
                    
                out_path = os.path.join(out_target_dir, out_fname)
                out_path = unique_filename(out_path)
                
                self._log(f"⚡ 正在處理 ({(idx+1)}/{total_tasks})：{os.path.basename(pdf_path)}")
                
                try:
                    doc = fitz.open(pdf_path)
                    # 如果原檔有密碼，先進行認證
                    if doc.is_encrypted:
                        doc.authenticate(task["pw"])
                        
                    if settings["mode"] == "encrypt":
                        # 🔒 模式：設定加密與限制
                        # 權限遮罩計算 (預設允許輔助功能)
                        perm = fitz.PDF_PERM_ACCESSIBILITY
                        
                        if not settings["restrict_print"]:
                            perm |= (fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ)
                        if not settings["restrict_copy"]:
                            perm |= fitz.PDF_PERM_COPY
                        if not settings["restrict_edit"]:
                            perm |= (fitz.PDF_PERM_MODIFY | fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_FORM | fitz.PDF_PERM_ASSEMBLE)
                            
                        # 使用 AES-256 加密保存
                        doc.save(
                            out_path, 
                            encryption=fitz.PDF_ENCRYPT_AES_256,
                            user_pw=settings["user_pw"],
                            owner_pw=settings["owner_pw"],
                            permissions=int(perm),
                            garbage=4,
                            deflate=True
                        )
                        self._log(f"  ➜ 成功設定加密防護，儲存至：{os.path.basename(out_path)}")
                    else:
                        # 🔓 模式：解除加密防護
                        doc.save(out_path, garbage=4, deflate=True)
                        self._log(f"  ➜ 成功解除密碼防護，另存為：{os.path.basename(out_path)}")
                        
                    doc.close()
                except Exception as e:
                    self._log(f"❌ 檔案「{os.path.basename(pdf_path)}」處理失敗：{str(e)}")
                    
                self.app.queue.put(("t5_progress", (idx + 1) / total_tasks))
                time.sleep(0.015)  # 釋放 GIL，維持 UI 流暢
                
            self.app.queue.put(("t5_done", final_out))
            
        except InterruptedError:
            self.app.queue.put(("t5_cancelled", None))
        except Exception as e:
            self.app.queue.put(("t5_error", str(e)))
