import os
import sys
import threading
import queue
import webbrowser  # 用於開啟瀏覽器
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image

import pypdfium2 as pdfium

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except Exception:
    TkinterDnD = tk.Tk
    DND_FILES = None
    DND_AVAILABLE = False

APP_TITLE = "PDF轉圖片小工具"

# ================== 🎨 現代模組化配色 (緊湊版) ==================
COLORS = {
    "bg": "#E5E7EB",          # 背景灰
    "card_bg": "#FFFFFF",     # 卡片白
    "header_bg": "#FFFFFF",   # 頂部白
    "primary": "#2563EB",     # 皇家藍
    "primary_hover": "#1D4ED8",
    "danger": "#DC2626",      # 警告紅
    "text_main": "#1F2937",   # 深灰黑
    "text_sub": "#6B7280",    # 淺灰 (也是 Placeholder 顏色)
    "border": "#D1D5DB",      # 邊框
    "input_bg": "#F9FAFB",    # 輸入框淺底
    "accent": "#3B82F6"       # 裝飾色條
}

def get_base_dir():
    """取得程式執行基底路徑 (修正支援 PyInstaller --onefile)"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.abspath(os.path.dirname(__file__))

# ================== 🔐 密碼視窗 ==================
class CleanPasswordDialog(tk.Toplevel):
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.password = None
        self.title("安全性驗證")
        self.geometry("420x220") 
        self.resizable(False, False)
        self.configure(bg=COLORS["card_bg"])
        
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 210
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 110
            self.geometry(f"+{x}+{y}")
        except:
            pass

        frame = tk.Frame(self, bg=COLORS["card_bg"], padx=25, pady=25)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="🔒 檔案受密碼保護", font=("Microsoft JhengHei", 12, "bold"), 
                 bg=COLORS["card_bg"], fg=COLORS["primary"]).pack(anchor="w", pady=(0, 6))
        
        tk.Label(frame, text=f"檔案「{filename}」需要密碼。\n請輸入開啟密碼：", 
                 font=("Microsoft JhengHei", 10), bg=COLORS["card_bg"], fg=COLORS["text_sub"], 
                 justify="left").pack(anchor="w", pady=(0, 10))

        self.entry = tk.Entry(frame, show="●", bg=COLORS["input_bg"], fg=COLORS["text_main"], 
                              relief="flat", font=("Helvetica", 11))
        self.entry.config(highlightthickness=1, highlightbackground=COLORS["border"], highlightcolor=COLORS["primary"])
        self.entry.pack(fill=tk.X, ipady=6, pady=(0, 20))
        self.entry.focus_set()

        btn_frame = tk.Frame(frame, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X)

        tk.Button(btn_frame, text="確認解鎖 🔓", command=self.on_submit, 
                  bg=COLORS["primary"], fg="white", bd=0, font=("Microsoft JhengHei", 10, "bold"),
                  activebackground=COLORS["primary_hover"], activeforeground="white", 
                  padx=16, pady=6, cursor="hand2").pack(side=tk.RIGHT)

        tk.Button(btn_frame, text="略過此檔", command=self.on_cancel, 
                  bg=COLORS["bg"], fg=COLORS["text_sub"], bd=0, font=("Microsoft JhengHei", 9),
                  activebackground=COLORS["border"], padx=12, pady=6, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 10))

        self.bind('<Return>', lambda e: self.on_submit())
        self.bind('<Escape>', lambda e: self.on_cancel())

    def on_submit(self):
        self.password = self.entry.get()
        self.destroy()

    def on_cancel(self):
        self.password = None
        self.destroy()

# ================== ℹ️ 關於視窗 ==================
class AboutDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("關於本程式")
        self.geometry("520x480") 
        self.resizable(False, False)
        self.configure(bg=COLORS["card_bg"])
        
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 260
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 240
            self.geometry(f"+{x}+{y}")
        except:
            pass

        frame = tk.Frame(self, bg=COLORS["card_bg"], padx=30, pady=30)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text=APP_TITLE, font=("Microsoft JhengHei", 14, "bold"), 
                 bg=COLORS["card_bg"], fg=COLORS["text_main"]).pack(anchor="w", pady=(0, 5))
        
        tk.Label(frame, text="版本: 0.2", font=("Microsoft JhengHei", 10), 
                 bg=COLORS["card_bg"], fg=COLORS["text_sub"]).pack(anchor="w", pady=(0, 15))

        self._create_link_row(frame, "本程式原始碼 (GitHub):", "https://github.com/kaoshou/pdf-image-converter")

        tk.Label(frame, text="開發者: 鄭郁翰 (Cheng, Yu-Han)", font=("Microsoft JhengHei", 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_main"]).pack(anchor="w", pady=(15, 2))
        tk.Label(frame, text="Email: kaoshou@gmail.com", font=("Microsoft JhengHei", 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_main"]).pack(anchor="w", pady=(0, 15))

        tk.Frame(frame, bg=COLORS["border"], height=1).pack(fill=tk.X, pady=10)

        tk.Label(frame, text="第三方套件授權:", font=("Microsoft JhengHei", 10, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_main"]).pack(anchor="w", pady=(0, 10))
        
        self._create_link_row(frame, "• pypdfium2 (Apache/BSD/MIT)", "https://github.com/pypdfium2-team/pypdfium2")
        self._create_link_row(frame, "• Pillow (HPKSA/MIT License)", "https://github.com/python-pillow/Pillow")
        self._create_link_row(frame, "• tkinterdnd2 (MIT License)", "https://github.com/pmgagne/tkinterdnd2")

        tk.Button(frame, text="關閉視窗", command=self.destroy,
                  bg=COLORS["bg"], fg=COLORS["text_main"], bd=0, 
                  font=("Microsoft JhengHei", 9), padx=20, pady=8,
                  activebackground=COLORS["border"], cursor="hand2").pack(side=tk.BOTTOM, pady=(20, 0))

    def _create_link_row(self, parent, label_text, url):
        row = tk.Frame(parent, bg=COLORS["card_bg"])
        row.pack(fill=tk.X, pady=2)
        
        tk.Label(row, text=label_text, font=("Microsoft JhengHei", 9),
                 bg=COLORS["card_bg"], fg=COLORS["text_main"]).pack(side=tk.LEFT)
        
        link = tk.Label(row, text=url, font=("Microsoft JhengHei", 9, "underline"),
                        bg=COLORS["card_bg"], fg=COLORS["primary"], cursor="hand2")
        link.pack(side=tk.LEFT, padx=(5, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(url))

# ================== 主程式 ==================
class PDFImageConverter:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("850x620") 
        self.root.minsize(750, 550)
        self.root.configure(bg=COLORS["bg"])

        self.base_dir = get_base_dir()
        self.queue = queue.Queue()
        self.stop_event = threading.Event()

        self.selected_files = []
        self.auto_open_var = tk.BooleanVar(value=False)
        self.rotation_var = tk.StringVar(value="0")
        
        # 修改: 初始化為空字串，以便顯示 Placeholder
        self.dpi_var = tk.StringVar(value="")
        self.page_start_var = tk.StringVar(value="")
        self.page_end_var = tk.StringVar(value="")
        
        self.output_format_var = tk.StringVar(value="PNG")
        self.output_mode_var = tk.StringVar(value="folder")
        self.file_summary_var = tk.StringVar(value="尚未選擇檔案")

        # 定義 Placeholder 文字 (用於後續比對)
        self.PH_DPI = "預設: 200"
        self.PH_START = "預設: 1"
        self.PH_END = "預設: 最末頁"

        self._setup_style()
        self._build_ui()

        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)
            
        self.root.after(100, self.process_queue)
        # 修改: 移除啟動時的 "pypdfium2 核心已載入" 訊息

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card_bg"], relief="flat")
        
        style.configure("TLabel", background=COLORS["card_bg"], foreground=COLORS["text_main"], font=("Microsoft JhengHei", 9))
        style.configure("Header.TLabel", background=COLORS["header_bg"], foreground=COLORS["text_main"], font=("Microsoft JhengHei", 16, "bold"))
        
        style.configure("SectionTitle.TLabel", background=COLORS["card_bg"], foreground=COLORS["text_main"], font=("Microsoft JhengHei", 11, "bold"))
        
        style.configure("Primary.TButton",
            font=("Microsoft JhengHei", 10, "bold"),
            background=COLORS["primary"], foreground="#FFFFFF",
            borderwidth=0, padding=(16, 6)
        )
        style.map("Primary.TButton",
            background=[("active", COLORS["primary_hover"]), ("disabled", "#9CA3AF")],
            foreground=[("disabled", "#F3F4F6")]
        )

        style.configure("Danger.TButton",
            font=("Microsoft JhengHei", 10, "bold"),
            background=COLORS["danger"], foreground="#FFFFFF",
            borderwidth=0, padding=(16, 6)
        )
        style.map("Danger.TButton", background=[("active", "#B91C1C")])

        style.configure("Secondary.TButton",
            font=("Microsoft JhengHei", 9),
            background=COLORS["input_bg"], foreground=COLORS["text_main"],
            borderwidth=1, bordercolor=COLORS["border"], padding=(12, 4)
        )
        style.map("Secondary.TButton",
            background=[("active", "#E5E7EB")]
        )

        style.configure("TCombobox", fieldbackground=COLORS["input_bg"], arrowcolor=COLORS["text_sub"])
        style.configure("TRadiobutton", background=COLORS["card_bg"], font=("Microsoft JhengHei", 9), foreground=COLORS["text_main"])
        style.configure("TCheckbutton", background=COLORS["card_bg"], font=("Microsoft JhengHei", 9), foreground=COLORS["text_main"])

    def _build_ui(self):
        header_frame = tk.Frame(self.root, bg=COLORS["header_bg"], height=50, padx=20)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        tk.Frame(self.root, bg=COLORS["border"], height=1).pack(fill=tk.X)

        left = tk.Frame(header_frame, bg=COLORS["header_bg"])
        left.pack(side=tk.LEFT, fill=tk.Y)
        tk.Frame(left, bg=COLORS["primary"], width=4).pack(side=tk.LEFT, fill=tk.Y, pady=12)
        ttk.Label(left, text=f"  {APP_TITLE}", style="Header.TLabel", background=COLORS["header_bg"]).pack(side=tk.LEFT, pady=10)

        tk.Button(header_frame, text="關於本程式", command=self.show_about,
                  bg=COLORS["header_bg"], fg=COLORS["text_sub"], bd=0, 
                  font=("Microsoft JhengHei", 9), cursor="hand2", activebackground=COLORS["bg"]).pack(side=tk.RIGHT)

        main_area = tk.Frame(self.root, bg=COLORS["bg"], padx=16, pady=16)
        main_area.pack(fill=tk.BOTH, expand=True)

        self._build_file_card(main_area)
        tk.Frame(main_area, bg=COLORS["bg"], height=10).pack(fill=tk.X)

        self._build_settings_card(main_area)
        tk.Frame(main_area, bg=COLORS["bg"], height=10).pack(fill=tk.X)

        self._build_action_card(main_area)

    def _create_card_frame(self, parent):
        card = tk.Frame(parent, bg=COLORS["card_bg"], padx=20, pady=15)
        card.pack(fill=tk.X)
        card.config(highlightbackground=COLORS["border"], highlightthickness=1)
        return card

    def _build_section_header(self, parent, text):
        row = tk.Frame(parent, bg=COLORS["card_bg"])
        row.pack(fill=tk.X, pady=(0, 10))
        tk.Frame(row, bg=COLORS["accent"], width=3, height=16).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row, text=text, style="SectionTitle.TLabel").pack(side=tk.LEFT)

    def _build_file_card(self, parent):
        card = self._create_card_frame(parent)
        self._build_section_header(card, "檔案來源")

        row = tk.Frame(card, bg=COLORS["card_bg"])
        row.pack(fill=tk.X)
        
        ttk.Button(row, text="選擇 PDF 檔案...", style="Secondary.TButton", command=self.select_pdfs).pack(side=tk.LEFT)
        tk.Label(row, textvariable=self.file_summary_var, font=("Microsoft JhengHei", 9), 
                 bg=COLORS["card_bg"], fg=COLORS["primary"]).pack(side=tk.LEFT, padx=(12, 0))

    def _build_settings_card(self, parent):
        card = self._create_card_frame(parent)
        self._build_section_header(card, "轉檔參數")

        grid = tk.Frame(card, bg=COLORS["card_bg"])
        grid.pack(fill=tk.X)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        # 修改: 傳入 placeholder 文字
        self._make_input(grid, 0, 0, "📄 起始頁碼", self.page_start_var, placeholder=self.PH_START)
        self._make_input(grid, 0, 2, "🔄 畫面旋轉", self.rotation_var, is_combo=True, values=["0", "90", "180", "270"])
        
        self._make_input(grid, 1, 0, "📄 結束頁碼", self.page_end_var, placeholder=self.PH_END)
        self._make_input(grid, 1, 2, "🎨 圖片格式", self.output_format_var, is_combo=True, values=["PNG", "JPG"])

        self._make_input(grid, 2, 0, "🔍 解析度 (DPI)", self.dpi_var, placeholder=self.PH_DPI)
        
        mode_f = tk.Frame(grid, bg=COLORS["card_bg"])
        mode_f.grid(row=2, column=2, columnspan=2, sticky="w", padx=10, pady=4)
        tk.Label(mode_f, text="📂 輸出位置：", bg=COLORS["card_bg"], font=("Microsoft JhengHei", 9)).pack(side=tk.LEFT)
        ttk.Radiobutton(mode_f, text="建立資料夾", variable=self.output_mode_var, value="folder").pack(side=tk.LEFT, padx=6)
        ttk.Radiobutton(mode_f, text="同層目錄", variable=self.output_mode_var, value="same").pack(side=tk.LEFT)

    def _build_action_card(self, parent):
        card = self._create_card_frame(parent)
        self._build_section_header(card, "執行作業")

        act_row = tk.Frame(card, bg=COLORS["card_bg"])
        act_row.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(act_row, text="完成後開啟資料夾", variable=self.auto_open_var).pack(side=tk.LEFT)

        self.btn_container = tk.Frame(act_row, bg=COLORS["card_bg"])
        self.btn_container.pack(side=tk.RIGHT)

        self.convert_btn = ttk.Button(self.btn_container, text="🚀 開始轉檔", style="Primary.TButton", command=self.start_convert)
        self.convert_btn.pack(side=tk.RIGHT)

        self.cancel_btn = ttk.Button(self.btn_container, text="⛔ 終止作業", style="Danger.TButton", command=self.cancel_convert)
        
        self.progress = ttk.Progressbar(card, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 8))

        log_box = tk.Frame(card, bg=COLORS["input_bg"], bd=1, relief="solid")
        log_box.config(highlightthickness=0)
        log_box.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_box, height=5, bg="#FAFAFA", fg="#374151", 
                                font=("Consolas", 9), relief="flat", padx=8, pady=8)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scr = ttk.Scrollbar(log_box, command=self.log_text.yview)
        scr.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scr.set)

    def _make_input(self, parent, r, c, label, var, is_combo=False, values=None, placeholder=None):
        pady_val = 4
        tk.Label(parent, text=label, bg=COLORS["card_bg"], font=("Microsoft JhengHei", 9)).grid(row=r, column=c, sticky="w", padx=(0, 8), pady=pady_val)
        
        f = tk.Frame(parent, bg=COLORS["card_bg"])
        f.grid(row=r, column=c+1, sticky="ew", padx=(0, 24), pady=pady_val)
        
        if is_combo:
            cb = ttk.Combobox(f, textvariable=var, values=values, width=10, state="readonly")
            cb.pack(fill=tk.X)
        else:
            e = tk.Entry(f, textvariable=var, width=12, relief="flat", bg=COLORS["input_bg"])
            e.config(highlightbackground=COLORS["border"], highlightthickness=1)
            e.pack(fill=tk.X, ipady=4, ipadx=4)

            # 修改: 實作 Placeholder 邏輯
            if placeholder:
                def on_focus_in(event):
                    if var.get() == placeholder:
                        var.set("")
                        e.config(fg=COLORS["text_main"])
                
                def on_focus_out(event):
                    if var.get() == "":
                        var.set(placeholder)
                        e.config(fg=COLORS["text_sub"])

                e.bind("<FocusIn>", on_focus_in)
                e.bind("<FocusOut>", on_focus_out)

                # 初始化狀態
                if not var.get():
                    var.set(placeholder)
                    e.config(fg=COLORS["text_sub"])
                elif var.get() == placeholder:
                    e.config(fg=COLORS["text_sub"])
                else:
                    e.config(fg=COLORS["text_main"])

    def select_pdfs(self):
        files = filedialog.askopenfilenames(title="選擇 PDF", filetypes=[("PDF", "*.pdf")])
        if files:
            self.selected_files = list(files)
            count = len(files)
            name = os.path.basename(files[0])
            msg = f"{name}" if count == 1 else f"{name} 等 {count} 個檔案"
            self.file_summary_var.set(msg)
            self.log(f"已選擇: {msg}")

    def on_drop(self, event):
        raw = event.data.strip()
        if not raw: return
        paths = []
        temp = ""
        in_brace = False
        for char in raw:
            if char == "{": in_brace = True
            elif char == "}": in_brace = False; paths.append(temp); temp = ""
            elif char == " " and not in_brace: 
                if temp: paths.append(temp); temp = ""
            else: temp += char
        if temp: paths.append(temp)
        
        pdfs = [p for p in paths if p.lower().endswith(".pdf")]
        if pdfs:
            self.selected_files = pdfs
            count = len(pdfs)
            name = os.path.basename(pdfs[0])
            msg = f"{name}" if count == 1 else f"{name} 等 {count} 個檔案"
            self.file_summary_var.set(msg)
            self.log(f"拖曳載入: {msg}")

    def log(self, msg):
        self.queue.put(("log", msg))

    def _update_log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def start_convert(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return

        # 修改: 解析輸入值時，處理 Placeholder 文字 (視為使用預設值)
        def parse_input(val_str, placeholder, default_val):
            val = val_str.strip()
            if not val or val == placeholder:
                return default_val
            try:
                return int(val)
            except:
                return default_val

        dpi = parse_input(self.dpi_var.get(), self.PH_DPI, 200)
        s = parse_input(self.page_start_var.get(), self.PH_START, None)
        e = parse_input(self.page_end_var.get(), self.PH_END, None)
        
        settings = {
            "dpi": dpi, "start": s, "end": e,
            "angle": int(self.rotation_var.get()),
            "fmt": self.output_format_var.get(),
            "mode": self.output_mode_var.get(),
            "open": self.auto_open_var.get()
        }

        self.convert_btn.pack_forget()
        self.cancel_btn.pack(side=tk.RIGHT)
        self.cancel_btn.config(state="normal")
        self.progress['value'] = 0
        self.log("===============================")
        self.log("🚀 轉檔作業開始...")
        self.stop_event.clear()

        threading.Thread(target=self.worker, args=(settings,), daemon=True).start()

    def cancel_convert(self):
        if messagebox.askyesno("取消", "確定要停止目前作業？"):
            self.stop_event.set()
            self.cancel_btn.config(state="disabled")
            self.log("🛑 正在停止...")

    def worker(self, settings):
        try:
            tasks = []
            total_pages = 0
            
            for f in self.selected_files:
                if self.stop_event.is_set(): raise InterruptedError()
                info = self._get_pdf_info(f)
                if not info: continue
                
                p_total = info["Pages"]
                s = settings["start"] or 1
                e = settings["end"] or p_total
                e = min(e, p_total)
                if s > e: continue
                pages = list(range(s, e+1))
                tasks.append({"path": f, "pages": pages, "pw": info.get("_pw")})
                total_pages += len(pages)

            if total_pages == 0:
                self.queue.put(("error", "無頁面可轉換"))
                return

            self.queue.put(("set_max", total_pages))
            self.log(f"📊 分析完成：共 {total_pages} 頁待處理")

            current = 0
            scale_factor = settings["dpi"] / 72.0 

            for task in tasks:
                if self.stop_event.is_set(): raise InterruptedError()
                base = os.path.splitext(os.path.basename(task["path"]))[0]
                out_dir = os.path.dirname(task["path"])
                if settings["mode"] == "folder":
                    out_dir = os.path.join(out_dir, base + "_images")
                    os.makedirs(out_dir, exist_ok=True)
                
                self.log(f"📂 正在處理：{base}")

                try:
                    pdf = pdfium.PdfDocument(task["path"], password=task["pw"])
                    
                    for p_num in task["pages"]:
                        if self.stop_event.is_set(): raise InterruptedError()
                        
                        page_index = p_num - 1
                        page = pdf[page_index]
                        bitmap = page.render(scale=scale_factor)
                        pil_image = bitmap.to_pil()
                        
                        bitmap.close()
                        page.close()

                        if settings["angle"]: 
                            pil_image = pil_image.rotate(settings["angle"], expand=True)

                        ext = settings["fmt"].lower()
                        fname = f"page_{p_num}.{ext}"
                        save_path = self._unique_path(os.path.join(out_dir, fname))
                        
                        fmt_param = "JPEG" if ext in ["jpg", "jpeg"] else "PNG"
                        if fmt_param == "JPEG": 
                            pil_image = pil_image.convert("RGB")
                            
                        pil_image.save(save_path, fmt_param)
                        self.log(f"  ➜ 第 {p_num} 頁轉換成功")

                        current += 1
                        self.queue.put(("progress", current))
                    
                    pdf.close()

                except Exception as e:
                    self.log(f"❌ 檔案處理錯誤: {str(e)}")
                    continue
                
                if settings["open"] and not self.stop_event.is_set():
                    try: os.startfile(out_dir)
                    except: pass
            
            self.queue.put(("done", None))

        except InterruptedError:
            self.queue.put(("cancelled", None))
        except Exception as e:
            self.queue.put(("error", str(e)))

    def _get_pdf_info(self, path):
        pw = None
        for _ in range(2):
            try:
                pdf = pdfium.PdfDocument(path, password=pw)
                page_count = len(pdf)
                pdf.close()
                return {"Pages": page_count, "_pw": pw}
            except Exception as e:
                err_str = str(e).lower()
                if "password" in err_str or "incorrect" in err_str or "crypt" in err_str:
                    pw = self.ask_password_ui(path)
                    if not pw: return None
                else:
                    if pw is None:
                         pw = self.ask_password_ui(path)
                         if not pw: return None
                    else:
                        self.log(f"讀取失敗: {os.path.basename(path)} ({e})")
                        return None
        return None

    def _unique_path(self, path):
        if not os.path.exists(path): return path
        base, ext = os.path.splitext(path)
        i = 1
        while True:
            new_p = f"{base}_{i}{ext}"
            if not os.path.exists(new_p): return new_p
            i += 1

    def ask_password_ui(self, path):
        evt = threading.Event()
        res = {}
        self.queue.put(("ask_pw", (path, evt, res)))
        evt.wait()
        return res.get("pw")

    def process_queue(self):
        try:
            while True:
                kind, data = self.queue.get_nowait()
                if kind == "log": self._update_log(data)
                elif kind == "set_max":
                    self.progress['maximum'] = data
                    self.progress['value'] = 0
                elif kind == "progress": self.progress['value'] = data
                elif kind == "ask_pw":
                    path, evt, res = data
                    dialog = CleanPasswordDialog(self.root, os.path.basename(path))
                    self.root.wait_window(dialog)
                    res["pw"] = dialog.password
                    evt.set()
                elif kind in ["done", "error", "cancelled"]:
                    self.cancel_btn.pack_forget()
                    self.convert_btn.pack(side=tk.RIGHT)
                    if kind == "done":
                        self.progress['value'] = self.progress['maximum']
                        self.log("✨ 恭喜！所有轉檔作業已完成。")
                        messagebox.showinfo("完成", "所有轉檔作業已完成！")
                    elif kind == "cancelled":
                        self.log("⚠️ 作業已手動取消")
                        messagebox.showinfo("取消", "作業已取消")
                    elif kind == "error":
                        messagebox.showerror("錯誤", f"發生錯誤: {data}")
        except queue.Empty: pass
        finally: self.root.after(100, self.process_queue)

    def show_about(self):
        AboutDialog(self.root)

if __name__ == "__main__":
    if DND_AVAILABLE: root = TkinterDnD.Tk()
    else: root = tk.Tk()
    PDFImageConverter(root)
    root.mainloop()