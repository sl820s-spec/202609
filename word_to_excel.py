# -*- coding: utf-8 -*-
"""
Word表格数据提取到Excel工具
功能：将Word文档中每页一个的相同结构表格提取并汇总到Excel
"""

import os
import sys
import threading
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Entry, StringVar, IntVar, BooleanVar,
    messagebox, ttk, scrolledtext, END, N, S, E, W,
    LEFT, RIGHT, TOP, BOTTOM, X, Y, BOTH, NORMAL, DISABLED
)
from tkinter import filedialog

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
    Workbook = None


# ============ 设计变量（与 HTML 预览保持一致） ============
class Colors:
    PRIMARY      = "#4F46E5"
    PRIMARY_HOV  = "#4338CA"
    PRIMARY_LIGHT= "#EEF2FF"
    SUCCESS      = "#10B981"
    SUCCESS_HOV  = "#059669"
    WARNING      = "#F59E0B"
    WARNING_HOV  = "#D97706"
    DANGER       = "#EF4444"

    BG_PAGE      = "#F3F4F6"
    BG_CARD      = "#FFFFFF"
    BG_HOVER     = "#F9FAFB"
    BG_INPUT     = "#FFFFFF"
    BG_LOG       = "#0F172A"

    BORDER       = "#E5E7EB"
    BORDER_FOCUS = "#4F46E5"

    TXT_PRIMARY   = "#111827"
    TXT_SECONDARY = "#4B5563"
    TXT_TERTIARY  = "#9CA3AF"
    TXT_INVERSE   = "#FFFFFF"
    TXT_LOG       = "#CBD5E1"

    HEADER_START = "#4F46E5"
    HEADER_END   = "#8B5CF6"


class WordToExcelApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Word 表格数据提取工具")
        self.root.geometry("880x760")
        self.root.minsize(820, 700)
        self.root.configure(bg=Colors.BG_PAGE)

        # 变量
        self.word_path   = StringVar()
        self.excel_path  = StringVar()
        self.col_count   = IntVar(value=5)
        self.row_count   = IntVar(value=30)
        self.title_col   = IntVar(value=0)
        self.data_start  = IntVar(value=1)
        self.data_end    = IntVar(value=3)
        self.skip_first  = BooleanVar(value=False)
        self.include_title = BooleanVar(value=True)

        self._setup_style()
        self._build_ui()
        self._check_dependencies()

    # ============ 样式系统 ============
    def _setup_style(self):
        try:
            style = ttk.Style()
            style.theme_use("clam")
        except Exception:
            style = ttk.Style()

        # Treeview（预览表）
        style.configure("Card.Treeview",
                        background=Colors.BG_CARD,
                        foreground=Colors.TXT_PRIMARY,
                        fieldbackground=Colors.BG_CARD,
                        rowheight=28,
                        font=("Microsoft YaHei", 10),
                        borderwidth=0)
        style.configure("Card.Treeview.Heading",
                        background="#F8FAFC",
                        foreground=Colors.TXT_SECONDARY,
                        font=("Microsoft YaHei", 9, "bold"),
                        padding=(8, 8),
                        borderwidth=0)
        style.map("Card.Treeview",
                  background=[("selected", Colors.PRIMARY_LIGHT)],
                  foreground=[("selected", Colors.TXT_PRIMARY)])
        style.configure("Card.Treeview", rowheight=30)

        # Progressbar
        style.configure("Card.Horizontal.TProgressbar",
                        troughcolor="#E5E7EB",
                        background=Colors.PRIMARY,
                        thickness=6,
                        borderwidth=0)

        # Checkbutton
        style.configure("Card.TCheckbutton",
                        background=Colors.BG_CARD,
                        foreground=Colors.TXT_SECONDARY,
                        font=("Microsoft YaHei", 10))

        # Scrollbar
        style.configure("Vertical.TScrollbar",
                        background="#E5E7EB",
                        troughcolor=Colors.BG_CARD,
                        borderwidth=0,
                        arrowcolor=Colors.TXT_TERTIARY)

    # ============ UI 构建 ============
    def _build_ui(self):
        # ---------- 顶部标题栏 ----------
        self._build_header()

        # ---------- 主体容器 ----------
        main = Frame(self.root, bg=Colors.BG_PAGE)
        main.pack(fill=BOTH, expand=True, padx=18, pady=(14, 14))

        # 1) 文件与参数卡
        self._build_file_param_card(main)

        # 2) 预览卡
        self._build_preview_card(main)

        # 3) 日志卡
        self._build_log_card(main)

        # 4) 底部操作条
        self._build_actionbar(main)

    def _build_header(self):
        # 用 Frame 模拟渐变顶栏（分段填色近似）
        header = Frame(self.root, bg=Colors.HEADER_START, height=72)
        header.pack(fill=X)
        header.pack_propagate(False)

        # 内部容器
        inner = Frame(header, bg=Colors.HEADER_START)
        inner.pack(fill=BOTH, expand=True, padx=28, pady=14)

        # Logo 方块
        logo = Label(inner, text="📊",
                     bg=Colors.HEADER_START, fg=Colors.TXT_INVERSE,
                     font=("Segoe UI Emoji", 22),
                     width=2, height=1)
        logo.pack(side=LEFT, padx=(0, 12))

        # 标题组
        title_box = Frame(inner, bg=Colors.HEADER_START)
        title_box.pack(side=LEFT)

        Label(title_box, text="Word 表格数据提取工具",
              bg=Colors.HEADER_START, fg=Colors.TXT_INVERSE,
              font=("Microsoft YaHei", 15, "bold")).pack(anchor=W)
        Label(title_box, text="一键将多页 Word 表格汇总到 Excel",
              bg=Colors.HEADER_START, fg="#E0E7FF",
              font=("Microsoft YaHei", 9)).pack(anchor=W, pady=(1, 0))

        # 版本徽章
        badge = Label(inner, text="v1.0 · Windows",
                      bg="#6366F1", fg=Colors.TXT_INVERSE,
                      font=("Microsoft YaHei", 9),
                      padx=12, pady=4)
        badge.place(relx=1.0, x=-4, y=12, anchor="ne")

    def _make_card(self, parent):
        card = Frame(parent, bg=Colors.BG_CARD,
                     highlightthickness=1,
                     highlightbackground=Colors.BORDER)
        return card

    def _make_card_head(self, card, icon_text, title, subtitle="",
                        icon_bg=None, icon_fg=None):
        head = Frame(card, bg=Colors.BG_CARD)
        head.pack(fill=X, padx=0, pady=0)
        # 底部细线
        sep = Frame(head, bg=Colors.BORDER, height=1)
        sep.place(relx=0, rely=1.0, anchor=W, relwidth=1.0)

        ic_bg = icon_bg or Colors.PRIMARY_LIGHT
        ic_fg = icon_fg or Colors.PRIMARY
        ic = Label(head, text=icon_text, bg=ic_bg, fg=ic_fg,
                   font=("Segoe UI Emoji", 11),
                   width=2, height=1)
        ic.pack(side=LEFT, padx=(16, 8), pady=12)

        Label(head, text=title,
              bg=Colors.BG_CARD, fg=Colors.TXT_PRIMARY,
              font=("Microsoft YaHei", 11, "bold")).pack(side=LEFT, pady=12)

        if subtitle:
            Label(head, text=subtitle,
                  bg=Colors.BG_CARD, fg=Colors.TXT_TERTIARY,
                  font=("Microsoft YaHei", 9)).pack(side=LEFT, padx=6, pady=12)
        return head

    def _build_file_param_card(self, parent):
        card = self._make_card(parent)
        card.pack(fill=X, pady=(0, 12))

        self._make_card_head(card, "📁", "文件与参数", "配置输入输出路径和表格结构")

        body = Frame(card, bg=Colors.BG_CARD)
        body.pack(fill=X, padx=18, pady=16)

        # --- 文件路径 ---
        def _file_row(parent, label, var, browse_cmd):
            row = Frame(parent, bg=Colors.BG_CARD)
            row.pack(fill=X, pady=(0, 10))

            Label(row, text=label, bg=Colors.BG_CARD, fg=Colors.TXT_SECONDARY,
                  font=("Microsoft YaHei", 9)).pack(side=LEFT, padx=(0, 10))

            entry = Entry(row, textvariable=var,
                          font=("Consolas", 10),
                          bg=Colors.BG_INPUT, fg=Colors.TXT_PRIMARY,
                          relief="flat",
                          highlightthickness=1,
                          highlightbackground=Colors.BORDER,
                          highlightcolor=Colors.BORDER_FOCUS,
                          bd=0)
            entry.pack(side=LEFT, fill=X, expand=True, ipady=7)

            self._mk_ghost_btn(row, "浏览…", browse_cmd).pack(side=LEFT, padx=(8, 0))

        _file_row(body, "Word 文件  ", self.word_path,  self._pick_word)
        _file_row(body, "Excel 输出", self.excel_path, self._pick_excel)

        # 间距
        Frame(body, bg=Colors.BG_CARD, height=14).pack()

        # --- 参数网格 (4列) ---
        params = Frame(body, bg=Colors.BG_CARD)
        params.pack(fill=X)

        def _param(label, var):
            box = Frame(params, bg=Colors.BG_CARD)
            box.pack(side=LEFT, fill=X, expand=True, padx=(0, 14))
            Label(box, text=label, bg=Colors.BG_CARD, fg=Colors.TXT_TERTIARY,
                  font=("Microsoft YaHei", 8)).pack(anchor=W)
            e = Entry(box, textvariable=var,
                      font=("Consolas", 11), justify="center",
                      bg=Colors.BG_INPUT, fg=Colors.TXT_PRIMARY,
                      relief="flat", bd=0,
                      highlightthickness=1,
                      highlightbackground=Colors.BORDER,
                      highlightcolor=Colors.BORDER_FOCUS)
            e.pack(fill=X, ipady=7, pady=(4, 0))
            return box

        _param("表格列数", self.col_count)
        _param("表格行数", self.row_count)
        _param("标题列索引 (0 起)", self.title_col)

        # 数据范围（两个数 + 中间 ~）
        range_box = Frame(params, bg=Colors.BG_CARD)
        range_box.pack(side=LEFT, fill=X, expand=True)
        Label(range_box, text="数据列范围", bg=Colors.BG_CARD,
              fg=Colors.TXT_TERTIARY,
              font=("Microsoft YaHei", 8)).pack(anchor=W)
        range_inner = Frame(range_box, bg=Colors.BG_CARD)
        range_inner.pack(fill=X, pady=(4, 0))
        for v in (self.data_start, self.data_end):
            Entry(range_inner, textvariable=v,
                  font=("Consolas", 11), justify="center",
                  bg=Colors.BG_INPUT, fg=Colors.TXT_PRIMARY,
                  relief="flat", bd=0, width=5,
                  highlightthickness=1,
                  highlightbackground=Colors.BORDER,
                  highlightcolor=Colors.BORDER_FOCUS
                  ).pack(side=LEFT, fill=X, expand=True, ipady=7)
            Label(range_inner, text="~", bg=Colors.BG_CARD,
                  fg=Colors.TXT_TERTIARY,
                  font=("Microsoft YaHei", 10)).pack(side=LEFT, padx=4)

        # --- 虚线分隔 + 复选框 + 自动检测 ---
        sep = Frame(body, bg=Colors.BORDER, height=1)
        sep.pack(fill=X, pady=(16, 12))

        checks = Frame(body, bg=Colors.BG_CARD)
        checks.pack(fill=X)

        ttk.Checkbutton(checks, text="跳过每页第一行（表头）",
                        variable=self.skip_first,
                        style="Card.TCheckbutton").pack(side=LEFT, padx=(0, 24))
        ttk.Checkbutton(checks, text="导出结果中包含标题列",
                        variable=self.include_title,
                        style="Card.TCheckbutton").pack(side=LEFT)

        self._mk_outline_btn(checks, "🔍 自动检测结构",
                             self._auto_detect).pack(side=RIGHT)

    def _build_preview_card(self, parent):
        card = self._make_card(parent)
        card.pack(fill=BOTH, expand=True, pady=(0, 12))

        self._make_card_head(card, "👁", "数据预览", "显示前 10 行提取结果",
                             icon_bg="#FEF3C7", icon_fg="#B45309")

        body = Frame(card, bg=Colors.BG_CARD)
        body.pack(fill=BOTH, expand=True, padx=18, pady=(8, 14))

        # Treeview
        self.preview_tree = ttk.Treeview(body, show="headings",
                                         style="Card.Treeview", height=8)
        vsb = ttk.Scrollbar(body, orient="vertical",
                            command=self.preview_tree.yview,
                            style="Vertical.TScrollbar")
        self.preview_tree.configure(yscrollcommand=vsb.set)
        self.preview_tree.pack(side=LEFT, fill=BOTH, expand=True)
        vsb.pack(side=RIGHT, fill=Y, padx=(6, 0))

        # 示例：设置 4 列
        self.preview_tree["columns"] = ("c1", "c2", "c3", "c4")
        self.preview_tree.heading("c1", text="Col 1 · 标题列")
        self.preview_tree.heading("c2", text="Col 2")
        self.preview_tree.heading("c3", text="Col 3")
        self.preview_tree.heading("c4", text="Col 4")
        self.preview_tree.column("c1", width=140, anchor=W)
        self.preview_tree.column("c2", width=150, anchor=W)
        self.preview_tree.column("c3", width=140, anchor=W)
        self.preview_tree.column("c4", width=140, anchor=W)

        # tag 用于高亮标题列行
        self.preview_tree.tag_configure("title_row",
                                        background=Colors.PRIMARY_LIGHT,
                                        foreground=Colors.PRIMARY)

    def _build_log_card(self, parent):
        card = self._make_card(parent)
        card.pack(fill=X, pady=(0, 12))

        self._make_card_head(card, "📜", "运行日志", "实时显示处理进度",
                             icon_bg="#DBEAFE", icon_fg="#1D4ED8")

        body = Frame(card, bg=Colors.BG_CARD)
        body.pack(fill=X, padx=18, pady=(10, 14))

        self.log_text = scrolledtext.ScrolledText(
            body, height=6, state=DISABLED,
            bg=Colors.BG_LOG, fg=Colors.TXT_LOG,
            font=("Consolas", 10),
            relief="flat", bd=0,
            padx=14, pady=10,
            insertbackground=Colors.TXT_LOG,
            selectbackground="#334155",
            wrap="word")
        self.log_text.pack(fill=X)

    def _build_actionbar(self, parent):
        bar = Frame(parent, bg=Colors.BG_CARD,
                    highlightthickness=1,
                    highlightbackground=Colors.BORDER)
        bar.pack(fill=X)

        inner = Frame(bar, bg=Colors.BG_CARD)
        inner.pack(fill=X, padx=18, pady=14)

        # 左侧统计
        stats_box = Frame(inner, bg=Colors.BG_CARD)
        stats_box.pack(side=LEFT)

        self.stat_tables = self._mk_stat(stats_box, "—", "待处理表格")
        self.stat_rows   = self._mk_stat(stats_box, "—", "预期行数")

        # 进度条
        prog_box = Frame(stats_box, bg=Colors.BG_CARD)
        prog_box.pack(side=LEFT, padx=(20, 0))
        self.progress = ttk.Progressbar(prog_box, length=140,
                                       style="Card.Horizontal.TProgressbar",
                                       mode="determinate")
        self.progress.pack()
        self.lbl_prog = Label(prog_box, text="进度 0%",
                              bg=Colors.BG_CARD, fg=Colors.TXT_TERTIARY,
                              font=("Microsoft YaHei", 8))
        self.lbl_prog.pack(anchor=W, pady=(2, 0))

        # 右侧按钮
        btns = Frame(inner, bg=Colors.BG_CARD)
        btns.pack(side=RIGHT)

        self.btn_preview = self._mk_warning_btn(btns, "👁 预览数据", self._do_preview)
        self.btn_preview.pack(side=LEFT, padx=(0, 10))

        self.btn_run = self._mk_success_btn(btns, "▶ 开始提取", self._do_extract)
        self.btn_run.pack(side=LEFT)

    def _mk_stat(self, parent, value, label):
        box = Frame(parent, bg=Colors.BG_CARD)
        box.pack(side=LEFT, padx=(0, 22))
        v = Label(box, text=value, bg=Colors.BG_CARD, fg=Colors.PRIMARY,
                  font=("Consolas", 18, "bold"))
        v.pack(anchor=W)
        Label(box, text=label, bg=Colors.BG_CARD, fg=Colors.TXT_TERTIARY,
              font=("Microsoft YaHei", 9)).pack(anchor=W)
        return v

    # ============ 按钮工厂 ============
    def _mk_btn(self, parent, text, cmd, bg, fg, bg_hover, fg_hover,
                width=None, height=None, padx=None, pady=None):
        """创建统一风格的按钮（tk.Button 以保证颜色可控）"""
        kw = dict(
            text=text, command=cmd,
            bg=bg, fg=fg,
            activebackground=bg_hover, activeforeground=fg_hover,
            relief="flat", bd=0,
            font=("Microsoft YaHei", 10, "bold"),
            cursor="hand2",
            highlightthickness=0,
        )
        if width is not None: kw["width"] = width
        if height is not None: kw["height"] = height
        if padx is not None: kw["padx"] = padx
        if pady is not None: kw["pady"] = pady
        b = Button(parent, **kw)
        # hover 绑定
        def _enter(_=None): b.config(bg=bg_hover, fg=fg_hover)
        def _leave(_=None): b.config(bg=bg, fg=fg)
        b.bind("<Enter>", _enter)
        b.bind("<Leave>", _leave)
        return b

    def _mk_ghost_btn(self, parent, text, cmd):
        return self._mk_btn(parent, text, cmd,
                            bg="#FFFFFF", fg=Colors.TXT_PRIMARY,
                            bg_hover="#F3F4F6", fg_hover=Colors.TXT_PRIMARY,
                            padx=14, pady=6,
                            font=("Microsoft YaHei", 9))

    def _mk_outline_btn(self, parent, text, cmd):
        b = Button(parent, text=text, command=cmd,
                   bg=Colors.BG_CARD, fg=Colors.PRIMARY,
                   activebackground=Colors.PRIMARY_LIGHT,
                   activeforeground=Colors.PRIMARY_HOV,
                   relief="flat", bd=1,
                   highlightthickness=1,
                   highlightbackground=Colors.PRIMARY,
                   font=("Microsoft YaHei", 9, "bold"),
                   padx=14, pady=6,
                   cursor="hand2")
        return b

    def _mk_success_btn(self, parent, text, cmd):
        return self._mk_btn(parent, text, cmd,
                            bg=Colors.SUCCESS, fg=Colors.TXT_INVERSE,
                            bg_hover=Colors.SUCCESS_HOV, fg_hover=Colors.TXT_INVERSE,
                            padx=24, pady=10)

    def _mk_warning_btn(self, parent, text, cmd):
        return self._mk_btn(parent, text, cmd,
                            bg=Colors.WARNING, fg=Colors.TXT_INVERSE,
                            bg_hover=Colors.WARNING_HOV, fg_hover=Colors.TXT_INVERSE,
                            padx=20, pady=10)

    # ============ 依赖检查 ============
    def _check_dependencies(self):
        missing = []
        if Document is None: missing.append("python-docx")
        if Workbook  is None: missing.append("openpyxl")
        if missing:
            self._log(f"缺少依赖: {', '.join(missing)}", tag="WARN")
            self._log("请在命令行执行: pip install python-docx openpyxl", tag="INFO")

    # ============ 文件选择 ============
    def _pick_word(self):
        path = filedialog.askopenfilename(
            title="选择Word文件",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")])
        if path:
            self.word_path.set(path)
            if not self.excel_path.get():
                base = os.path.splitext(path)[0]
                self.excel_path.set(base + "_提取结果.xlsx")

    def _pick_excel(self):
        path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self.excel_path.set(path)

    # ============ 自动检测 ============
    def _auto_detect(self):
        path = self.word_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先选择有效的Word文件")
            return
        if Document is None:
            messagebox.showerror("错误", "缺少依赖 python-docx")
            return
        try:
            doc = Document(path)
            if not doc.tables:
                self._log("文档中未找到任何表格", tag="ERR")
                return
            t = doc.tables[0]
            rows = len(t.rows); cols = len(t.columns)
            self.row_count.set(rows); self.col_count.set(cols)
            self._log(f"检测到第1个表格: {rows}行 × {cols}列（共 {len(doc.tables)} 个）", tag="OK")
            messagebox.showinfo("检测完成",
                                f"检测到 {len(doc.tables)} 个表格\n结构: {rows}行 × {cols}列")
        except Exception as e:
            self._log(f"自动检测失败: {e}", tag="ERR")
            messagebox.showerror("错误", f"自动检测失败:\n{e}")

    # ============ 预览 ============
    def _do_preview(self):
        threading.Thread(target=self._preview_worker, daemon=True).start()

    def _preview_worker(self):
        try:
            data = self._extract_data(preview=True)
            self.root.after(0, lambda: self._show_preview(data))
        except Exception as e:
            self.root.after(0, lambda: self._log(f"预览失败: {e}", tag="ERR"))

    def _show_preview(self, rows_data):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        if not rows_data:
            messagebox.showinfo("预览", "未提取到数据")
            return

        cols = len(rows_data[0])
        col_ids = [f"c{i}" for i in range(cols)]
        self.preview_tree["columns"] = col_ids
        for i in range(cols):
            tag_txt = " · 标题列" if (i == 0 and self.include_title.get()) else ""
            self.preview_tree.heading(col_ids[i], text=f"Col {i+1}{tag_txt}")
            self.preview_tree.column(col_ids[i], width=120, anchor=W)

        for r in rows_data[:10]:
            tags = ("title_row",) if self.include_title.get() else ()
            self.preview_tree.insert("", END, values=r, tags=tags)

        self._log(f"共提取 {len(rows_data)} 行数据，显示前 10 行", tag="OK")

    # ============ 提取主逻辑 ============
    def _extract_data(self, preview=False):
        path = self.word_path.get().strip()
        if not path: raise ValueError("请选择Word文件")
        if not os.path.exists(path): raise ValueError("文件不存在")
        if Document is None: raise RuntimeError("缺少依赖 python-docx")

        col_count  = self.col_count.get()
        row_count  = self.row_count.get()
        title_col  = self.title_col.get()
        data_start = self.data_start.get()
        data_end   = self.data_end.get()
        skip_first = self.skip_first.get()
        include_title = self.include_title.get()

        self._log("正在解析 Word 文档...", tag="INFO")
        doc = Document(path)
        total = len(doc.tables)
        self._log(f"共发现 {total} 个表格", tag="INFO")

        all_rows = []
        for ti, table in enumerate(doc.tables):
            actual_rows = len(table.rows); actual_cols = len(table.columns)
            if ti == 0:
                self._log(f"每个表格约 {actual_rows}行 × {actual_cols}列", tag="INFO")

            start_row = 1 if skip_first else 0
            for ri in range(start_row, min(row_count, actual_rows)):
                cells = table.rows[ri].cells
                title_text = cells[title_col].text.strip() if title_col < len(cells) else ""
                data_values = []
                for c in range(data_start, data_end + 1):
                    data_values.append(cells[c].text.strip() if c < len(cells) else "")
                out_row = ([title_text] if include_title else []) + data_values
                if any(out_row):
                    all_rows.append(out_row)

            if not preview:
                self.root.after(0, lambda p=ti+1, t=total: self._update_progress(p, t))

        self._log(f"提取完成，共 {len(all_rows)} 行有效数据", tag="OK")
        return all_rows

    # ============ 导出Excel ============
    def _export_excel(self, rows_data):
        if Workbook is None: raise RuntimeError("缺少依赖 openpyxl")

        out_path = self.excel_path.get().strip()
        if not out_path: raise ValueError("请设置 Excel 输出路径")

        wb = Workbook(); ws = wb.active; ws.title = "提取结果"

        # 紫色表头（与新 UI 主色一致）
        header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
        header_font = Font(bold=True, size=11, color="FFFFFF")
        thin = Side(border_style="thin", color="E5E7EB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        title_col  = self.title_col.get()
        data_start = self.data_start.get()
        data_end   = self.data_end.get()
        include_title = self.include_title_col.get()

        headers = []
        if include_title: headers.append(f"标题列(Col{title_col+1})")
        for c in range(data_start, data_end + 1): headers.append(f"数据列(Col{c+1})")

        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.font = header_font; cell.fill = header_fill
            cell.alignment = center; cell.border = border

        for ri, row in enumerate(rows_data, 2):
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.border = border; cell.alignment = center

        # 自动列宽
        for ci in range(1, len(headers) + 1):
            max_len = len(str(headers[ci-1]))
            for ri in range(2, min(len(rows_data) + 2, 52)):
                v = ws.cell(row=ri, column=ci).value
                if v: max_len = max(max_len, len(str(v)))
            col_letter = chr(64 + ci) if ci <= 26 else "A" + chr(64 + ci - 26)
            ws.column_dimensions[col_letter].width = min(max_len * 2 + 4, 40)

        wb.save(out_path)
        self._log(f"Excel 已保存: {out_path}", tag="INFO")
        return out_path

    # ============ 开始提取 ============
    def _do_extract(self):
        threading.Thread(target=self._extract_worker, daemon=True).start()

    def _extract_worker(self):
        try:
            self.root.after(0, lambda: self.btn_run.config(state=DISABLED))
            rows_data = self._extract_data()
            out_path = self._export_excel(rows_data)
            self._log("✅ 提取成功！", tag="OK")
            self.root.after(0, lambda: self.btn_run.config(state=NORMAL))
            self.root.after(0, lambda: self.progress.config(value=100))
            self.root.after(0, lambda: self.lbl_prog.config(text="进度 100%"))
            self.root.after(0, lambda: messagebox.showinfo("完成",
                f"提取成功！\n\n共 {len(rows_data)} 行数据\n输出文件: {out_path}"))
        except Exception as e:
            self._log(str(e), tag="ERR")
            self.root.after(0, lambda: self.btn_run.config(state=NORMAL))
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))

    # ============ 辅助 ============
    def _update_progress(self, current, total):
        pct = int(current / max(total, 1) * 100)
        self.progress.config(value=pct)
        self.lbl_prog.config(text=f"进度 {pct}%")

        # 更新统计
        if current == total:
            self.stat_tables.config(text=str(total))
            try:
                row_per_table = self.row_count.get()
                self.stat_rows.config(text=str(total * row_per_table))
            except Exception:
                pass

    def _log(self, msg, tag="INFO"):
        def _write():
            ts = datetime.now().strftime("%H:%M:%S")
            tag_colors = {
                "INFO": "#60A5FA", "OK": "#34D399",
                "WARN": "#FBBF24", "ERR": "#F87171"
            }
            tcolor = tag_colors.get(tag, "#94A3B8")

            self.log_text.config(state=NORMAL)
            self.log_text.insert(END, f"{ts}  ")
            # tag 徽章
            self.log_text.insert(END, f"[{tag}] ", f"tag_{tag}")
            self.log_text.insert(END, msg + "\n")
            # 配置颜色
            self.log_text.tag_config(f"tag_{tag}", foreground=tcolor,
                                     font=("Consolas", 10, "bold"))
            self.log_text.see(END)
            self.log_text.config(state=DISABLED)
        self.root.after(0, _write)


def main():
    root = Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    WordToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
