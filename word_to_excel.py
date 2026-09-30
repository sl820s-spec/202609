# -*- coding: utf-8 -*-
"""
Word表格数据提取到Excel工具
功能：将Word文档中每页一个的相同结构表格提取并汇总到Excel
"""

import os
import sys
import threading
from tkinter import (
    Tk, Frame, Label, Button, Entry, StringVar, IntVar, BooleanVar,
    FileDialog, messagebox, ttk, scrolledtext, END, N, S, E, W,
    LEFT, RIGHT, TOP, BOTTOM, X, Y, BOTH, SUNKEN, GROOVE, NORMAL, DISABLED
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


class WordToExcelApp:
    """主应用类"""

    def __init__(self, root):
        self.root = root
        self.root.title("Word表格数据提取工具 v1.0")
        self.root.geometry("820x640")
        self.root.minsize(760, 580)

        # 变量
        self.word_path = StringVar()
        self.excel_path = StringVar()
        self.col_count = IntVar(value=5)
        self.row_count = IntVar(value=30)
        self.title_col_index = IntVar(value=0)       # 标题列索引（0-based）
        self.data_col_start = IntVar(value=1)         # 数据起始列
        self.data_col_end = IntVar(value=3)           # 数据结束列
        self.skip_first_row = BooleanVar(value=False) # 是否跳过每页首行（表头）
        self.include_title_col = BooleanVar(value=True) # 导出时包含标题列

        self._build_ui()
        self._check_dependencies()

    # ========== UI 构建 ==========
    def _build_ui(self):
        """构建界面布局"""
        pad = {"padx": 10, "pady": 6}

        # ---------- 文件选择区 ----------
        file_frame = LabelFrame(self.root, text=" 文件选择 ", font=("Microsoft YaHei", 10, "bold"))
        file_frame.pack(fill=X, padx=12, pady=(12, 4))

        Label(file_frame, text="Word文件:", font=("Microsoft YaHei", 9)).grid(
            row=0, column=0, sticky=W, padx=8, pady=8)
        Entry(file_frame, textvariable=self.word_path, width=60).grid(
            row=0, column=1, padx=4, pady=8)
        Button(file_frame, text="浏览...", command=self._pick_word, width=10).grid(
            row=0, column=2, padx=6, pady=8)

        Label(file_frame, text="Excel输出:", font=("Microsoft YaHei", 9)).grid(
            row=1, column=0, sticky=W, padx=8, pady=8)
        Entry(file_frame, textvariable=self.excel_path, width=60).grid(
            row=1, column=1, padx=4, pady=8)
        Button(file_frame, text="浏览...", command=self._pick_excel, width=10).grid(
            row=1, column=2, padx=6, pady=8)

        # ---------- 参数配置区 ----------
        param_frame = LabelFrame(self.root, text=" 表格参数配置 ", font=("Microsoft YaHei", 10, "bold"))
        param_frame.pack(fill=X, padx=12, pady=4)

        # 行列数
        Label(param_frame, text="每页表格列数:", font=("Microsoft YaHei", 9)).grid(
            row=0, column=0, sticky=W, padx=8, pady=8)
        Entry(param_frame, textvariable=self.col_count, width=8).grid(row=0, column=1, padx=4, pady=8)

        Label(param_frame, text="每页表格行数:", font=("Microsoft YaHei", 9)).grid(
            row=0, column=2, sticky=W, padx=(16, 4), pady=8)
        Entry(param_frame, textvariable=self.row_count, width=8).grid(row=0, column=3, padx=4, pady=8)

        # 列映射
        Label(param_frame, text="标题列索引:", font=("Microsoft YaHei", 9)).grid(
            row=1, column=0, sticky=W, padx=8, pady=8)
        Entry(param_frame, textvariable=self.title_col_index, width=8).grid(row=1, column=1, padx=4, pady=8)

        Label(param_frame, text="数据列范围:", font=("Microsoft YaHei", 9)).grid(
            row=1, column=2, sticky=W, padx=(16, 4), pady=8)
        Entry(param_frame, textvariable=self.data_col_start, width=6).grid(row=1, column=3, padx=(0, 2), pady=8, sticky=E)
        Label(param_frame, text="~", font=("Microsoft YaHei", 9)).grid(row=1, column=3, padx=(36, 0), pady=8, sticky=W)
        Entry(param_frame, textvariable=self.data_col_end, width=6).grid(row=1, column=3, padx=(44, 4), pady=8, sticky=W)

        # 选项
        check_frame = Frame(param_frame)
        check_frame.grid(row=2, column=0, columnspan=4, sticky=W, padx=8, pady=4)
        ttk.Checkbutton(check_frame, text="跳过每页第一行（若为表头）",
                        variable=self.skip_first_row).pack(side=LEFT, padx=(0, 20))
        ttk.Checkbutton(check_frame, text="导出结果中包含标题列",
                        variable=self.include_title_col).pack(side=LEFT)

        # 按钮：自动检测
        Button(param_frame, text="自动检测表格结构",
               command=self._auto_detect, width=18,
               bg="#E8F4FD").grid(row=0, column=4, rowspan=2, padx=12, pady=8)

        # ---------- 预览区 ----------
        preview_frame = LabelFrame(self.root, text=" 数据预览（前10行） ", font=("Microsoft YaHei", 10, "bold"))
        preview_frame.pack(fill=BOTH, expand=True, padx=12, pady=4)

        self.preview_tree = ttk.Treeview(preview_frame, show="headings", height=8)
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=vsb.set)
        self.preview_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=(4, 0), pady=4)
        vsb.pack(side=RIGHT, fill=Y, padx=(0, 4), pady=4)

        # ---------- 日志区 ----------
        log_frame = LabelFrame(self.root, text=" 运行日志 ", font=("Microsoft YaHei", 10, "bold"))
        log_frame.pack(fill=X, padx=12, pady=4)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=5, font=("Consolas", 9), state=DISABLED,
            background="#1E1E1E", foreground="#D4D4D4"
        )
        self.log_text.pack(fill=X, padx=4, pady=4)

        # ---------- 底部操作区 ----------
        bottom_frame = Frame(self.root)
        bottom_frame.pack(fill=X, padx=12, pady=(4, 12))

        self.progress = ttk.Progressbar(bottom_frame, mode="determinate", length=400)
        self.progress.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))

        self.btn_preview = Button(bottom_frame, text="预览数据", command=self._do_preview,
                                  width=12, bg="#FFC107", font=("Microsoft YaHei", 9, "bold"))
        self.btn_preview.pack(side=RIGHT, padx=4)

        self.btn_run = Button(bottom_frame, text="开始提取", command=self._do_extract,
                              width=12, bg="#4CAF50", fg="white",
                              font=("Microsoft YaHei", 9, "bold"))
        self.btn_run.pack(side=RIGHT, padx=4)

    # ========== 依赖检查 ==========
    def _check_dependencies(self):
        missing = []
        if Document is None:
            missing.append("python-docx")
        if Workbook is None:
            missing.append("openpyxl")
        if missing:
            self._log(f"[警告] 缺少依赖: {', '.join(missing)}")
            self._log("请在命令行执行: pip install python-docx openpyxl")

    # ========== 文件选择 ==========
    def _pick_word(self):
        path = filedialog.askopenfilename(
            title="选择Word文件",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if path:
            self.word_path.set(path)
            # 自动生成输出路径
            if not self.excel_path.get():
                base = os.path.splitext(path)[0]
                self.excel_path.set(base + "_提取结果.xlsx")

    def _pick_excel(self):
        path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_path.set(path)

    # ========== 自动检测 ==========
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
                self._log("[错误] 文档中未找到任何表格")
                return
            t = doc.tables[0]
            rows = len(t.rows)
            cols = len(t.columns)
            self.row_count.set(rows)
            self.col_count.set(cols)
            self._log(f"[检测] 第一个表格: {rows}行 × {cols}列（共{len(doc.tables)}个表格）")
            messagebox.showinfo("检测完成", f"检测到 {len(doc.tables)} 个表格\n结构: {rows}行 × {cols}列")
        except Exception as e:
            self._log(f"[错误] 自动检测失败: {e}")
            messagebox.showerror("错误", f"自动检测失败:\n{e}")

    # ========== 预览 ==========
    def _do_preview(self):
        threading.Thread(target=self._preview_worker, daemon=True).start()

    def _preview_worker(self):
        try:
            data = self._extract_data(preview=True)
            self.root.after(0, lambda: self._show_preview(data))
        except Exception as e:
            self.root.after(0, lambda: self._log(f"[错误] 预览失败: {e}"))

    def _show_preview(self, rows_data):
        # 清空
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        if not rows_data:
            messagebox.showinfo("预览", "未提取到数据")
            return

        cols = rows_data[0]
        self.preview_tree["columns"] = [str(i) for i in range(len(cols))]
        for i, _ in enumerate(cols):
            self.preview_tree.heading(i, text=f"列{i+1}")
            self.preview_tree.column(i, width=100, anchor="center")

        for r in rows_data[:10]:
            self.preview_tree.insert("", END, values=r)

        self._log(f"[预览] 共提取 {len(rows_data)} 行数据，显示前10行")

    # ========== 提取主逻辑 ==========
    def _extract_data(self, preview=False):
        path = self.word_path.get().strip()
        if not path:
            raise ValueError("请选择Word文件")
        if not os.path.exists(path):
            raise ValueError("文件不存在")
        if Document is None:
            raise RuntimeError("缺少依赖 python-docx")

        col_count = self.col_count.get()
        row_count = self.row_count.get()
        title_col = self.title_col_index.get()
        data_start = self.data_col_start.get()
        data_end = self.data_col_end.get()
        skip_first = self.skip_first_row.get()
        include_title = self.include_title_col.get()

        self._log(f"[读取] 正在解析Word文档...")
        doc = Document(path)
        total_tables = len(doc.tables)
        self._log(f"[读取] 共发现 {total_tables} 个表格")

        all_rows = []  # [ (table_idx, row_idx, title_text, data_1, data_2, ...) ]

        for ti, table in enumerate(doc.tables):
            actual_rows = len(table.rows)
            actual_cols = len(table.columns)

            if ti == 0:
                self._log(f"[结构] 每个表格约 {actual_rows}行 × {actual_cols}列")

            # 使用实际列数（更灵活）
            use_cols = min(col_count, actual_cols)
            start_row = 1 if skip_first else 0

            for ri in range(start_row, min(row_count, actual_rows)):
                cells = table.rows[ri].cells
                # 取标题文本
                title_text = cells[title_col].text.strip() if title_col < len(cells) else ""

                # 取数据列
                data_values = []
                for c in range(data_start, data_end + 1):
                    if c < len(cells):
                        data_values.append(cells[c].text.strip())
                    else:
                        data_values.append("")

                # 组装输出行
                out_row = []
                if include_title:
                    out_row.append(title_text)
                out_row.extend(data_values)

                # 跳过全空行
                if any(out_row):
                    all_rows.append(out_row)

            if not preview:
                self.root.after(0, lambda p=ti+1, t=total_tables: self._update_progress(p, t))

        self._log(f"[提取] 完成，共 {len(all_rows)} 行有效数据")
        return all_rows

    # ========== 导出Excel ==========
    def _export_excel(self, rows_data):
        if Workbook is None:
            raise RuntimeError("缺少依赖 openpyxl")

        out_path = self.excel_path.get().strip()
        if not out_path:
            raise ValueError("请设置Excel输出路径")

        wb = Workbook()
        ws = wb.active
        ws.title = "提取结果"

        # 样式
        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font_white = Font(bold=True, size=11, color="FFFFFF")
        thin = Side(border_style="thin", color="B4B4B4")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # 写表头
        col_count = self.col_count.get()
        title_col = self.title_col_index.get()
        data_start = self.data_col_start.get()
        data_end = self.data_col_end.get()
        include_title = self.include_title_col.get()

        headers = []
        if include_title:
            headers.append(f"标题列(Col{title_col+1})")
        for c in range(data_start, data_end + 1):
            headers.append(f"数据列(Col{c+1})")

        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

        # 写数据
        for ri, row in enumerate(rows_data, 2):
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.border = border
                cell.alignment = center

        # 自动列宽（简单估算）
        for ci in range(1, len(headers) + 1):
            max_len = len(str(headers[ci-1]))
            for ri in range(2, min(len(rows_data) + 2, 52)):  # 采样前50行
                v = ws.cell(row=ri, column=ci).value
                if v:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[chr(64 + ci) if ci <= 26 else "A" + chr(64 + ci - 26)].width = min(max_len * 2 + 4, 40)

        wb.save(out_path)
        self._log(f"[保存] Excel已保存: {out_path}")
        return out_path

    # ========== 开始提取 ==========
    def _do_extract(self):
        threading.Thread(target=self._extract_worker, daemon=True).start()

    def _extract_worker(self):
        try:
            self.root.after(0, lambda: self.btn_run.config(state=DISABLED))
            self._log("=" * 50)
            rows_data = self._extract_data()
            out_path = self._export_excel(rows_data)
            self._log("[完成] ✅ 提取成功！")
            self.root.after(0, lambda: self.btn_run.config(state=NORMAL))
            self.root.after(0, lambda: self.progress.config(value=100))
            self.root.after(0, lambda: messagebox.showinfo("完成",
                f"提取成功！\n\n共 {len(rows_data)} 行数据\n输出文件: {out_path}"))
        except Exception as e:
            self._log(f"[错误] {e}")
            self.root.after(0, lambda: self.btn_run.config(state=NORMAL))
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))

    # ========== 辅助 ==========
    def _update_progress(self, current, total):
        pct = int(current / max(total, 1) * 100)
        self.progress.config(value=pct)

    def _log(self, msg):
        def _write():
            self.log_text.config(state=NORMAL)
            self.log_text.insert(END, msg + "\n")
            self.log_text.see(END)
            self.log_text.config(state=DISABLED)
        self.root.after(0, _write)


def main():
    root = Tk()
    # 设置主题尝试
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    app = WordToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
