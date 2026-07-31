"""保险行业季度调研报告生成器：跨平台桌面界面（Windows / macOS）。"""

from __future__ import annotations

import os
import queue
import shutil
import sys
import tempfile
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from app_utils import (
    build_output_paths,
    build_section_preview_rows,
    copy_selected_files,
    default_desktop_dir,
    default_doc_fonts,
    is_windows,
    layer_to_level,
    open_path,
    ui_font_family,
)
from report_aggregator import (
    FontConfig,
    FormatTemplate,
    aggregate,
    load_format_template,
    parse_input_dir,
    _convert_docx_to_pdf_with_word,
)

APP_TITLE = "保险行业季度调研报告生成器"
_DOC_FONTS = default_doc_fonts()
FONT_OPTIONS = list(_DOC_FONTS["options"])
UI_FONT = ui_font_family()
COLOR_OPTIONS = {"黑色": "000000", "深蓝": "1F3864", "红色": "C00000", "绿色": "375623", "灰色": "595959"}
SIZE_OPTIONS = ["10", "10.5", "11", "12", "14", "16", "18", "22", "24"]
EDITABLE_COLS = {"#2": "section_id", "#3": "title", "#4": "layer"}
LAYER_CHOICES = ("章", "节", "小节", "前言")


class ReportDesktopApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1220x820")
        self.root.minsize(1000, 680)
        self.root.configure(bg="#EAF0F7")

        self.files: list[str] = []
        self.events: queue.Queue = queue.Queue()
        self.last_docx = self.last_pdf = None
        self.busy = False
        self.worker = None
        self.pending_temp_dir = None
        self.preview_ready = False
        self.preview_failures: list = []
        self._preview_original: dict[str, dict] = {}
        self._edit_widget = self._edit_item = self._edit_field = None
        self.format_template: FormatTemplate | None = None

        self.vars: dict[str, tk.StringVar] = {}
        self.header_vars: dict[str, tk.BooleanVar] = {}
        self.font_vars: dict[str, tk.StringVar] = {}
        self.indent_var = tk.StringVar(value="2")
        self.status_var = tk.StringVar(value="请先添加本季度的 DOCX 材料")
        self.use_template_var = tk.BooleanVar(value=False)
        self.template_path_var = tk.StringVar(value="")
        self.template_info_var = tk.StringVar(value="未选择格式模板（可选）")
        self.use_tpl_titles_var = tk.BooleanVar(value=True)

        self._build_style()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(120, self._process_events)
        self._try_default_template()
        # 演示录屏：环境变量 REPORT_DEMO=1 时自动走通「加载样例→预览→生成」
        if os.environ.get("REPORT_DEMO", "").strip() in ("1", "true", "TRUE", "yes"):
            self.root.after(800, self._demo_autorun)

    def _try_default_template(self) -> None:
        """若同目录/上一级存在 Q2 终版，预填为默认模板路径（不强制启用）。"""
        here = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(here, "..", "保险行业季度调研报告_终版_2026Q2.docx"),
            os.path.join(here, "保险行业季度调研报告_终版_2026Q2.docx"),
            os.path.join(os.path.expanduser("~"), "Desktop", "报告合成", "保险行业季度调研报告_终版_2026Q2.docx"),
            os.path.join(os.path.expanduser("~"), "桌面", "报告合成", "保险行业季度调研报告_终版_2026Q2.docx"),
        ]
        # PyInstaller 打包后：可执行文件旁
        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(os.path.abspath(sys.executable))
            candidates.insert(0, os.path.join(exe_dir, "保险行业季度调研报告_终版_2026Q2.docx"))
        for c in candidates:
            ap = os.path.abspath(c)
            if os.path.isfile(ap):
                self.template_path_var.set(ap)
                self.template_info_var.set(f"可选用模板：{os.path.basename(ap)}")
                break

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        for theme in (("vista",) if is_windows() else ()) + ("clam", "aqua", "default"):
            try:
                style.theme_use(theme)
                break
            except tk.TclError:
                continue
        style.configure("App.TFrame", background="#EAF0F7")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure("Title.TLabel", background="#17365D", foreground="#FFFFFF", font=(UI_FONT, 16, "bold"))
        style.configure("Subtitle.TLabel", background="#17365D", foreground="#DCE8F5", font=(UI_FONT, 10))
        style.configure("CardTitle.TLabel", background="#FFFFFF", foreground="#17365D", font=(UI_FONT, 11, "bold"))
        style.configure("Body.TLabel", background="#FFFFFF", foreground="#334155", font=(UI_FONT, 9))
        style.configure("Primary.TButton", font=(UI_FONT, 10, "bold"))
        style.configure("Action.TButton", font=(UI_FONT, 9))

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, bg="#17365D", height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        ttk.Label(header, text="📊  " + APP_TITLE, style="Title.TLabel").pack(anchor="w", padx=18, pady=(10, 0))
        ttk.Label(
            header,
            text="多份输入合成 · 可选格式模板 · 手动调序/改标题 · 导出 DOCX/PDF",
            style="Subtitle.TLabel",
        ).pack(anchor="w", padx=22, pady=(2, 10))

        body = ttk.Frame(self.root, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=12, pady=12)
        body.columnconfigure(0, weight=0, minsize=340)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self.left = ttk.Frame(body, style="Card.TFrame", padding=12)
        self.left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.right = ttk.Frame(body, style="Card.TFrame", padding=12)
        self.right.grid(row=0, column=1, sticky="nsew")
        self._build_parameters()
        self._build_workspace()

    def _build_parameters(self) -> None:
        # ---- 格式模板 ----
        ttk.Label(self.left, text="格式模板（可选）", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Checkbutton(
            self.left,
            text="启用格式模板（参考历史终版的章节顺序/标题/字体）",
            variable=self.use_template_var,
            command=self._on_template_toggle,
        ).pack(anchor="w", pady=(6, 2))
        ttk.Checkbutton(
            self.left,
            text="匹配章节时优先使用模板标题",
            variable=self.use_tpl_titles_var,
        ).pack(anchor="w", pady=(0, 4))
        tpl_row = ttk.Frame(self.left, style="Card.TFrame")
        tpl_row.pack(fill="x", pady=2)
        ttk.Entry(tpl_row, textvariable=self.template_path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(tpl_row, text="浏览…", command=self.choose_template, style="Action.TButton").pack(side="left", padx=(6, 0))
        ttk.Button(tpl_row, text="加载", command=self.load_template, style="Action.TButton").pack(side="left", padx=(4, 0))
        ttk.Label(self.left, textvariable=self.template_info_var, style="Body.TLabel").pack(anchor="w", pady=(2, 6))
        ttk.Label(
            self.left,
            text="例：选 Q2 终版 docx 作模板，再合成 Q3 输入材料",
            style="Body.TLabel",
        ).pack(anchor="w")

        ttk.Separator(self.left).pack(fill="x", pady=10)
        ttk.Label(self.left, text="报告参数", style="CardTitle.TLabel").pack(anchor="w")
        defaults = {
            "period": "2026年第三季度",
            "title": "中国保险行业调研报告",
            "subtitle": "市场环境 · 保费增长 · 渠道变革 · 趋势展望",
            "org": "保险行业研究中心",
            "date": "2026年10月",
            "output_dir": default_desktop_dir(),
        }
        for key, label in {
            "period": "报告周期", "title": "报告主标题", "subtitle": "副标题",
            "org": "编制单位", "date": "发布日期",
        }.items():
            ttk.Label(self.left, text=label, style="Body.TLabel").pack(anchor="w", pady=(8, 0))
            self.vars[key] = tk.StringVar(value=defaults[key])
            ttk.Entry(self.left, textvariable=self.vars[key]).pack(fill="x", pady=2)

        ttk.Separator(self.left).pack(fill="x", pady=10)
        ttk.Label(self.left, text="页眉内容", style="CardTitle.TLabel").pack(anchor="w")
        hf = ttk.Frame(self.left, style="Card.TFrame")
        hf.pack(fill="x", pady=4)
        for key, text, default in (("title", "标题", True), ("org", "单位", False), ("date", "日期", False), ("period", "周期", False)):
            bv = tk.BooleanVar(value=default)
            self.header_vars[key] = bv
            ttk.Checkbutton(hf, text=text, variable=bv).pack(side="left", padx=(0, 8))

        ttk.Separator(self.left).pack(fill="x", pady=10)
        ttk.Label(self.left, text="格式设置（可被模板覆盖）", style="CardTitle.TLabel").pack(anchor="w")
        for level, title in (("h1", "一级标题"), ("h2", "二级标题"), ("h3", "三级标题"), ("body", "正文")):
            row = ttk.Frame(self.left, style="Card.TFrame")
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=title, style="Body.TLabel", width=8).pack(side="left")
            self.font_vars[f"{level}_font"] = tk.StringVar(
                value=_DOC_FONTS["heading"] if level != "body" else _DOC_FONTS["body"]
            )
            self.font_vars[f"{level}_size"] = tk.StringVar(value={"h1": "18", "h2": "14", "h3": "12", "body": "12"}[level])
            self.font_vars[f"{level}_bold"] = tk.StringVar(value="是" if level != "body" else "否")
            self.font_vars[f"{level}_color"] = tk.StringVar(value="深蓝" if level != "body" else "黑色")
            ttk.Combobox(row, textvariable=self.font_vars[f"{level}_font"], values=FONT_OPTIONS, width=8, state="readonly").pack(side="left", padx=2)
            ttk.Combobox(row, textvariable=self.font_vars[f"{level}_size"], values=SIZE_OPTIONS, width=4, state="readonly").pack(side="left", padx=2)
            ttk.Combobox(row, textvariable=self.font_vars[f"{level}_bold"], values=["是", "否"], width=3, state="readonly").pack(side="left", padx=2)
            ttk.Combobox(row, textvariable=self.font_vars[f"{level}_color"], values=list(COLOR_OPTIONS), width=5, state="readonly").pack(side="left", padx=2)

        ir = ttk.Frame(self.left, style="Card.TFrame")
        ir.pack(fill="x", pady=4)
        ttk.Label(ir, text="正文首行缩进", style="Body.TLabel").pack(side="left")
        ttk.Entry(ir, textvariable=self.indent_var, width=6).pack(side="left", padx=6)
        ttk.Label(ir, text="字符", style="Body.TLabel").pack(side="left")

        ttk.Separator(self.left).pack(fill="x", pady=10)
        ttk.Label(self.left, text="输出位置", style="CardTitle.TLabel").pack(anchor="w")
        out_row = ttk.Frame(self.left, style="Card.TFrame")
        out_row.pack(fill="x", pady=4)
        self.vars["output_dir"] = tk.StringVar(value=defaults["output_dir"])
        ttk.Entry(out_row, textvariable=self.vars["output_dir"]).pack(side="left", fill="x", expand=True)
        ttk.Button(out_row, text="浏览", command=self.choose_output_dir, style="Action.TButton").pack(side="left", padx=(6, 0))

    def _build_workspace(self) -> None:
        self.right.rowconfigure(2, weight=1)
        self.right.columnconfigure(0, weight=1)

        top = ttk.Frame(self.right, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="材料与处理中心", style="CardTitle.TLabel").pack(side="left")
        ttk.Button(top, text="＋ 添加 DOCX", command=self.add_files, style="Action.TButton").pack(side="right", padx=3)
        self.preview_btn = ttk.Button(top, text="🔍 预览章节", command=self.start_preview, style="Action.TButton")
        self.preview_btn.pack(side="right", padx=3)
        ttk.Button(top, text="清空", command=self.clear_files, style="Action.TButton").pack(side="right", padx=3)
        ttk.Button(top, text="移除选中", command=self.remove_selected, style="Action.TButton").pack(side="right", padx=3)

        mid = ttk.Frame(self.right, style="Card.TFrame")
        mid.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(mid, text="双击改编号/标题/层级；右侧按钮调序（=终版顺序）", style="Body.TLabel").pack(side="left")
        ttk.Button(mid, text="恢复解析结果", command=self.reset_preview_edits, style="Action.TButton").pack(side="right")

        notebook = ttk.Notebook(self.right)
        notebook.grid(row=2, column=0, sticky="nsew", pady=8)
        file_tab = ttk.Frame(notebook, style="Card.TFrame")
        preview_tab = ttk.Frame(notebook, style="Card.TFrame")
        log_tab = ttk.Frame(notebook, style="Card.TFrame")
        notebook.add(file_tab, text="输入文件")
        notebook.add(preview_tab, text="章节预览")
        notebook.add(log_tab, text="处理日志")

        self.file_list = tk.Listbox(file_tab, activestyle="dotbox", font=("Consolas", 10))
        self.file_list.pack(fill="both", expand=True, padx=4, pady=4)

        wrap = ttk.Frame(preview_tab, style="Card.TFrame")
        wrap.pack(fill="both", expand=True)
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        tree_frame = ttk.Frame(wrap, style="Card.TFrame")
        tree_frame.grid(row=0, column=0, sticky="nsew")
        cols = ("order", "section_id", "title", "layer", "block_count", "source_count", "confidence", "status")
        self.preview_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        heads = {"order": "顺序", "section_id": "章节编号", "title": "识别标题", "layer": "层级",
                 "block_count": "内容块", "source_count": "来源数", "confidence": "置信度", "status": "状态"}
        widths = {"order": 50, "section_id": 90, "title": 230, "layer": 70, "block_count": 70,
                  "source_count": 70, "confidence": 70, "status": 90}
        for c in cols:
            self.preview_tree.heading(c, text=heads[c])
            self.preview_tree.column(c, width=widths[c], anchor="center" if c != "title" else "w")
        ys = ttk.Scrollbar(tree_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=ys.set)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.preview_tree.tag_configure("warning", foreground="#C00000")
        self.preview_tree.tag_configure("merged", foreground="#9A6700")
        self.preview_tree.tag_configure("preamble", foreground="#1F3864")
        self.preview_tree.tag_configure("edited", foreground="#0B6E4F")
        self.preview_tree.bind("<Double-1>", self._on_tree_double_click)
        self.preview_tree.bind("<Control-Up>", lambda e: self._move_selected(-1))
        self.preview_tree.bind("<Control-Down>", lambda e: self._move_selected(1))

        order_btns = ttk.Frame(wrap, style="Card.TFrame", padding=(6, 0))
        order_btns.grid(row=0, column=1, sticky="ns")
        ttk.Label(order_btns, text="调整顺序", style="CardTitle.TLabel").pack(pady=(0, 6))
        ttk.Button(order_btns, text="↑ 上移", command=lambda: self._move_selected(-1), style="Action.TButton").pack(fill="x", pady=2)
        ttk.Button(order_btns, text="↓ 下移", command=lambda: self._move_selected(1), style="Action.TButton").pack(fill="x", pady=2)
        ttk.Button(order_btns, text="⤒ 置顶", command=lambda: self._move_selected("top"), style="Action.TButton").pack(fill="x", pady=2)
        ttk.Button(order_btns, text="⤓ 置底", command=lambda: self._move_selected("bottom"), style="Action.TButton").pack(fill="x", pady=2)
        ttk.Label(order_btns, text="Ctrl+↑/↓", style="Body.TLabel").pack(pady=(10, 0))

        self.log = ScrolledText(log_tab, height=12, font=("Consolas", 9), relief="flat", bg="#F8FAFC")
        self.log.pack(fill="both", expand=True, padx=4, pady=4)

        bottom = ttk.Frame(self.right, style="Card.TFrame")
        bottom.grid(row=3, column=0, sticky="ew")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 6))
        ttk.Label(bottom, textvariable=self.status_var, style="Body.TLabel").pack(anchor="w")
        btn_row = ttk.Frame(bottom, style="Card.TFrame")
        btn_row.pack(fill="x", pady=6)
        self.generate_btn = ttk.Button(btn_row, text="🚀 生成终版报告", command=self.start_generation, style="Primary.TButton", state="disabled")
        self.generate_btn.pack(side="left")
        self.open_docx_btn = ttk.Button(btn_row, text="打开 DOCX", command=lambda: self._open_path(self.last_docx), style="Action.TButton", state="disabled")
        self.open_docx_btn.pack(side="left", padx=6)
        self.open_pdf_btn = ttk.Button(btn_row, text="打开 PDF", command=lambda: self._open_path(self.last_pdf), style="Action.TButton", state="disabled")
        self.open_pdf_btn.pack(side="left", padx=6)
        ttk.Button(btn_row, text="打开输出目录", command=self.open_output_dir, style="Action.TButton").pack(side="left", padx=6)

    # ---- template ----
    def _on_template_toggle(self) -> None:
        if self.use_template_var.get() and self.template_path_var.get().strip():
            self.load_template()

    def choose_template(self) -> None:
        cur = self.template_path_var.get().strip()
        initial = os.path.dirname(cur) if cur else default_desktop_dir()
        path = filedialog.askopenfilename(
            title="选择格式模板（历史合成终版 DOCX）",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            initialdir=initial if os.path.isdir(initial) else default_desktop_dir(),
        )
        if path:
            self.template_path_var.set(path)
            self.load_template()

    def load_template(self) -> None:
        path = self.template_path_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "请先选择模板 DOCX 路径。")
            return
        if not os.path.isfile(path):
            messagebox.showerror(APP_TITLE, f"文件不存在：\n{path}")
            return
        try:
            ft = load_format_template(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"加载模板失败：\n{exc}")
            self.format_template = None
            return
        self.format_template = ft
        self.use_template_var.set(True)
        self.template_info_var.set(
            f"已加载：{os.path.basename(path)} · {len(ft.specs)} 个章节 · "
            f"H1 {ft.font_config.h1_font}/{ft.font_config.h1_size}"
        )
        self._append_log(f"已加载格式模板：{path}")
        self._append_log(f"  模板章节 {len(ft.specs)} 个；封面「{ft.title}」/ {ft.period}")
        # 可选：用模板封面字段填充空参数
        if ft.period and (not self.vars["period"].get() or "第三季度" in self.vars["period"].get() or True):
            # 仅提示用户，不强制覆盖周期（Q3 应自己改）
            pass
        if ft.title and not self.vars["title"].get().strip():
            self.vars["title"].set(ft.title)
        if ft.subtitle and not self.vars["subtitle"].get().strip():
            self.vars["subtitle"].set(ft.subtitle)
        if ft.org and not self.vars["org"].get().strip():
            self.vars["org"].set(ft.org)
        # 应用模板字体到 UI
        if ft.font_config:
            fc = ft.font_config
            mapping = [
                ("h1", fc.h1_font, fc.h1_size, fc.h1_bold),
                ("h2", fc.h2_font, fc.h2_size, fc.h2_bold),
                ("h3", fc.h3_font, fc.h3_size, fc.h3_bold),
                ("body", fc.body_font, fc.body_size, fc.body_bold),
            ]
            for level, font, size, bold in mapping:
                if font:
                    self.font_vars[f"{level}_font"].set(font)
                if size:
                    # 选最接近的选项
                    s = str(int(size) if float(size).is_integer() else size)
                    if s not in SIZE_OPTIONS:
                        SIZE_OPTIONS.append(s)
                    self.font_vars[f"{level}_size"].set(s)
                self.font_vars[f"{level}_bold"].set("是" if bold else "否")
        self.status_var.set(f"格式模板已就绪：{len(ft.specs)} 个参考章节")

    # ---- files ----
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="选择板块 DOCX 文件", filetypes=[("Word 文档", "*.docx")])
        for p in paths:
            ap = os.path.abspath(p)
            if ap not in self.files:
                self.files.append(ap)
                self.file_list.insert(tk.END, ap)
        self._invalidate_preview()
        self.status_var.set(f"已选择 {len(self.files)} 份材料，请预览章节")

    def remove_selected(self) -> None:
        for idx in reversed(self.file_list.curselection()):
            self.file_list.delete(idx)
            del self.files[idx]
        self._invalidate_preview()
        self.status_var.set(f"已选择 {len(self.files)} 份材料，请预览章节")

    def clear_files(self) -> None:
        self.files.clear()
        self.file_list.delete(0, tk.END)
        self._invalidate_preview()
        self.status_var.set("请先添加本季度的 DOCX 材料")

    def _invalidate_preview(self) -> None:
        self.preview_ready = False
        self.preview_failures = []
        self._preview_original.clear()
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.generate_btn.configure(state="disabled")
        if self.pending_temp_dir:
            shutil.rmtree(self.pending_temp_dir, ignore_errors=True)
            self.pending_temp_dir = None
        self._close_edit_widget(save=False)

    def reset_preview_edits(self) -> None:
        if not self.preview_tree.get_children():
            return
        rows = sorted(
            ((orig.get("order", 0), iid, orig) for iid, orig in self._preview_original.items()),
            key=lambda x: x[0],
        )
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        for order, iid, orig in rows:
            self.preview_tree.insert(
                "", tk.END, iid=iid,
                values=(order, orig["section_id"], orig["title"], orig["layer"],
                        orig.get("block_count", ""), orig.get("source_count", ""),
                        orig.get("confidence", ""), orig.get("status", "正常")),
                tags=orig.get("tags", ("normal",)),
            )
        self._renumber_order()
        self.status_var.set("已恢复全部为解析结果（含顺序）")

    def _move_selected(self, direction) -> str:
        if self.busy:
            return "break"
        sel = self.preview_tree.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "请先选中一行。")
            return "break"
        item = sel[0]
        children = list(self.preview_tree.get_children())
        idx = children.index(item)
        if direction == "top":
            new_idx = 0
        elif direction == "bottom":
            new_idx = len(children) - 1
        else:
            new_idx = idx + int(direction)
        if new_idx < 0 or new_idx >= len(children) or new_idx == idx:
            return "break"
        self.preview_tree.move(item, "", new_idx)
        self.preview_tree.selection_set(item)
        self.preview_tree.see(item)
        self._renumber_order()
        self.status_var.set(f"顺序已调整（第 {new_idx + 1} 位）")
        return "break"

    def _renumber_order(self) -> None:
        for i, item in enumerate(self.preview_tree.get_children(), 1):
            vals = list(self.preview_tree.item(item, "values"))
            if vals:
                vals[0] = i
                self.preview_tree.item(item, values=vals)

    def _collect_section_order(self) -> list[str]:
        return list(self.preview_tree.get_children())

    def _on_tree_double_click(self, event) -> None:
        if self.busy or self.preview_tree.identify("region", event.x, event.y) != "cell":
            return
        field = EDITABLE_COLS.get(self.preview_tree.identify_column(event.x))
        item = self.preview_tree.identify_row(event.y)
        if field and item:
            self._start_cell_edit(item, field)

    def _start_cell_edit(self, item: str, field: str) -> None:
        self._close_edit_widget(save=True)
        bbox = self.preview_tree.bbox(item, field)
        if not bbox:
            return
        x, y, w, h = bbox
        vals = list(self.preview_tree.item(item, "values"))
        col_index = {"section_id": 1, "title": 2, "layer": 3}[field]
        current = str(vals[col_index]) if len(vals) > col_index else ""
        if field == "layer":
            widget: tk.Widget = ttk.Combobox(self.preview_tree, values=LAYER_CHOICES, state="readonly", font=(UI_FONT, 9))
            widget.set(current if current in LAYER_CHOICES else "节")
            widget.bind("<<ComboboxSelected>>", lambda _e: self._close_edit_widget(save=True))
        else:
            widget = tk.Entry(self.preview_tree, font=(UI_FONT, 9))
            widget.insert(0, current)
            widget.select_range(0, tk.END)
            widget.bind("<Return>", lambda _e: self._close_edit_widget(save=True))
            widget.bind("<Escape>", lambda _e: self._close_edit_widget(save=False))
        widget.place(x=x, y=y, width=max(w, 80), height=h)
        widget.focus_set()
        widget.bind("<FocusOut>", lambda _e: self.root.after(80, lambda: self._close_edit_widget(save=True)))
        self._edit_widget, self._edit_item, self._edit_field = widget, item, field

    def _close_edit_widget(self, save: bool) -> None:
        widget, item, field = self._edit_widget, self._edit_item, self._edit_field
        self._edit_widget = self._edit_item = self._edit_field = None
        if not widget or not item or not field:
            return
        try:
            text = widget.get().strip() if save else None
        finally:
            try:
                widget.destroy()
            except Exception:
                pass
        if not save or text is None:
            return
        vals = list(self.preview_tree.item(item, "values"))
        orig = self._preview_original.get(item, {})
        if field == "section_id":
            vals[1] = text or str(orig.get("section_id") or vals[1])
        elif field == "title":
            vals[2] = text or str(orig.get("title") or vals[2])
        elif field == "layer":
            if text not in LAYER_CHOICES:
                text = str(orig.get("layer") or "节")
            if item != "_preamble" and text == "前言":
                text = "章"
            vals[3] = text
        changed = (
            str(vals[1]) != str(orig.get("section_id", vals[1]))
            or str(vals[2]) != str(orig.get("title", vals[2]))
            or str(vals[3]) != str(orig.get("layer", vals[3]))
        )
        self.preview_tree.item(item, values=vals, tags=("edited",) if changed else tuple(orig.get("tags") or ("normal",)))

    def _collect_section_overrides(self) -> dict[str, dict]:
        overrides: dict[str, dict] = {}
        for item in self.preview_tree.get_children():
            orig = self._preview_original.get(item)
            if not orig:
                continue
            vals = self.preview_tree.item(item, "values")
            new_sid, new_title, new_layer = str(vals[1]).strip(), str(vals[2]).strip(), str(vals[3]).strip()
            new_level = layer_to_level(new_layer, new_sid or item)
            ov: dict = {}
            if new_sid and new_sid != str(orig["section_id"]):
                ov["section_id"] = new_sid
            if new_title and new_title != str(orig["title"]):
                ov["title"] = new_title
            if new_level != int(orig.get("level") or 0):
                ov["level"] = new_level
            if ov:
                if "title" not in ov and new_title:
                    ov["title"] = new_title
                if "level" not in ov:
                    ov["level"] = new_level
                overrides[item] = ov
        return overrides

    def start_preview(self) -> None:
        if not self.files:
            messagebox.showwarning(APP_TITLE, "请先添加至少一份 DOCX 文件。")
            return
        if self.busy:
            return
        # 若启用模板但尚未加载，尝试加载
        if self.use_template_var.get() and self.format_template is None and self.template_path_var.get().strip():
            self.load_template()
        self._set_busy(True, "正在识别章节……")
        self._append_log("开始解析输入文件……")
        self.worker = threading.Thread(target=self._preview_worker, args=(list(self.files),), daemon=True)
        self.worker.start()

    def _preview_worker(self, files: list[str]) -> None:
        staged_root = None
        try:
            staged_root = tempfile.mkdtemp(prefix="insurance-report-preview-")
            input_dir = os.path.join(staged_root, "inputs")
            count, conflicts, _ = copy_selected_files(files, input_dir)
            parsed = parse_input_dir(input_dir)
            # 预览排序：若启用模板，按模板顺序展示
            if self.use_template_var.get() and self.format_template and self.format_template.specs:
                from report_aggregator import sort_section_ids
                order = self.format_template.section_order()
                # build rows in template order then extras
                rows_map = {r["section_id"]: r for r in build_section_preview_rows(parsed)}
                rows = []
                seen = set()
                for sid in order:
                    if sid in rows_map and sid not in seen:
                        rows.append(rows_map[sid])
                        seen.add(sid)
                for r in build_section_preview_rows(parsed):
                    if r["section_id"] not in seen:
                        rows.append(r)
                        seen.add(r["section_id"])
                for i, r in enumerate(rows, 1):
                    r["order"] = i
            else:
                rows = build_section_preview_rows(parsed)
            if not rows:
                raise ValueError("没有识别到可生成的章节")
            self.events.put({
                "kind": "preview_done", "rows": rows,
                "failures": list(getattr(parsed, "failures", []) or []),
                "count": count, "conflicts": conflicts, "staged_root": staged_root,
            })
            staged_root = None
        except Exception as exc:
            self.events.put({"kind": "error", "message": str(exc), "detail": traceback.format_exc()})
        finally:
            if staged_root:
                shutil.rmtree(staged_root, ignore_errors=True)

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择报告保存目录", initialdir=self.vars["output_dir"].get())
        if path:
            self.vars["output_dir"].set(path)

    def _font_config(self) -> FontConfig:
        def bold_of(level: str) -> bool:
            return self.font_vars[f"{level}_bold"].get() == "是"

        def color_of(level: str) -> str:
            return COLOR_OPTIONS.get(self.font_vars[f"{level}_color"].get(), "000000")

        try:
            indent = int(float(self.indent_var.get() or "2"))
        except ValueError:
            indent = 2
        return FontConfig(
            h1_font=self.font_vars["h1_font"].get(), h1_size=float(self.font_vars["h1_size"].get()),
            h1_bold=bold_of("h1"), h1_color=color_of("h1"),
            h2_font=self.font_vars["h2_font"].get(), h2_size=float(self.font_vars["h2_size"].get()),
            h2_bold=bold_of("h2"), h2_color=color_of("h2"),
            h3_font=self.font_vars["h3_font"].get(), h3_size=float(self.font_vars["h3_size"].get()),
            h3_bold=bold_of("h3"), h3_color=color_of("h3"),
            body_font=self.font_vars["body_font"].get(), body_size=float(self.font_vars["body_size"].get()),
            body_bold=bold_of("body"), body_color=color_of("body"), body_indent=indent,
        )

    def start_generation(self) -> None:
        if not self.preview_ready or not self.pending_temp_dir:
            messagebox.showwarning(APP_TITLE, "请先预览章节后再生成。")
            return
        if not self.files:
            messagebox.showwarning(APP_TITLE, "请先添加 DOCX 文件。")
            return
        output_dir = self.vars["output_dir"].get().strip()
        if not output_dir:
            messagebox.showwarning(APP_TITLE, "请选择输出目录。")
            return
        try:
            os.makedirs(output_dir, exist_ok=True)
            if not os.access(output_dir, os.W_OK):
                raise OSError("目录不可写")
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"无法使用输出目录：{exc}")
            return

        if self.use_template_var.get():
            if self.format_template is None:
                if self.template_path_var.get().strip():
                    self.load_template()
                if self.format_template is None:
                    messagebox.showwarning(APP_TITLE, "已勾选格式模板，但尚未成功加载。")
                    return

        paths = build_output_paths(output_dir, self.vars["period"].get())
        demo = os.environ.get("REPORT_DEMO", "").strip() in ("1", "true", "TRUE", "yes")
        existing = [p for p in (paths["docx"], paths["pdf"]) if os.path.exists(p)]
        if existing and not demo:
            if not messagebox.askyesno(APP_TITLE, "以下文件已存在：\n\n" + "\n".join(existing) + "\n\n是否覆盖？"):
                return

        allow_skip = False
        if self.preview_failures:
            details = "\n".join(f"{getattr(f, 'file_name', '?')} · {getattr(f, 'message', '')}" for f in self.preview_failures)
            if demo:
                allow_skip = True
                self._append_log(f"演示模式：跳过无法解析文件\n{details}")
            elif not messagebox.askyesno(APP_TITLE, f"发现无法解析的文件：\n\n{details}\n\n是否跳过继续生成？"):
                return
            else:
                allow_skip = True

        # 字体：启用模板时优先用模板字体，再允许 UI 覆盖（UI 已在 load 时同步）
        font_config = self._font_config()
        ft = self.format_template if self.use_template_var.get() else None
        if ft and ft.font_config:
            # merge: keep UI values already applied from template load
            pass

        options = {
            "files": list(self.files),
            "output_dir": output_dir,
            "period": self.vars["period"].get(),
            "title": self.vars["title"].get(),
            "subtitle": self.vars["subtitle"].get(),
            "org": self.vars["org"].get(),
            "date": self.vars["date"].get(),
            "header_items": [k for k, v in self.header_vars.items() if v.get()],
            "font_config": font_config,
            "allow_skip_failures": allow_skip,
            "section_overrides": self._collect_section_overrides(),
            "section_order": self._collect_section_order(),
            "format_template": ft,
            "use_template_titles": self.use_tpl_titles_var.get() if ft else False,
            "staged_root": self.pending_temp_dir,
        }
        self._set_busy(True, "开始处理报告……")
        self._append_log("开始处理报告……")
        if ft:
            self._append_log(f"使用格式模板：{ft.source_path}（{len(ft.specs)} 章节）")
        self._append_log(f"章节顺序：{len(options['section_order'])} 项")
        self.worker = threading.Thread(target=self._generate_worker, args=(options,), daemon=True)
        self.worker.start()

    def _generate_worker(self, options: dict) -> None:
        pythoncom = None
        try:
            # Windows 上 Word COM 需要在工作线程初始化 COM
            if is_windows():
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except ImportError:
                    pythoncom = None

            preserve = options.get("staged_root")
            if preserve and os.path.isdir(os.path.join(preserve, "inputs")):
                input_dir = os.path.join(preserve, "inputs")
                count, conflicts = len(options["files"]), []
            else:
                input_root = tempfile.mkdtemp(prefix="insurance-report-inputs-")
                input_dir = os.path.join(input_root, "inputs")
                count, conflicts, _ = copy_selected_files(options["files"], input_dir)

            self.events.put({"kind": "log", "message": f"已准备 {count} 份输入文件。"})
            for c in conflicts:
                self.events.put({"kind": "log", "message": f"同名文件：{c['name']} × {c['count']}，已全部保留。"})

            paths = build_output_paths(options["output_dir"], options["period"])
            publish_dir = tempfile.mkdtemp(prefix=".insurance-report-", dir=options["output_dir"])
            temp_docx = os.path.join(publish_dir, "report.docx")
            temp_pdf = os.path.join(publish_dir, "report.pdf")

            self.events.put({"kind": "log", "message": "正在生成 DOCX……"})
            result = aggregate(
                input_dir, temp_docx,
                period=options["period"], title=options["title"], subtitle=options["subtitle"],
                org=options["org"], date=options["date"], header_items=options["header_items"],
                font_config=options["font_config"], silent=True,
                allow_skip_failures=bool(options.get("allow_skip_failures")),
                section_overrides=options.get("section_overrides") or None,
                section_order=options.get("section_order") or None,
                format_template=options.get("format_template"),
                use_template_titles=options.get("use_template_titles"),
            )
            for line in result.get("log", []):
                self.events.put({"kind": "log", "message": line})

            pdf_ok, pdf_msg = _convert_docx_to_pdf_with_word(temp_docx, temp_pdf)
            self.events.put({"kind": "log", "message": pdf_msg})
            shutil.move(temp_docx, paths["docx"])
            if pdf_ok and os.path.exists(temp_pdf):
                shutil.move(temp_pdf, paths["pdf"])
            elif os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            shutil.rmtree(publish_dir, ignore_errors=True)
            self.events.put({
                "kind": "done", "docx": paths["docx"],
                "pdf": paths["pdf"] if pdf_ok and os.path.exists(paths["pdf"]) else "",
                "pdf_msg": pdf_msg,
            })
        except Exception as exc:
            self.events.put({"kind": "error", "message": str(exc), "detail": traceback.format_exc()})
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _process_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event.get("kind")
                if kind == "log":
                    self._append_log(event.get("message", ""))
                elif kind == "preview_done":
                    self._set_busy(False)
                    if self.pending_temp_dir and self.pending_temp_dir != event.get("staged_root"):
                        shutil.rmtree(self.pending_temp_dir, ignore_errors=True)
                    self.pending_temp_dir = event.get("staged_root")
                    self.preview_failures = event.get("failures") or []
                    for item in self.preview_tree.get_children():
                        self.preview_tree.delete(item)
                    self._preview_original.clear()
                    for row in event.get("rows") or []:
                        status = row.get("status", "正常")
                        tag = "normal"
                        if "冲突" in status or "低置信" in status:
                            tag = "warning"
                        elif str(status).startswith("合并"):
                            tag = "merged"
                        elif status == "前言" or row.get("section_id") == "_preamble":
                            tag = "preamble"
                        sid = str(row.get("section_id", ""))
                        title = str(row.get("title", ""))
                        layer = str(row.get("layer", ""))
                        level = int(row.get("level") or layer_to_level(layer, sid))
                        iid = sid if not self.preview_tree.exists(sid) else f"{sid}#{row.get('order')}"
                        self._preview_original[iid] = {
                            "section_id": sid, "title": title, "layer": layer, "level": level,
                            "order": row.get("order", 0), "block_count": row.get("block_count", 0),
                            "source_count": row.get("source_count", 0), "confidence": row.get("confidence", ""),
                            "status": status, "tags": (tag,),
                        }
                        self.preview_tree.insert(
                            "", tk.END, iid=iid,
                            values=(row.get("order", ""), sid, title, layer, row.get("block_count", 0),
                                    row.get("source_count", 0), row.get("confidence", ""), status),
                            tags=(tag,),
                        )
                    nsec = len(event.get("rows") or [])
                    self.preview_ready = True
                    self.generate_btn.configure(state="normal")
                    self.status_var.set(f"预览完成：{nsec} 个章节（可调序/改字段后生成）")
                    self._append_log(f"预览完成：{event.get('count', 0)} 份文件 → {nsec} 个章节。")
                elif kind == "done":
                    self._set_busy(False)
                    self.last_docx = event.get("docx") or None
                    self.last_pdf = event.get("pdf") or None
                    self.open_docx_btn.configure(state="normal" if self.last_docx else "disabled")
                    self.open_pdf_btn.configure(state="normal" if self.last_pdf else "disabled")
                    docx, pdf, pdf_msg = event.get("docx", ""), event.get("pdf", ""), event.get("pdf_msg", "")
                    msg = f"DOCX 已保存：\n{docx}\n\n" + (f"PDF 已保存：\n{pdf}" if pdf else f"PDF 未生成：{pdf_msg}")
                    self.status_var.set("报告生成完成")
                    self._append_log(msg.replace("\n", " | "))
                    # 演示录屏时不弹阻塞对话框，避免打断时间线
                    if os.environ.get("REPORT_DEMO", "").strip() not in ("1", "true", "TRUE", "yes"):
                        messagebox.showinfo(APP_TITLE, msg)
                elif kind == "error":
                    self._set_busy(False)
                    self._append_log(event.get("detail") or event.get("message") or "")
                    self.status_var.set("生成失败，请查看日志")
                    if os.environ.get("REPORT_DEMO", "").strip() not in ("1", "true", "TRUE", "yes"):
                        messagebox.showerror(APP_TITLE, f"生成失败：{event.get('message', '')}")
        except queue.Empty:
            pass
        self.root.after(120, self._process_events)

    def _set_busy(self, busy: bool, status: str = "") -> None:
        self.busy = busy
        self.preview_btn.configure(state="disabled" if busy else "normal")
        if busy:
            self.generate_btn.configure(state="disabled")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.generate_btn.configure(state="normal" if self.preview_ready else "disabled")
        if status:
            self.status_var.set(status)

    def _append_log(self, message: str) -> None:
        self.log.insert(tk.END, str(message).rstrip() + "\n")
        self.log.see(tk.END)

    def _open_path(self, path) -> None:
        if not path or not os.path.exists(path):
            messagebox.showwarning(APP_TITLE, "文件不存在。")
            return
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"无法打开：{exc}")

    def open_output_dir(self) -> None:
        path = self.vars["output_dir"].get()
        if not os.path.isdir(path):
            messagebox.showwarning(APP_TITLE, "输出目录不存在。")
            return
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"无法打开目录：{exc}")

    def on_close(self) -> None:
        if self.busy:
            messagebox.showwarning(APP_TITLE, "任务进行中，请稍候再关闭。")
            return
        if self.pending_temp_dir:
            shutil.rmtree(self.pending_temp_dir, ignore_errors=True)
        self.root.destroy()

    # ------------------------------------------------------------------
    # 演示录屏（真实 GUI 自动操作，非视频生成）
    # ------------------------------------------------------------------
    def _demo_root_dir(self) -> str:
        """定位「报告合成」根目录（含输入样例与模板）。"""
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        here = os.path.dirname(os.path.abspath(__file__))
        return os.path.abspath(os.path.join(here, ".."))

    def _demo_autorun(self) -> None:
        """分步演示：加载 Q3 样例 → 启用模板 → 预览 → 生成。"""
        root_dir = self._demo_root_dir()
        q3 = os.path.join(root_dir, "输入文件_2026Q3")
        tpl = os.path.join(root_dir, "保险行业季度调研报告_终版_2026Q2.docx")
        out = os.path.join(root_dir, "演示输出")
        os.makedirs(out, exist_ok=True)

        self.status_var.set("【演示】准备中：定位样例材料…")
        self._append_log("======== 演示模式开始 ========")
        self._append_log(f"样例目录：{q3}")
        # 后台录制：不抢焦点、不置顶；半隐藏窗口便于 PrintWindow 抓到真实画面
        bg = os.environ.get("REPORT_DEMO_BG", "").strip() in ("1", "true", "TRUE", "yes")
        if bg:
            self.root.update_idletasks()
            w, h = 1220, 820
            try:
                sw = int(self.root.winfo_screenwidth())
                sh = int(self.root.winfo_screenheight())
            except Exception:
                sw, sh = 1920, 1080
            # 挪到屏幕右下角外侧，仅露出极少像素，避免挡住你的工作区
            x = max(sw - 40, 0)
            y = max(sh - 40, 0)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
            try:
                self.root.attributes("-alpha", 0.08)  # 几乎看不见
                self.root.attributes("-topmost", False)
            except Exception:
                pass
            # 不 lift / 不 focus_force
            self._append_log("后台模式：低透明靠边窗口，尽量不打扰前台操作")
        else:
            try:
                self.root.lift()
            except Exception:
                pass

        def step_load():
            files = []
            if os.path.isdir(q3):
                files = sorted(
                    os.path.join(q3, f)
                    for f in os.listdir(q3)
                    if f.lower().endswith(".docx")
                )
            if not files:
                self.status_var.set("【演示】未找到 输入文件_2026Q3，已中止")
                self._append_log("演示失败：无样例 DOCX")
                return
            self.files = files
            self.file_list.delete(0, tk.END)
            for p in files:
                self.file_list.insert(tk.END, os.path.basename(p))
            self.vars["output_dir"].set(out)
            self.vars["period"].set("2026年第三季度")
            self.vars["date"].set("2026年10月")
            if os.path.isfile(tpl):
                self.template_path_var.set(tpl)
                self.use_template_var.set(True)
                try:
                    self.load_template()
                except Exception as exc:
                    self._append_log(f"模板加载提示：{exc}")
            self.status_var.set(f"【演示】已载入 {len(files)} 份 Q3 材料，即将预览章节…")
            self._append_log(f"已载入 {len(files)} 个文件，输出目录：{out}")
            self.root.after(1800, step_preview)

        def step_preview():
            self.status_var.set("【演示】正在识别章节结构…")
            self.start_preview()
            self._demo_wait_preview(0)

        self.root.after(1200, step_load)

    def _demo_wait_preview(self, ticks: int) -> None:
        if self.preview_ready and not self.busy:
            self.status_var.set("【演示】章节预览完成，展示结构后开始生成…")
            # 切到章节预览页
            try:
                for child in self.right.winfo_children():
                    pass
                # notebook 在 _build_workspace 中；通过遍历找 ttk.Notebook
                def find_nb(w):
                    from tkinter import ttk as _ttk
                    if isinstance(w, _ttk.Notebook):
                        return w
                    for c in w.winfo_children():
                        found = find_nb(c)
                        if found:
                            return found
                    return None
                nb = find_nb(self.right)
                if nb is not None and nb.index("end") >= 2:
                    nb.select(1)
            except Exception:
                pass
            self.root.after(2500, self._demo_start_generate)
            return
        if ticks > 120:  # ~60s
            self.status_var.set("【演示】预览超时，已中止")
            return
        self.root.after(500, lambda: self._demo_wait_preview(ticks + 1))

    def _demo_start_generate(self) -> None:
        if not self.preview_ready:
            return
        self.status_var.set("【演示】正在生成终版 DOCX/PDF…")
        self._append_log("演示：开始生成终版报告")
        # 绕过可能的失败确认：清空 preview_failures 以免弹窗卡住录屏
        # 若确有失败仍按正常逻辑
        self.start_generation()
        self._demo_wait_done(0)

    def _demo_wait_done(self, ticks: int) -> None:
        if self.last_docx and not self.busy:
            self.status_var.set("【演示】完成！终版已生成")
            self._append_log(f"演示完成：{self.last_docx}")
            # 写完成标记，供录屏脚本收尾
            try:
                root_dir = self._demo_root_dir()
                marker = os.path.join(root_dir, "演示输出", "_demo_done.txt")
                with open(marker, "w", encoding="utf-8") as f:
                    f.write(self.last_docx + "\n")
                    if self.last_pdf:
                        f.write(self.last_pdf + "\n")
            except Exception:
                pass
            # 再展示几秒供录屏停留
            self.root.after(4000, lambda: None)
            return
        if ticks > 180:  # ~90s
            self.status_var.set("【演示】生成超时")
            return
        self.root.after(500, lambda: self._demo_wait_done(ticks + 1))


def main() -> None:
    root = tk.Tk()
    ReportDesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
