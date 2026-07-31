"""保险行业季度调研报告生成器：跨平台桌面界面（Windows / macOS）。"""

from __future__ import annotations

import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import traceback
from dataclasses import replace
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from app_utils import (
    LAYER_NAMES,
    MAX_HEADING_LEVEL,
    build_output_paths,
    build_section_preview_rows,
    copy_selected_files,
    level_to_layer,
    next_layer_name,
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
from i18n import (
    LANG_LABELS,
    LANG_ZH,
    SUPPORTED,
    color_label_map,
    get_lang,
    hex_to_color_label as _hex_to_color_label,
    is_rtl,
    is_yes,
    layer_display_names,
    layer_display_to_internal,
    layer_internal_to_display,
    load_saved_lang,
    save_lang,
    set_lang,
    t,
    yes_no_values,
)

_DOC_FONTS = default_doc_fonts()
FONT_OPTIONS = list(_DOC_FONTS["options"])
UI_FONT = ui_font_family()
# 常用中文字号（磅）+ 扩展档
SIZE_OPTIONS = [
    "8", "9", "10", "10.5", "11", "12", "13", "14", "15", "16",
    "18", "20", "21", "22", "24", "26", "28", "32", "36", "42", "48",
]
LINE_SPACING_OPTIONS = ["1.0", "1.15", "1.25", "1.5", "1.75", "2.0", "2.5", "3.0"]
EDITABLE_COLS = {"#2": "section_id", "#3": "title", "#4": "layer"}


def _color_options() -> dict:
    return color_label_map()


def _yes_no() -> list:
    return yes_no_values()


def _layer_choices() -> list:
    return layer_display_names() + [t("layer_preamble")]


class ReportDesktopApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        load_saved_lang()
        self.root.title(t("app_title"))
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
        self._deleted_sids: set[str] = set()  # 从预览删除的原解析章节（合成时排除）
        self._edit_widget = self._edit_item = self._edit_field = None
        self.format_template: FormatTemplate | None = None
        self._i18n_widgets: list = []  # (widget, key) 或 (widget, key, "text"/"title")
        self._lang_var = tk.StringVar(value=LANG_LABELS.get(get_lang(), "中文"))

        self.vars: dict[str, tk.StringVar] = {}
        self.header_vars: dict[str, tk.BooleanVar] = {}
        self.font_vars: dict[str, tk.StringVar] = {}
        # 启用模板时：勾选行 = 该级参与覆盖；不勾选 = 该级全跟模板
        self.fmt_override_vars: dict[str, tk.BooleanVar] = {
            k: tk.BooleanVar(value=False)
            for k in ("h1", "h2", "h3", "h4", "h5", "h6", "body", "indent", "spacing")
        }
        # 属性覆盖：仅对「已勾选的行」生效；默认全开 = 等价整行覆盖
        self.fmt_attr_override_vars: dict[str, tk.BooleanVar] = {
            k: tk.BooleanVar(value=True)
            for k in ("font", "size", "bold", "italic", "underline", "color")
        }
        self._fmt_override_checks: list = []  # 行/缩进/行距勾选 + 属性勾选，随模板开关启用
        self._fmt_override_action_btns: list = []  # 一键按钮
        self.indent_var = tk.StringVar(value="2")
        self.line_spacing_var = tk.StringVar(value="1.5")
        self.status_var = tk.StringVar(value=t("status_add_files"))
        self.use_template_var = tk.BooleanVar(value=False)
        self.template_path_var = tk.StringVar(value="")
        self.template_info_var = tk.StringVar(value=t("template_none"))
        self.use_tpl_titles_var = tk.BooleanVar(value=True)

        self._build_style()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(120, self._process_events)
        self._try_default_template()
        # 演示录屏：环境变量 REPORT_DEMO=1 时自动走通「加载样例→预览→生成」
        if os.environ.get("REPORT_DEMO", "").strip() in ("1", "true", "TRUE", "yes"):
            self.root.after(800, self._demo_autorun)

    def _tw(self, widget, key: str) -> None:
        """注册可随语言刷新的控件（configure text）。"""
        self._i18n_widgets.append((widget, key))
        try:
            widget.configure(text=t(key))
        except tk.TclError:
            pass

    def _on_language_change(self, _event=None) -> None:
        label = self._lang_var.get()
        code = LANG_ZH
        for k, lab in LANG_LABELS.items():
            if lab == label:
                code = k
                break
        if code == get_lang():
            return
        # 颜色/是否：切换语言时尽量保留 hex 与真值
        color_hex_by_level = {}
        yn_by_key = {}
        colors_old = _color_options()
        for level in ("h1", "h2", "h3", "h4", "h5", "h6", "body"):
            if f"{level}_color" in self.font_vars:
                color_hex_by_level[level] = colors_old.get(
                    self.font_vars[f"{level}_color"].get(), "000000"
                )
            for attr in ("bold", "italic", "underline"):
                k = f"{level}_{attr}"
                if k in self.font_vars:
                    yn_by_key[k] = is_yes(self.font_vars[k].get())
        set_lang(code)
        save_lang(code)
        self.root.title(t("app_title"))
        # 阿拉伯语/乌尔都语为 RTL（Tk 支持有限，仅作标记）
        try:
            self.root.tk.call("tk", "scaling")  # keep app alive
            if is_rtl(code):
                self.status_var.set(t("status_add_files") if not self.files else self.status_var.get())
        except Exception:
            pass
        # 刷新已注册文案
        for item in self._i18n_widgets:
            w, key = item[0], item[1]
            try:
                w.configure(text=t(key))
            except Exception:
                pass
        # 标题/副标题
        try:
            self._lbl_title.configure(text="📊  " + t("app_title"))
            self._lbl_subtitle.configure(text=t("app_subtitle"))
        except Exception:
            pass
        # 状态若仍是默认提示则更新
        if self.status_var.get() in (
            "请先添加本季度的 DOCX 材料",
            "Please add this quarter’s DOCX materials first",
            "Bitte zuerst die DOCX-Materialien dieses Quartals hinzufügen",
        ) or not self.files:
            if not self.files:
                self.status_var.set(t("status_add_files"))
        if not self.format_template and not self.template_path_var.get().strip():
            self.template_info_var.set(t("template_none"))
        # 更新颜色/是否/层级下拉
        colors = _color_options()
        inv_hex = {v.upper(): k for k, v in colors.items()}
        yn = _yes_no()
        layers = _layer_choices()
        for level in ("h1", "h2", "h3", "h4", "h5", "h6", "body"):
            for attr, width_key in (
                ("font", None), ("size", None),
                ("bold", None), ("italic", None), ("underline", None), ("color", None),
            ):
                pass
            if f"{level}_bold" in self.font_vars:
                self.font_vars[f"{level}_bold"].set(yn[0] if yn_by_key.get(f"{level}_bold") else yn[1])
            if f"{level}_italic" in self.font_vars:
                self.font_vars[f"{level}_italic"].set(yn[0] if yn_by_key.get(f"{level}_italic") else yn[1])
            if f"{level}_underline" in self.font_vars:
                self.font_vars[f"{level}_underline"].set(yn[0] if yn_by_key.get(f"{level}_underline") else yn[1])
            if f"{level}_color" in self.font_vars:
                hx = color_hex_by_level.get(level, "000000").upper()
                self.font_vars[f"{level}_color"].set(inv_hex.get(hx, t("color_black")))
        # 刷新 combobox values
        for cb in getattr(self, "_format_combos", []):
            try:
                kind = cb._i18n_kind  # type: ignore[attr-defined]
                if kind == "yn":
                    cb.configure(values=yn)
                elif kind == "color":
                    cb.configure(values=list(colors.keys()))
            except Exception:
                pass
        # 树表头
        if hasattr(self, "preview_tree"):
            heads = {
                "order": t("tree_order"), "section_id": t("tree_sid"), "title": t("tree_title"),
                "layer": t("tree_layer"), "block_count": t("tree_blocks"),
                "source_count": t("tree_sources"), "confidence": t("tree_conf"), "status": t("tree_status"),
            }
            for c, text in heads.items():
                try:
                    self.preview_tree.heading(c, text=text)
                except Exception:
                    pass
        # Notebook 页签
        if hasattr(self, "_notebook"):
            try:
                self._notebook.tab(0, text=t("tab_files"))
                self._notebook.tab(1, text=t("tab_preview"))
                self._notebook.tab(2, text=t("tab_log"))
            except Exception:
                pass
        # 参数标签
        for key, lbl in getattr(self, "_param_labels", {}).items():
            try:
                lbl.configure(text=t(f"param_{key}"))
            except Exception:
                pass
        for key, cb in getattr(self, "_header_checks", {}).items():
            try:
                cb.configure(text=t(f"hdr_{key}"))
            except Exception:
                pass
        # 格式行名
        for key, lbl in getattr(self, "_fmt_row_labels", {}).items():
            try:
                lbl.configure(text=t(key))
            except Exception:
                pass
        for col, lbl in getattr(self, "_fmt_col_labels", {}).items():
            try:
                lbl.configure(text=t(col))
            except Exception:
                pass
        for item in getattr(self, "_fmt_attr_checks", []) or []:
            try:
                cb, label_key = item
                cb.configure(text=t(label_key))
            except Exception:
                pass
        # 覆盖勾选提示随模板状态刷新（键可能是 on/off 两套）
        try:
            self._update_fmt_override_ui()
        except Exception:
            pass
        try:
            self._schedule_cover_preview()
        except Exception:
            pass
        # 刷新树中层级列显示名
        if hasattr(self, "preview_tree"):
            for item in self.preview_tree.get_children():
                try:
                    vals = list(self.preview_tree.item(item, "values"))
                    if len(vals) > 3:
                        vals[3] = layer_internal_to_display(layer_display_to_internal(str(vals[3])))
                        if len(vals) > 7 and str(vals[7]) in ("新增", "New", "Neu", t("status_new")):
                            vals[7] = t("status_new")
                        self.preview_tree.item(item, values=vals)
                except Exception:
                    pass

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
                self.template_info_var.set(t("template_available", name=os.path.basename(ap)))
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

    def _on_global_mousewheel(self, event) -> None:
        """滚轮作用于当前鼠标所在的左右面板。"""
        canvas = getattr(self, "_active_scroll_canvas", None)
        if canvas is None:
            return
        try:
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass

    def _on_global_linux_scroll(self, event) -> None:
        canvas = getattr(self, "_active_scroll_canvas", None)
        if canvas is None:
            return
        try:
            step = -1 if getattr(event, "num", 0) == 4 else 1
            canvas.yview_scroll(step, "units")
        except tk.TclError:
            pass

    def _bind_panel_mousewheel(self, widget, canvas) -> None:
        """进入该面板任意子控件时，标记滚轮目标为对应 canvas。"""

        def _enter(_event=None, c=canvas) -> None:
            self._active_scroll_canvas = c

        widget.bind("<Enter>", _enter, add="+")
        for child in widget.winfo_children():
            self._bind_panel_mousewheel(child, canvas)

    def _make_scrollable_panel(self, parent: ttk.Frame, *, min_width: int = 0) -> ttk.Frame:
        """在 parent 内创建带纵向滚动条的卡片区域，内容贴顶，返回内层 Frame。"""
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        canvas = tk.Canvas(
            parent,
            highlightthickness=0,
            bd=0,
            bg="#FFFFFF",
            width=min_width or 1,
        )
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        inner = ttk.Frame(canvas, style="Card.TFrame", padding=12)
        # 必须 anchor=nw，内容从左上角贴顶排布，避免上方留白
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _sync_scrollregion(_event=None) -> None:
            # 仅按实际内容定 scrollregion，不要把 canvas 空白算进去
            canvas.update_idletasks()
            bbox = canvas.bbox(window_id)
            if bbox:
                x1, y1, x2, y2 = bbox
                canvas.configure(scrollregion=(0, 0, max(x2, 1), max(y2, 1)))
            else:
                canvas.configure(scrollregion=(0, 0, 1, 1))

        def _sync_inner_width(event) -> None:
            # 只同步宽度，高度严格由内容决定（不把 canvas 高度硬塞给 inner）
            w = max(int(event.width), 1)
            canvas.itemconfigure(window_id, width=w)
            _sync_scrollregion()

        inner.bind("<Configure>", lambda _e: _sync_scrollregion())
        canvas.bind("<Configure>", _sync_inner_width)

        inner._scroll_canvas = canvas  # type: ignore[attr-defined]
        return inner

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, bg="#17365D", height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        top_h = tk.Frame(header, bg="#17365D")
        top_h.pack(fill="x", padx=18, pady=(10, 0))
        self._lbl_title = ttk.Label(top_h, text="📊  " + t("app_title"), style="Title.TLabel")
        self._lbl_title.pack(side="left", anchor="w")
        lang_fr = tk.Frame(top_h, bg="#17365D")
        lang_fr.pack(side="right")
        ttk.Label(lang_fr, text=t("language") + ":", style="Subtitle.TLabel").pack(side="left", padx=(0, 6))
        self._lang_combo = ttk.Combobox(
            lang_fr, textvariable=self._lang_var,
            values=[LANG_LABELS[c] for c in SUPPORTED],
            width=16, state="readonly",
        )
        self._lang_combo.pack(side="left")
        self._lang_combo.bind("<<ComboboxSelected>>", self._on_language_change)
        self._lbl_subtitle = ttk.Label(header, text=t("app_subtitle"), style="Subtitle.TLabel")
        self._lbl_subtitle.pack(anchor="w", padx=22, pady=(2, 10))

        body = ttk.Frame(self.root, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=12, pady=12)

        # 左右可拖动调宽：拖中间竖条调整比例
        paned = tk.PanedWindow(
            body,
            orient=tk.HORIZONTAL,
            sashrelief=tk.RAISED,
            sashwidth=8,
            sashpad=1,
            bd=0,
            bg="#C5D4E8",
            showhandle=False,
        )
        paned.pack(fill="both", expand=True)
        self._main_paned = paned

        left_shell = ttk.Frame(paned, style="Card.TFrame")
        right_shell = ttk.Frame(paned, style="Card.TFrame")
        # minsize 防止一侧拖没；width 为初始左侧宽度
        paned.add(left_shell, minsize=280, width=440, stretch="never")
        paned.add(right_shell, minsize=360, stretch="always")

        right_shell.rowconfigure(0, weight=1)
        right_shell.columnconfigure(0, weight=1)

        self._active_scroll_canvas = None
        self.left = self._make_scrollable_panel(left_shell, min_width=280)
        self.right = ttk.Frame(right_shell, style="Card.TFrame", padding=12)
        self.right.grid(row=0, column=0, sticky="nsew")
        self._build_parameters()
        self._build_workspace()

        # 左侧滚轮；右侧表格/列表用自身滚动条
        self._bind_panel_mousewheel(self.left, self.left._scroll_canvas)  # type: ignore[attr-defined]
        self.root.bind_all("<MouseWheel>", self._on_global_mousewheel)
        self.root.bind_all("<Button-4>", self._on_global_linux_scroll)
        self.root.bind_all("<Button-5>", self._on_global_linux_scroll)
        # 构建完后滚到顶部
        self.root.after_idle(lambda: self.left._scroll_canvas.yview_moveto(0))  # type: ignore[attr-defined]

    def _build_parameters(self) -> None:
        self._format_combos = []
        self._param_labels = {}
        self._header_checks = {}
        self._fmt_row_labels = {}
        self._fmt_col_labels = {}
        self._fmt_attr_checks = []  # (checkbutton, label_key)
        yn = _yes_no()
        colors = _color_options()

        # ---- 格式模板 ----
        w = ttk.Label(self.left, text=t("sec_template"), style="CardTitle.TLabel")
        w.pack(anchor="w")
        self._tw(w, "sec_template")
        cb1 = ttk.Checkbutton(
            self.left, text=t("chk_use_template"),
            variable=self.use_template_var, command=self._on_template_toggle,
        )
        cb1.pack(anchor="w", pady=(6, 2))
        self._tw(cb1, "chk_use_template")
        cb2 = ttk.Checkbutton(
            self.left, text=t("chk_tpl_titles"), variable=self.use_tpl_titles_var,
        )
        cb2.pack(anchor="w", pady=(0, 4))
        self._tw(cb2, "chk_tpl_titles")
        tpl_row = ttk.Frame(self.left, style="Card.TFrame")
        tpl_row.pack(fill="x", pady=2)
        ttk.Entry(tpl_row, textvariable=self.template_path_var).pack(side="left", fill="x", expand=True)
        b1 = ttk.Button(tpl_row, text=t("btn_browse"), command=self.choose_template, style="Action.TButton")
        b1.pack(side="left", padx=(6, 0))
        self._tw(b1, "btn_browse")
        b2 = ttk.Button(tpl_row, text=t("btn_load"), command=self.load_template, style="Action.TButton")
        b2.pack(side="left", padx=(4, 0))
        self._tw(b2, "btn_load")
        ttk.Label(self.left, textvariable=self.template_info_var, style="Body.TLabel").pack(anchor="w", pady=(2, 6))
        ex = ttk.Label(self.left, text=t("tpl_example"), style="Body.TLabel")
        ex.pack(anchor="w")
        self._tw(ex, "tpl_example")

        ttk.Separator(self.left).pack(fill="x", pady=10)
        w = ttk.Label(self.left, text=t("sec_params"), style="CardTitle.TLabel")
        w.pack(anchor="w")
        self._tw(w, "sec_params")
        defaults = {
            "period": "2026年第三季度",
            "title": "中国保险行业调研报告",
            "subtitle": "市场环境 · 保费增长 · 渠道变革 · 趋势展望",
            "org": "保险行业研究中心",
            "date": "2026年10月",
            "output_dir": default_desktop_dir(),
        }
        for key in ("period", "title", "subtitle", "org", "date"):
            lbl = ttk.Label(self.left, text=t(f"param_{key}"), style="Body.TLabel")
            lbl.pack(anchor="w", pady=(8, 0))
            self._param_labels[key] = lbl
            self.vars[key] = tk.StringVar(value=defaults[key])
            ttk.Entry(self.left, textvariable=self.vars[key]).pack(fill="x", pady=2)
            # 参数变更 → 防抖刷新封面示意
            self.vars[key].trace_add("write", lambda *_a, _k=key: self._schedule_cover_preview())

        ttk.Separator(self.left).pack(fill="x", pady=10)
        w = ttk.Label(self.left, text=t("sec_header"), style="CardTitle.TLabel")
        w.pack(anchor="w")
        self._tw(w, "sec_header")
        hf = ttk.Frame(self.left, style="Card.TFrame")
        hf.pack(fill="x", pady=4)
        for key, default in (("title", True), ("org", False), ("date", False), ("period", False)):
            bv = tk.BooleanVar(value=default)
            self.header_vars[key] = bv
            cb = ttk.Checkbutton(hf, text=t(f"hdr_{key}"), variable=bv)
            cb.pack(side="left", padx=(0, 8))
            self._header_checks[key] = cb

        ttk.Separator(self.left).pack(fill="x", pady=10)
        w = ttk.Label(self.left, text=t("sec_format"), style="CardTitle.TLabel")
        w.pack(anchor="w")
        self._tw(w, "sec_format")
        # 模板启用时：勾选行 + 属性 覆盖模板；未启用时勾选无效（全部用界面设置）
        self._fmt_override_hint = ttk.Label(
            self.left, text=t("hint_fmt_override"), style="Body.TLabel", wraplength=320, justify="left",
        )
        self._fmt_override_hint.pack(anchor="w", pady=(2, 0))
        self._tw(self._fmt_override_hint, "hint_fmt_override")

        # 一键：全部跟模板 / 全部跟界面
        fmt_quick = ttk.Frame(self.left, style="Card.TFrame")
        fmt_quick.pack(fill="x", pady=(4, 2))
        b_all_tpl = ttk.Button(
            fmt_quick, text=t("btn_fmt_all_template"),
            command=self._fmt_override_all_template, style="Action.TButton",
        )
        b_all_tpl.pack(side="left", padx=(0, 6))
        self._tw(b_all_tpl, "btn_fmt_all_template")
        self._fmt_override_action_btns.append(b_all_tpl)
        b_all_ui = ttk.Button(
            fmt_quick, text=t("btn_fmt_all_ui"),
            command=self._fmt_override_all_ui, style="Action.TButton",
        )
        b_all_ui.pack(side="left")
        self._tw(b_all_ui, "btn_fmt_all_ui")
        self._fmt_override_action_btns.append(b_all_ui)

        # 属性覆盖：勾选才从界面取该字段（字体/字号/加粗/…）
        attr_fr = ttk.Frame(self.left, style="Card.TFrame")
        attr_fr.pack(fill="x", pady=(2, 2))
        w = ttk.Label(attr_fr, text=t("fmt_attr_override"), style="Body.TLabel")
        w.pack(side="left", padx=(0, 6))
        self._tw(w, "fmt_attr_override")
        for attr, label_key in (
            ("font", "col_font"),
            ("size", "col_size"),
            ("bold", "col_bold"),
            ("italic", "col_italic"),
            ("underline", "col_underline"),
            ("color", "col_color"),
        ):
            cb = ttk.Checkbutton(
                attr_fr,
                text=t(label_key),
                variable=self.fmt_attr_override_vars[attr],
                command=self._on_fmt_override_changed,
            )
            cb.pack(side="left", padx=(0, 4))
            self._fmt_override_checks.append(cb)
            self._fmt_attr_checks.append((cb, label_key))

        fmt_grid = ttk.Frame(self.left, style="Card.TFrame")
        fmt_grid.pack(fill="x", pady=(4, 0))
        # col0=覆盖勾选, col1=行名, col2..7=字体属性
        fmt_grid.columnconfigure(0, minsize=28)
        fmt_grid.columnconfigure(1, minsize=56)
        for col, wdt in ((2, 100), (3, 48), (4, 40), (5, 40), (6, 48), (7, 56)):
            fmt_grid.columnconfigure(col, minsize=wdt, weight=0)
        col_keys = ("col_override", "", "col_font", "col_size", "col_bold", "col_italic", "col_underline", "col_color")
        for col, ckey in enumerate(col_keys):
            text = t(ckey) if ckey else ""
            lbl = ttk.Label(fmt_grid, text=text, style="Body.TLabel", anchor="center" if col != 1 else "w")
            lbl.grid(row=0, column=col, sticky="ew", padx=1, pady=(0, 2))
            if ckey:
                self._fmt_col_labels[ckey] = lbl
        row_defs = (
            ("h1", "fmt_h1"), ("h2", "fmt_h2"), ("h3", "fmt_h3"),
            ("h4", "fmt_h4"), ("h5", "fmt_h5"), ("h6", "fmt_h6"), ("body", "fmt_body"),
        )
        for i, (level, title_key) in enumerate(row_defs, start=1):
            self.font_vars[f"{level}_font"] = tk.StringVar(
                value=_DOC_FONTS["heading"] if level != "body" else _DOC_FONTS["body"]
            )
            self.font_vars[f"{level}_size"] = tk.StringVar(
                value={
                    "h1": "18", "h2": "14", "h3": "12",
                    "h4": "12", "h5": "12", "h6": "11", "body": "12",
                }[level]
            )
            self.font_vars[f"{level}_bold"] = tk.StringVar(value=yn[0] if level != "body" else yn[1])
            self.font_vars[f"{level}_italic"] = tk.StringVar(value=yn[1])
            self.font_vars[f"{level}_underline"] = tk.StringVar(value=yn[1])
            default_color = t("color_black") if level == "body" else t("color_navy")
            self.font_vars[f"{level}_color"] = tk.StringVar(value=default_color)
            ov = ttk.Checkbutton(
                fmt_grid, variable=self.fmt_override_vars[level],
                command=self._on_fmt_override_changed,
            )
            ov.grid(row=i, column=0, sticky="w", padx=(0, 2), pady=2)
            self._fmt_override_checks.append(ov)
            rl = ttk.Label(fmt_grid, text=t(title_key), style="Body.TLabel")
            rl.grid(row=i, column=1, sticky="w", padx=(0, 2), pady=2)
            self._fmt_row_labels[title_key] = rl
            ttk.Combobox(
                fmt_grid, textvariable=self.font_vars[f"{level}_font"],
                values=FONT_OPTIONS, width=10, state="readonly",
            ).grid(row=i, column=2, sticky="ew", padx=1, pady=2)
            ttk.Combobox(
                fmt_grid, textvariable=self.font_vars[f"{level}_size"],
                values=SIZE_OPTIONS, width=5, state="readonly",
            ).grid(row=i, column=3, sticky="ew", padx=1, pady=2)
            for col, attr, kind in (
                (4, "bold", "yn"), (5, "italic", "yn"), (6, "underline", "yn"), (7, "color", "color"),
            ):
                vals = yn if kind == "yn" else list(colors.keys())
                cb = ttk.Combobox(
                    fmt_grid, textvariable=self.font_vars[f"{level}_{attr}"],
                    values=vals, width=5 if kind == "color" else 3, state="readonly",
                )
                cb.grid(row=i, column=col, sticky="ew", padx=1, pady=2)
                cb._i18n_kind = kind  # type: ignore[attr-defined]
                self._format_combos.append(cb)

        ir = ttk.Frame(self.left, style="Card.TFrame")
        ir.pack(fill="x", pady=4)
        ov_i = ttk.Checkbutton(
            ir, variable=self.fmt_override_vars["indent"],
            command=self._on_fmt_override_changed,
        )
        ov_i.pack(side="left")
        self._fmt_override_checks.append(ov_i)
        w = ttk.Label(ir, text=t("body_indent"), style="Body.TLabel")
        w.pack(side="left")
        self._tw(w, "body_indent")
        ttk.Entry(ir, textvariable=self.indent_var, width=6).pack(side="left", padx=6)
        w = ttk.Label(ir, text=t("chars"), style="Body.TLabel")
        w.pack(side="left")
        self._tw(w, "chars")

        ls = ttk.Frame(self.left, style="Card.TFrame")
        ls.pack(fill="x", pady=4)
        ov_s = ttk.Checkbutton(
            ls, variable=self.fmt_override_vars["spacing"],
            command=self._on_fmt_override_changed,
        )
        ov_s.pack(side="left")
        self._fmt_override_checks.append(ov_s)
        w = ttk.Label(ls, text=t("body_spacing"), style="Body.TLabel")
        w.pack(side="left")
        self._tw(w, "body_spacing")
        ttk.Combobox(
            ls, textvariable=self.line_spacing_var, values=LINE_SPACING_OPTIONS,
            width=6, state="readonly",
        ).pack(side="left", padx=6)
        w = ttk.Label(ls, text=t("times"), style="Body.TLabel")
        w.pack(side="left")
        self._tw(w, "times")

        # 初始：未启用模板时覆盖勾选禁用
        self.root.after(0, self._update_fmt_override_ui)

        ttk.Separator(self.left).pack(fill="x", pady=10)
        w = ttk.Label(self.left, text=t("sec_output"), style="CardTitle.TLabel")
        w.pack(anchor="w")
        self._tw(w, "sec_output")
        out_row = ttk.Frame(self.left, style="Card.TFrame")
        out_row.pack(fill="x", pady=4)
        self.vars["output_dir"] = tk.StringVar(value=defaults["output_dir"])
        ttk.Entry(out_row, textvariable=self.vars["output_dir"]).pack(side="left", fill="x", expand=True)
        b = ttk.Button(out_row, text=t("btn_browse_short"), command=self.choose_output_dir, style="Action.TButton")
        b.pack(side="left", padx=(6, 0))
        self._tw(b, "btn_browse_short")

        # ---- 封面预览（在「输出位置」下方；Canvas 仿 Word 封面，不落盘）----
        ttk.Separator(self.left).pack(fill="x", pady=10)
        w = ttk.Label(self.left, text=t("sec_cover_preview"), style="CardTitle.TLabel")
        w.pack(anchor="w")
        self._tw(w, "sec_cover_preview")
        hint = ttk.Label(self.left, text=t("hint_cover_preview"), style="Body.TLabel", wraplength=300)
        hint.pack(anchor="w", pady=(2, 4))
        self._tw(hint, "hint_cover_preview")
        cover_wrap = ttk.Frame(self.left, style="Card.TFrame")
        cover_wrap.pack(fill="x", pady=(0, 4))
        self._cover_shadow = 8
        self._cover_paper_w, self._cover_paper_h = 216, 306
        self._cover_w = self._cover_paper_w + self._cover_shadow + 4
        self._cover_h = self._cover_paper_h + self._cover_shadow + 4
        self._cover_canvas = tk.Canvas(
            cover_wrap,
            width=self._cover_w,
            height=self._cover_h,
            bg="#E8EEF5",
            highlightthickness=0,
            bd=0,
        )
        self._cover_canvas.pack(anchor="center", pady=6)
        self._cover_preview_job = None
        self.root.after_idle(self._draw_cover_preview)

    def _build_workspace(self) -> None:
        # 右栏：顶栏/说明固定，中间 Notebook+表格随窗口拉伸，底栏固定
        self.right.rowconfigure(0, weight=0)
        self.right.rowconfigure(1, weight=0)
        self.right.rowconfigure(2, weight=1)  # 表格区域可拉伸
        self.right.rowconfigure(3, weight=0)
        self.right.columnconfigure(0, weight=1)

        top = ttk.Frame(self.right, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        w = ttk.Label(top, text=t("workspace_title"), style="CardTitle.TLabel")
        w.pack(side="left")
        self._tw(w, "workspace_title")
        b = ttk.Button(top, text=t("btn_add_docx"), command=self.add_files, style="Action.TButton")
        b.pack(side="right", padx=3)
        self._tw(b, "btn_add_docx")
        self.preview_btn = ttk.Button(top, text=t("btn_preview"), command=self.start_preview, style="Action.TButton")
        self.preview_btn.pack(side="right", padx=3)
        self._tw(self.preview_btn, "btn_preview")
        b = ttk.Button(top, text=t("btn_clear"), command=self.clear_files, style="Action.TButton")
        b.pack(side="right", padx=3)
        self._tw(b, "btn_clear")
        b = ttk.Button(top, text=t("btn_remove"), command=self.remove_selected, style="Action.TButton")
        b.pack(side="right", padx=3)
        self._tw(b, "btn_remove")

        mid = ttk.Frame(self.right, style="Card.TFrame")
        mid.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        w = ttk.Label(mid, text=t("hint_edit"), style="Body.TLabel")
        w.pack(side="left")
        self._tw(w, "hint_edit")
        b = ttk.Button(mid, text=t("btn_reset_preview"), command=self.reset_preview_edits, style="Action.TButton")
        b.pack(side="right")
        self._tw(b, "btn_reset_preview")

        notebook = ttk.Notebook(self.right)
        notebook.grid(row=2, column=0, sticky="nsew", pady=8)
        self._notebook = notebook
        file_tab = ttk.Frame(notebook, style="Card.TFrame")
        preview_tab = ttk.Frame(notebook, style="Card.TFrame")
        log_tab = ttk.Frame(notebook, style="Card.TFrame")
        notebook.add(file_tab, text=t("tab_files"))
        notebook.add(preview_tab, text=t("tab_preview"))
        notebook.add(log_tab, text=t("tab_log"))

        file_wrap = ttk.Frame(file_tab, style="Card.TFrame")
        file_wrap.pack(fill="both", expand=True, padx=4, pady=4)
        file_wrap.rowconfigure(0, weight=1)
        file_wrap.columnconfigure(0, weight=1)
        self.file_list = tk.Listbox(file_wrap, activestyle="dotbox", font=("Consolas", 10))
        file_ys = ttk.Scrollbar(file_wrap, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=file_ys.set)
        self.file_list.grid(row=0, column=0, sticky="nsew")
        file_ys.grid(row=0, column=1, sticky="ns")

        wrap = ttk.Frame(preview_tab, style="Card.TFrame")
        wrap.pack(fill="both", expand=True)
        wrap.columnconfigure(0, weight=1)
        wrap.columnconfigure(1, weight=0)
        wrap.rowconfigure(0, weight=1)
        tree_frame = ttk.Frame(wrap, style="Card.TFrame")
        tree_frame.grid(row=0, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        cols = ("order", "section_id", "title", "layer", "block_count", "source_count", "confidence", "status")
        # extended：支持 Ctrl/Shift 多选，便于批量删除/隐藏/改层级
        self.preview_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")
        heads = {
            "order": t("tree_order"), "section_id": t("tree_sid"), "title": t("tree_title"),
            "layer": t("tree_layer"), "block_count": t("tree_blocks"),
            "source_count": t("tree_sources"), "confidence": t("tree_conf"), "status": t("tree_status"),
        }
        # 固定窄列 + 标题列随窗口拉伸
        widths = {"order": 50, "section_id": 90, "title": 260, "layer": 70, "block_count": 60,
                  "source_count": 60, "confidence": 60, "status": 80}
        stretch_cols = {"title"}
        for c in cols:
            self.preview_tree.heading(c, text=heads[c])
            self.preview_tree.column(
                c,
                width=widths[c],
                minwidth=40 if c != "title" else 120,
                stretch=(c in stretch_cols),
                anchor="center" if c != "title" else "w",
            )
        ys = ttk.Scrollbar(tree_frame, orient="vertical", command=self.preview_tree.yview)
        xs = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.preview_tree.tag_configure("warning", foreground="#C00000")
        self.preview_tree.tag_configure("merged", foreground="#9A6700")
        self.preview_tree.tag_configure("preamble", foreground="#1F3864")
        self.preview_tree.tag_configure("edited", foreground="#0B6E4F")
        self.preview_tree.tag_configure("manual", foreground="#2E75B6")
        self.preview_tree.tag_configure("hidden", foreground="#94A3B8")
        self.preview_tree.bind("<Double-1>", self._on_tree_double_click)
        self.preview_tree.bind("<Control-Up>", lambda e: self._move_selected(-1))
        self.preview_tree.bind("<Control-Down>", lambda e: self._move_selected(1))
        self.preview_tree.bind("<Delete>", lambda e: self._delete_selected_sections())

        order_btns = ttk.Frame(wrap, style="Card.TFrame", padding=(6, 0))
        order_btns.grid(row=0, column=1, sticky="ns")
        w = ttk.Label(order_btns, text=t("order_panel"), style="CardTitle.TLabel")
        w.pack(pady=(0, 6))
        self._tw(w, "order_panel")
        for key, cmd in (
            ("btn_up", lambda: self._move_selected(-1)),
            ("btn_down", lambda: self._move_selected(1)),
            ("btn_top", lambda: self._move_selected("top")),
            ("btn_bottom", lambda: self._move_selected("bottom")),
        ):
            b = ttk.Button(order_btns, text=t(key), command=cmd, style="Action.TButton")
            b.pack(fill="x", pady=2)
            self._tw(b, key)
        ttk.Label(order_btns, text="Ctrl+↑/↓", style="Body.TLabel").pack(pady=(10, 0))
        ttk.Separator(order_btns, orient="horizontal").pack(fill="x", pady=10)
        w = ttk.Label(order_btns, text=t("struct_panel"), style="CardTitle.TLabel")
        w.pack(pady=(0, 6))
        self._tw(w, "struct_panel")
        b = ttk.Button(
            order_btns, text=t("btn_add_heading"),
            command=self._add_subheading, style="Action.TButton",
        )
        b.pack(fill="x", pady=2)
        self._tw(b, "btn_add_heading")
        b = ttk.Button(
            order_btns, text=t("btn_delete_section"),
            command=self._delete_selected_sections, style="Action.TButton",
        )
        b.pack(fill="x", pady=2)
        self._tw(b, "btn_delete_section")
        b = ttk.Button(
            order_btns, text=t("btn_toggle_hide"),
            command=self._toggle_hide_selected, style="Action.TButton",
        )
        b.pack(fill="x", pady=2)
        self._tw(b, "btn_toggle_hide")
        b = ttk.Button(
            order_btns, text=t("btn_batch_structure"),
            command=self._batch_structure_dialog, style="Action.TButton",
        )
        b.pack(fill="x", pady=2)
        self._tw(b, "btn_batch_structure")
        w = ttk.Label(order_btns, text=t("hint_heading_only"), style="Body.TLabel", justify="center")
        w.pack(pady=(6, 0))
        self._tw(w, "hint_heading_only")
        w = ttk.Label(order_btns, text=t("hint_multi_select"), style="Body.TLabel", justify="center")
        w.pack(pady=(4, 0))
        self._tw(w, "hint_multi_select")

        self.log = ScrolledText(log_tab, height=12, font=("Consolas", 9), relief="flat", bg="#F8FAFC")
        self.log.pack(fill="both", expand=True, padx=4, pady=4)

        bottom = ttk.Frame(self.right, style="Card.TFrame")
        bottom.grid(row=3, column=0, sticky="ew")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 6))
        ttk.Label(bottom, textvariable=self.status_var, style="Body.TLabel").pack(anchor="w")
        btn_row = ttk.Frame(bottom, style="Card.TFrame")
        btn_row.pack(fill="x", pady=6)
        self.generate_btn = ttk.Button(
            btn_row, text=t("btn_generate"), command=self.start_generation,
            style="Primary.TButton", state="disabled",
        )
        self.generate_btn.pack(side="left")
        self._tw(self.generate_btn, "btn_generate")
        self.open_docx_btn = ttk.Button(
            btn_row, text=t("btn_open_docx"), command=lambda: self._open_path(self.last_docx),
            style="Action.TButton", state="disabled",
        )
        self.open_docx_btn.pack(side="left", padx=6)
        self._tw(self.open_docx_btn, "btn_open_docx")
        self.open_pdf_btn = ttk.Button(
            btn_row, text=t("btn_open_pdf"), command=lambda: self._open_path(self.last_pdf),
            style="Action.TButton", state="disabled",
        )
        self.open_pdf_btn.pack(side="left", padx=6)
        self._tw(self.open_pdf_btn, "btn_open_pdf")
        b = ttk.Button(btn_row, text=t("btn_open_outdir"), command=self.open_output_dir, style="Action.TButton")
        b.pack(side="left", padx=6)
        self._tw(b, "btn_open_outdir")

    # ---- cover preview (Canvas, mirrors Word add_cover layout; no disk) ----
    def _schedule_cover_preview(self) -> None:
        """参数变更后防抖重绘封面示意。"""
        job = getattr(self, "_cover_preview_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        try:
            self._cover_preview_job = self.root.after(280, self._draw_cover_preview)
        except Exception:
            pass

    def _cover_heading_color(self) -> str:
        """与终版封面一致：标题/周期用深蓝（FontConfig.h2 → COLOR_HEADING）。"""
        ft = self.format_template if self.use_template_var.get() else None
        if ft and ft.font_config:
            for attr in ("h2_color", "h1_color"):
                hx = str(getattr(ft.font_config, attr, "") or "").strip().lstrip("#")
                if len(hx) == 6:
                    try:
                        int(hx, 16)
                        return f"#{hx.upper()}"
                    except ValueError:
                        pass
        return "#1F3864"

    @staticmethod
    def _pick_ui_font(candidates: list[str]) -> str:
        """选本机可用的中文字体。"""
        for name in candidates:
            if not name:
                continue
            try:
                f = tkfont.Font(family=name, size=10)
                actual = (f.actual("family") or "").replace(" ", "").lower()
                want = name.replace(" ", "").lower()
                # 宽松匹配：系统可能返回 "Microsoft YaHei UI"
                if want in actual or actual in want or name.lower() in (f.actual("family") or "").lower():
                    return f.actual("family") or name
            except Exception:
                continue
        return UI_FONT

    def _cover_font_families(self) -> tuple[str, str]:
        heading = self._pick_ui_font([
            "微软雅黑", "Microsoft YaHei", "Microsoft YaHei UI",
            "黑体", "SimHei", "思源黑体", "Source Han Sans SC", UI_FONT,
        ])
        body = self._pick_ui_font([
            "微软雅黑", "Microsoft YaHei", "Microsoft YaHei UI",
            "宋体", "SimSun", "Noto Serif CJK SC", UI_FONT,
        ])
        return heading, body

    @staticmethod
    def _wrap_canvas_text(
        text: str, font: tuple, max_width: int, *, max_lines: int = 8,
    ) -> list[str]:
        """按像素宽度折行；仅当内容超出 max_lines 时在末行加省略号。"""
        text = (text or "").strip()
        if not text:
            return []
        try:
            fnt = tkfont.Font(font=font)
            measure = fnt.measure
        except Exception:
            avg = max(8, int(font[1]) if len(font) > 1 else 10)

            def measure(s: str, _a=avg) -> int:
                return len(s) * _a

        lines: list[str] = []
        truncated = False
        for para in text.replace("\r", "").split("\n"):
            para = para.strip()
            if not para:
                continue
            current = ""
            for ch in para:
                trial = current + ch
                if measure(trial) <= max_width or not current:
                    current = trial
                else:
                    if len(lines) + 1 >= max_lines:
                        # 当前行写满后还要换行 → 截断
                        lines.append(current)
                        truncated = True
                        current = ""
                        break
                    lines.append(current)
                    current = ch
            if truncated:
                break
            if current:
                if len(lines) >= max_lines:
                    truncated = True
                    break
                lines.append(current)
        if truncated and lines:
            last = lines[-1]
            while last and measure(last + "…") > max_width:
                last = last[:-1]
            lines[-1] = (last + "…") if last else "…"
        return lines

    def _draw_cover_preview(self) -> None:
        """仿 report_aggregator.add_cover 的版式：顶留白 → 周期/标题 → 副标题 → 大留白 → 单位/日期。"""
        self._cover_preview_job = None
        cv = getattr(self, "_cover_canvas", None)
        if cv is None:
            return
        try:
            cv.delete("all")
        except tk.TclError:
            return

        sh = int(getattr(self, "_cover_shadow", 8))
        pw = int(getattr(self, "_cover_paper_w", 216))
        ph = int(getattr(self, "_cover_paper_h", 306))
        ox, oy = 2, 2  # 纸面左上（阴影向右下）

        # 背景
        cv.create_rectangle(0, 0, self._cover_w, self._cover_h, fill="#E8EEF5", outline="")

        # 多层软阴影
        for i, col in enumerate(("#C5D0DE", "#D2DBE8", "#DEE5EF")):
            d = sh - i * 2
            if d <= 0:
                break
            cv.create_rectangle(
                ox + d, oy + d, ox + pw + d, oy + ph + d,
                fill=col, outline="",
            )
        # 纸面
        cv.create_rectangle(ox, oy, ox + pw, oy + ph, fill="#FFFFFF", outline="#B8C4D4", width=1)
        # 内细边（印刷页边感）
        inset = 5
        cv.create_rectangle(
            ox + inset, oy + inset, ox + pw - inset, oy + ph - inset,
            fill="", outline="#E8EEF4", width=1,
        )

        period = (self.vars.get("period").get() if self.vars.get("period") else "") or ""
        title = (self.vars.get("title").get() if self.vars.get("title") else "") or ""
        subtitle = (self.vars.get("subtitle").get() if self.vars.get("subtitle") else "") or ""
        org = (self.vars.get("org").get() if self.vars.get("org") else "") or ""
        date = (self.vars.get("date").get() if self.vars.get("date") else "") or ""

        heading_family, body_family = self._cover_font_families()
        color_heading = self._cover_heading_color()
        color_sub = "#666666"
        color_meta = "#333333"

        # 字号相对纸宽缩放（对照 Word：标题约 26pt、副文 14pt）
        # 迷你纸上适当放大以保证可读
        title_size = 13
        sub_size = 9
        meta_size = 9
        font_title = (heading_family, title_size, "bold")
        font_sub = (body_family, sub_size)
        font_meta = (body_family, meta_size)

        pad_x = 18
        max_tw = pw - pad_x * 2
        cx = ox + pw // 2

        # —— 版式比例对齐 add_cover ——
        # space_before≈120pt、标题间距、副标题后大段空白、底部单位日期
        y = oy + int(ph * 0.16)

        # 周期（与标题同级：黑体粗、深蓝、居中）
        period_lines = self._wrap_canvas_text(period.strip(), font_title, max_tw, max_lines=2)
        line_h_title = title_size + 8
        for line in period_lines:
            cv.create_text(cx, y, text=line, font=font_title, fill=color_heading, anchor="n")
            y += line_h_title
        if period_lines:
            y += 4  # space_after≈8pt 缩放

        # 主标题
        title_lines = self._wrap_canvas_text(title.strip() or "—", font_title, max_tw, max_lines=4)
        if not title_lines:
            title_lines = ["—"]
        for line in title_lines:
            cv.create_text(cx, y, text=line, font=font_title, fill=color_heading, anchor="n")
            y += line_h_title
        y += 14  # space_after≈36pt 缩放

        # 副标题（灰、略小）
        line_h_sub = sub_size + 6
        sub_lines = self._wrap_canvas_text(subtitle.strip(), font_sub, max_tw, max_lines=3)
        for line in sub_lines:
            cv.create_text(cx, y, text=line, font=font_sub, fill=color_sub, anchor="n")
            y += line_h_sub

        # 大留白（对应 space_after≈180pt）— 单位/日期贴底
        bot_block_h = 48
        y_bot = oy + ph - bot_block_h

        # 若上方文字侵入底部，压缩间距（极端长标题）
        if y + 20 > y_bot:
            y_bot = min(oy + ph - 36, y + 16)

        # 底部细线装饰（轻量，不抢真封面）
        cv.create_line(
            ox + pad_x + 24, y_bot - 12,
            ox + pw - pad_x - 24, y_bot - 12,
            fill="#D0D7E2",
        )

        org_lines = self._wrap_canvas_text(org.strip() or "—", font_meta, max_tw, max_lines=2)
        date_lines = self._wrap_canvas_text(date.strip() or "—", font_meta, max_tw, max_lines=1)
        yy = y_bot
        for line in org_lines:
            cv.create_text(cx, yy, text=line, font=font_meta, fill=color_meta, anchor="n")
            yy += meta_size + 5
        yy += 2
        for line in date_lines:
            cv.create_text(cx, yy, text=line, font=font_meta, fill=color_meta, anchor="n")
            yy += meta_size + 5

    # ---- template ----
    def _on_template_toggle(self) -> None:
        if self.use_template_var.get() and self.template_path_var.get().strip():
            self.load_template()
        self._update_fmt_override_ui()
        self._schedule_cover_preview()

    def _on_fmt_override_changed(self) -> None:
        """勾选变化时在状态栏提示当前覆盖范围。"""
        if not self.use_template_var.get() or not self.format_template:
            return
        rows = [k for k, v in self.fmt_override_vars.items() if v.get()]
        attrs = [k for k, v in self.fmt_attr_override_vars.items() if v.get()]
        if not rows:
            self.status_var.set(t("status_fmt_all_template"))
            return
        attr_s = ",".join(attrs) if attrs else t("fmt_attr_none")
        self.status_var.set(t("status_fmt_override", keys=", ".join(rows), attrs=attr_s))

    def _fmt_override_all_template(self) -> None:
        """一键：所有级别/缩进/行距/属性都不覆盖 → 全部跟模板。"""
        if not self.use_template_var.get() or not self.format_template:
            messagebox.showinfo(t("app_title"), t("msg_fmt_need_template"))
            return
        for v in self.fmt_override_vars.values():
            v.set(False)
        # 属性勾选一并取消，避免界面仍显示「字号/颜色…」为选中态
        for v in self.fmt_attr_override_vars.values():
            v.set(False)
        self.status_var.set(t("status_fmt_all_template"))
        self._append_log("format override: all → template (rows+attrs cleared)")

    def _fmt_override_all_ui(self) -> None:
        """一键：所有级别 + 全部属性 + 缩进/行距 → 全部跟界面。"""
        if not self.use_template_var.get() or not self.format_template:
            messagebox.showinfo(t("app_title"), t("msg_fmt_need_template"))
            return
        for v in self.fmt_override_vars.values():
            v.set(True)
        for v in self.fmt_attr_override_vars.values():
            v.set(True)
        self.status_var.set(
            t("status_fmt_override", keys="h1–h6,body,indent,spacing", attrs="all")
        )
        self._append_log("format override: all → UI (all attributes)")

    def _update_fmt_override_ui(self) -> None:
        """启用格式模板时允许行/属性覆盖与一键按钮；否则禁用。"""
        active = bool(self.use_template_var.get() and self.format_template is not None)
        state = "normal" if active else "disabled"
        for w in getattr(self, "_fmt_override_checks", []):
            try:
                w.configure(state=state)
            except tk.TclError:
                pass
        for w in getattr(self, "_fmt_override_action_btns", []):
            try:
                w.configure(state=state)
            except tk.TclError:
                pass
        hint = getattr(self, "_fmt_override_hint", None)
        if hint is not None:
            try:
                hint.configure(text=t("hint_fmt_override" if active else "hint_fmt_override_off"))
            except tk.TclError:
                pass

    def choose_template(self) -> None:
        cur = self.template_path_var.get().strip()
        initial = os.path.dirname(cur) if cur else default_desktop_dir()
        path = filedialog.askopenfilename(
            title=t("dlg_pick_template"),
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            initialdir=initial if os.path.isdir(initial) else default_desktop_dir(),
        )
        if path:
            self.template_path_var.set(path)
            self.load_template()

    def load_template(self) -> None:
        path = self.template_path_var.get().strip()
        if not path:
            messagebox.showwarning(t("app_title"), t("msg_tpl_path"))
            return
        if not os.path.isfile(path):
            messagebox.showerror(t("app_title"), t("msg_file_missing", path=path))
            return
        try:
            ft = load_format_template(path)
        except Exception as exc:
            messagebox.showerror(t("app_title"), t("msg_tpl_load_fail", err=exc))
            self.format_template = None
            self._update_fmt_override_ui()
            return
        self.format_template = ft
        self.use_template_var.set(True)
        self._update_fmt_override_ui()
        self._schedule_cover_preview()
        self.template_info_var.set(
            f"已加载：{os.path.basename(path)} · {len(ft.specs)} 个章节 · "
            f"H1 {ft.font_config.h1_font}/{ft.font_config.h1_size}/#{ft.font_config.h1_color} · "
            f"正文 #{ft.font_config.body_color}"
        )
        self._append_log(f"已加载格式模板：{path}")
        self._append_log(f"  模板章节 {len(ft.specs)} 个；封面「{ft.title}」/ {ft.period}")
        if ft.font_config:
            fc = ft.font_config
            self._append_log(
                f"  版式（合成时严格按模板）："
                f" H1=#{fc.h1_color} H2=#{fc.h2_color} H3=#{fc.h3_color} 正文=#{fc.body_color}"
            )
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
        # 应用模板完整版式到 UI，便于预览；合成时仍直接用模板 FontConfig
        if ft.font_config:
            fc = ft.font_config
            mapping = [
                ("h1", fc.h1_font, fc.h1_size, fc.h1_bold, fc.h1_italic, fc.h1_underline, fc.h1_color),
                ("h2", fc.h2_font, fc.h2_size, fc.h2_bold, fc.h2_italic, fc.h2_underline, fc.h2_color),
                ("h3", fc.h3_font, fc.h3_size, fc.h3_bold, fc.h3_italic, fc.h3_underline, fc.h3_color),
                ("h4", fc.h4_font, fc.h4_size, fc.h4_bold, fc.h4_italic, fc.h4_underline, fc.h4_color),
                ("h5", fc.h5_font, fc.h5_size, fc.h5_bold, fc.h5_italic, fc.h5_underline, fc.h5_color),
                ("h6", fc.h6_font, fc.h6_size, fc.h6_bold, fc.h6_italic, fc.h6_underline, fc.h6_color),
                ("body", fc.body_font, fc.body_size, fc.body_bold, fc.body_italic, fc.body_underline, fc.body_color),
            ]
            for level, font, size, bold, italic, under, color in mapping:
                if font:
                    if font not in FONT_OPTIONS:
                        FONT_OPTIONS.append(font)
                    self.font_vars[f"{level}_font"].set(font)
                if size:
                    s = str(int(size) if float(size).is_integer() else size)
                    if s not in SIZE_OPTIONS:
                        SIZE_OPTIONS.append(s)
                    self.font_vars[f"{level}_size"].set(s)
                yn = _yes_no()
                self.font_vars[f"{level}_bold"].set(yn[0] if bold else yn[1])
                self.font_vars[f"{level}_italic"].set(yn[0] if italic else yn[1])
                self.font_vars[f"{level}_underline"].set(yn[0] if under else yn[1])
                self.font_vars[f"{level}_color"].set(_hex_to_color_label(color))
            # 行距
            ls = getattr(fc, "body_line_spacing", None)
            if ls is not None:
                s = str(ls)
                if s.endswith(".0"):
                    s = s[:-2] if ls in (1.0, 2.0, 3.0) else str(ls)
                # 规范为选项中的写法
                s_norm = f"{float(ls):g}"
                if s_norm not in LINE_SPACING_OPTIONS:
                    LINE_SPACING_OPTIONS.append(s_norm)
                    LINE_SPACING_OPTIONS.sort(key=float)
                self.line_spacing_var.set(s_norm if s_norm in LINE_SPACING_OPTIONS else str(ls))
        self.status_var.set(t("status_template_ready", n=len(ft.specs)))

    # ---- files ----
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title=t("dlg_pick_docx"), filetypes=[(t("filetype_docx"), "*.docx")]
        )
        for p in paths:
            ap = os.path.abspath(p)
            if ap not in self.files:
                self.files.append(ap)
                self.file_list.insert(tk.END, ap)
        self._invalidate_preview()
        self.status_var.set(t("status_files_selected", n=len(self.files)))

    def remove_selected(self) -> None:
        for idx in reversed(self.file_list.curselection()):
            self.file_list.delete(idx)
            del self.files[idx]
        self._invalidate_preview()
        self.status_var.set(t("status_files_selected", n=len(self.files)))

    def clear_files(self) -> None:
        self.files.clear()
        self.file_list.delete(0, tk.END)
        self._invalidate_preview()
        self.status_var.set(t("status_add_files"))

    def _invalidate_preview(self) -> None:
        self.preview_ready = False
        self.preview_failures = []
        self._preview_original.clear()
        self._deleted_sids.clear()
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.generate_btn.configure(state="disabled")
        if self.pending_temp_dir:
            shutil.rmtree(self.pending_temp_dir, ignore_errors=True)
            self.pending_temp_dir = None
        self._close_edit_widget(save=False)

    def reset_preview_edits(self) -> None:
        if not self.preview_tree.get_children() and not self._preview_original:
            return
        rows = sorted(
            ((orig.get("order", 0), iid, orig) for iid, orig in self._preview_original.items()),
            key=lambda x: x[0],
        )
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self._deleted_sids.clear()
        for order, iid, orig in rows:
            layer_disp = layer_internal_to_display(str(orig.get("layer") or "节"))
            self.preview_tree.insert(
                "", tk.END, iid=iid,
                values=(order, orig["section_id"], orig["title"], layer_disp,
                        orig.get("block_count", ""), orig.get("source_count", ""),
                        orig.get("confidence", ""), orig.get("status", t("status_normal"))),
                tags=orig.get("tags", ("normal",)),
            )
        self._renumber_order()
        self.status_var.set(t("status_reset_preview"))

    def _move_selected(self, direction) -> str:
        if self.busy:
            return "break"
        sel = self.preview_tree.selection()
        if not sel:
            messagebox.showinfo(t("app_title"), t("msg_select_row"))
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
        self.status_var.set(t("status_order", n=new_idx + 1))
        return "break"

    def _existing_section_ids(self) -> set[str]:
        ids: set[str] = set()
        for item in self.preview_tree.get_children():
            vals = self.preview_tree.item(item, "values")
            if vals and len(vals) > 1:
                sid = str(vals[1]).strip()
                if sid:
                    ids.add(sid)
            ids.add(str(item))
        return ids

    def _suggest_child_section_id(self, parent_sid: str) -> str:
        """根据父编号生成下一个未占用的子编号，如 2 → 2.1，2.1 → 2.1.1。"""
        parent_sid = (parent_sid or "").strip()
        if not parent_sid or parent_sid.startswith("_") or parent_sid == "_preamble":
            # 无父级时给一级编号
            used = self._existing_section_ids()
            n = 1
            while str(n) in used:
                n += 1
            return str(n)
        used = self._existing_section_ids()
        n = 1
        while True:
            cand = f"{parent_sid}.{n}"
            if cand not in used:
                return cand
            n += 1
            if n > 200:
                return f"{parent_sid}.{n}"

    def _add_subheading(self) -> None:
        """在章节预览中新增一个仅标题、无正文的条目（可插入到任意位置）。"""
        if self.busy:
            return
        if not self.preview_ready and not self.preview_tree.get_children():
            messagebox.showwarning(t("app_title"), t("msg_preview_first"))
            return

        # 先固定当前列表与选中项（点按钮后焦点变化也可能丢选中）
        children = list(self.preview_tree.get_children())
        sel = list(self.preview_tree.selection())
        selected_item = sel[0] if sel else None

        parent_sid = ""
        default_layer = "章"
        if selected_item:
            vals = list(self.preview_tree.item(selected_item, "values"))
            parent_sid = str(vals[1]).strip() if len(vals) > 1 else str(selected_item)
            parent_layer = layer_display_to_internal(
                str(vals[3]).strip() if len(vals) > 3 else "章"
            )
            if parent_layer == "前言":
                default_layer = "章"
                parent_sid = ""
            elif parent_layer in LAYER_NAMES:
                if parent_layer == LAYER_NAMES[-1]:
                    default_layer = parent_layer
                    if "." in parent_sid:
                        parent_sid = parent_sid.rsplit(".", 1)[0]
                else:
                    default_layer = next_layer_name(parent_layer)
            else:
                default_layer = "节"

        if default_layer == "章":
            suggest_sid = self._suggest_child_section_id("")
        else:
            suggest_sid = self._suggest_child_section_id(parent_sid)

        # 插入位置选项：最前 / 在某行之后 / 最后
        pos_choices: list[tuple[str, str, str | None]] = [
            (t("pos_first"), "start", None),
        ]
        for item in children:
            vals = list(self.preview_tree.item(item, "values"))
            order = str(vals[0]) if vals else ""
            sid = str(vals[1]).strip() if len(vals) > 1 else item
            title = str(vals[2]).strip() if len(vals) > 2 else ""
            short = title if len(title) <= 18 else title[:18] + "…"
            label = t("pos_after", order=order, sid=sid, title=short)
            pos_choices.append((label, "after", item))
        pos_choices.append((t("pos_last"), "end", None))

        default_pos_label = pos_choices[-1][0]
        if selected_item:
            for lab, mode, ref in pos_choices:
                if mode == "after" and ref == selected_item:
                    default_pos_label = lab
                    break

        dlg = tk.Toplevel(self.root)
        dlg.title(t("dlg_add_heading"))
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=t("dlg_add_hint"), style="Body.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        ttk.Label(frm, text=t("dlg_insert_pos"), style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        pos_var = tk.StringVar(value=default_pos_label)
        ttk.Combobox(
            frm, textvariable=pos_var, values=[c[0] for c in pos_choices],
            width=36, state="readonly",
        ).grid(row=1, column=1, sticky="ew", pady=4, padx=(8, 0))

        ttk.Label(frm, text=t("dlg_sid"), style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        sid_var = tk.StringVar(value=suggest_sid)
        ttk.Entry(frm, textvariable=sid_var, width=22).grid(row=2, column=1, sticky="ew", pady=4, padx=(8, 0))

        ttk.Label(frm, text=t("dlg_title"), style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=4)
        title_var = tk.StringVar(value=t("default_new_title"))
        title_entry = ttk.Entry(frm, textvariable=title_var, width=22)
        title_entry.grid(row=3, column=1, sticky="ew", pady=4, padx=(8, 0))

        ttk.Label(frm, text=t("dlg_layer"), style="Body.TLabel").grid(row=4, column=0, sticky="w", pady=4)
        disp_layers = layer_display_names()
        default_disp = layer_internal_to_display(
            default_layer if default_layer in LAYER_NAMES else "节"
        )
        layer_var = tk.StringVar(value=default_disp)
        ttk.Combobox(
            frm, textvariable=layer_var, values=disp_layers,
            width=20, state="readonly",
        ).grid(row=4, column=1, sticky="ew", pady=4, padx=(8, 0))
        frm.columnconfigure(1, weight=1)

        result: dict = {}
        pos_map = {lab: (mode, ref) for lab, mode, ref in pos_choices}

        def _ok() -> None:
            sid = sid_var.get().strip()
            title = title_var.get().strip() or t("default_new_title")
            layer = layer_display_to_internal(layer_var.get().strip() or t("layer_2"))
            if layer == "前言":
                layer = "章"
            if not sid or sid.startswith("_"):
                messagebox.showwarning(t("app_title"), t("msg_need_sid"), parent=dlg)
                return
            if not re.fullmatch(r"\d+(?:\.\d+)*", sid):
                messagebox.showwarning(t("app_title"), t("msg_sid_format"), parent=dlg)
                return
            used = self._existing_section_ids()
            if sid in used or self.preview_tree.exists(sid):
                messagebox.showwarning(t("app_title"), t("msg_sid_exists", sid=sid), parent=dlg)
                return
            pos_label = pos_var.get()
            mode, ref = pos_map.get(pos_label, ("end", None))
            result.update(sid=sid, title=title, layer=layer, pos_mode=mode, pos_ref=ref, pos_label=pos_label)
            dlg.destroy()

        def _cancel() -> None:
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text=t("dlg_cancel"), command=_cancel, style="Action.TButton").pack(side="right", padx=(6, 0))
        ttk.Button(btns, text=t("dlg_add"), command=_ok, style="Primary.TButton").pack(side="right")
        title_entry.focus_set()
        title_entry.select_range(0, tk.END)
        dlg.bind("<Return>", lambda _e: _ok())
        dlg.bind("<Escape>", lambda _e: _cancel())
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.root.wait_window(dlg)

        if not result:
            return
        sid, title, layer = result["sid"], result["title"], result["layer"]
        pos_mode = result.get("pos_mode") or "end"
        pos_ref = result.get("pos_ref")

        # 唯一 iid：优先用编号本身
        iid = sid
        n = 1
        while self.preview_tree.exists(iid):
            iid = f"{sid}#new{n}"
            n += 1

        # 先插到末尾，再 move 到目标位置（比直接用 index 更稳）
        self.preview_tree.insert(
            "", tk.END, iid=iid,
            values=("", sid, title, layer_internal_to_display(layer), 0, 0, "—", t("status_new")),
            tags=("manual",),
        )
        try:
            if pos_mode == "start":
                self.preview_tree.move(iid, "", 0)
            elif pos_mode == "after" and pos_ref and self.preview_tree.exists(pos_ref):
                # 插到参考行之后：move 到 ref 的当前下标 + 1
                # 注意：iid 已在末尾，ref 的 index 不受影响
                ref_idx = int(self.preview_tree.index(pos_ref))
                self.preview_tree.move(iid, "", ref_idx + 1)
            # end：保持在末尾
        except tk.TclError:
            pass

        self._renumber_order()
        self.preview_tree.selection_set(iid)
        self.preview_tree.focus(iid)
        self.preview_tree.see(iid)
        self.preview_ready = True
        self.generate_btn.configure(state="normal")
        order_now = int(self.preview_tree.index(iid)) + 1
        self.status_var.set(t("status_added_heading", sid=sid, title=title, n=order_now))
        self._append_log(
            f"add heading: {sid}  {title}  layer={layer}  "
            f"pos={result.get('pos_label', '')} → #{order_now}"
        )

    def _renumber_order(self) -> None:
        for i, item in enumerate(self.preview_tree.get_children(), 1):
            vals = list(self.preview_tree.item(item, "values"))
            if vals:
                vals[0] = i
                self.preview_tree.item(item, values=vals)

    def _is_hidden_item(self, item: str) -> bool:
        return "hidden" in (self.preview_tree.item(item, "tags") or ())

    def _item_display_tags(self, item: str, *, hidden: bool | None = None, edited: bool | None = None) -> tuple:
        """组合标签：hidden 优先（灰色），其次 edited / 原始状态。"""
        orig = self._preview_original.get(item, {})
        base = list(orig.get("tags") or ("normal",))
        # 去掉会冲突的状态标签，后面按需再加
        base = [x for x in base if x not in ("hidden", "edited", "normal")]
        cur = list(self.preview_tree.item(item, "tags") or ())
        is_hidden = self._is_hidden_item(item) if hidden is None else bool(hidden)
        if edited is None:
            is_edited = "edited" in cur
        else:
            is_edited = bool(edited)
        tags: list[str] = []
        if is_hidden:
            tags.append("hidden")
        if is_edited:
            tags.append("edited")
        if not tags:
            tags = base or ["normal"]
        else:
            for b in base:
                if b not in tags:
                    tags.append(b)
        return tuple(tags)

    def _resolve_item_sid(self, item: str) -> str:
        """解析树行对应的原始/当前章节编号（用于排除列表）。"""
        orig = self._preview_original.get(item, {})
        if orig.get("section_id"):
            return str(orig["section_id"]).strip()
        vals = list(self.preview_tree.item(item, "values") or ())
        if len(vals) > 1 and str(vals[1]).strip():
            return str(vals[1]).strip()
        return str(item)

    def _delete_selected_sections(self) -> str:
        """从预览删除选中章节；原解析章节记入排除列表，终版不再写入。"""
        if self.busy:
            return "break"
        sel = list(self.preview_tree.selection())
        if not sel:
            messagebox.showinfo(t("app_title"), t("msg_select_row"))
            return "break"
        n = len(sel)
        if not messagebox.askyesno(
            t("app_title"),
            t("msg_confirm_delete", n=n),
        ):
            return "break"
        deleted = 0
        for item in sel:
            if not self.preview_tree.exists(item):
                continue
            tags = self.preview_tree.item(item, "tags") or ()
            is_manual = ("manual" in tags) and (item not in self._preview_original)
            sid = self._resolve_item_sid(item)
            # 原解析章节（含前言）记入排除，避免终版回流；纯手动新增仅从预览移除
            if item in self._preview_original or sid == "_preamble":
                if sid:
                    self._deleted_sids.add(sid)
                self._deleted_sids.add(item)
                orig = self._preview_original.get(item) or {}
                if orig.get("section_id"):
                    self._deleted_sids.add(str(orig["section_id"]).strip())
            elif not is_manual and sid and not str(sid).startswith("_"):
                self._deleted_sids.add(sid)
                self._deleted_sids.add(item)
            try:
                self.preview_tree.delete(item)
                deleted += 1
            except tk.TclError:
                pass
        self._close_edit_widget(save=False)
        self._renumber_order()
        remaining = len(self.preview_tree.get_children())
        self.preview_ready = remaining > 0
        if not remaining:
            self.generate_btn.configure(state="disabled")
        self.status_var.set(t("status_deleted_n", n=deleted, left=remaining))
        self._append_log(f"delete sections: {deleted} removed, excluded={sorted(self._deleted_sids)}")
        return "break"

    def _toggle_hide_selected(self) -> None:
        """隐藏/取消隐藏选中章节（终版跳过，预览中保留灰色行）。"""
        if self.busy:
            return
        sel = list(self.preview_tree.selection())
        if not sel:
            messagebox.showinfo(t("app_title"), t("msg_select_row"))
            return
        hide_n = show_n = 0
        for item in sel:
            if not self.preview_tree.exists(item):
                continue
            vals = list(self.preview_tree.item(item, "values") or ())
            while len(vals) < 8:
                vals.append("")
            currently_hidden = self._is_hidden_item(item)
            if currently_hidden:
                # 恢复显示
                orig = self._preview_original.get(item, {})
                prev_status = str(orig.get("status") or t("status_normal"))
                tags = list(self.preview_tree.item(item, "tags") or ())
                if "edited" in tags or "manual" in tags:
                    # 保持编辑/手动标记；状态列若曾是「已隐藏」则回正常/新增
                    if str(vals[7]) in (t("status_hidden"), "已隐藏", "Hidden"):
                        if "manual" in tags:
                            vals[7] = t("status_new")
                        else:
                            vals[7] = t("status_normal") if not ("edited" in tags) else t("status_normal")
                else:
                    vals[7] = prev_status
                self.preview_tree.item(item, values=vals, tags=self._item_display_tags(item, hidden=False))
                show_n += 1
            else:
                vals[7] = t("status_hidden")
                self.preview_tree.item(item, values=vals, tags=self._item_display_tags(item, hidden=True))
                hide_n += 1
        if hide_n and not show_n:
            self.status_var.set(t("status_hidden_n", n=hide_n))
            self._append_log(f"hide sections: {hide_n}")
        elif show_n and not hide_n:
            self.status_var.set(t("status_shown_n", n=show_n))
            self._append_log(f"show sections: {show_n}")
        else:
            self.status_var.set(t("status_hide_toggle", hide=hide_n, show=show_n))
            self._append_log(f"toggle hide: hide={hide_n} show={show_n}")

    def _batch_structure_dialog(self) -> None:
        """批量改层级 / 按层级规则重编号。"""
        if self.busy:
            return
        children = list(self.preview_tree.get_children())
        if not children:
            messagebox.showinfo(t("app_title"), t("msg_preview_first"))
            return
        sel = list(self.preview_tree.selection())
        dlg = tk.Toplevel(self.root)
        dlg.title(t("dlg_batch_structure"))
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=t("dlg_batch_hint"), style="Body.TLabel", wraplength=380).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        # 作用范围
        ttk.Label(frm, text=t("dlg_batch_scope"), style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        scope_choices = [t("dlg_batch_scope_selected"), t("dlg_batch_scope_all")]
        default_scope = scope_choices[0] if sel else scope_choices[1]
        scope_var = tk.StringVar(value=default_scope)
        ttk.Combobox(
            frm, textvariable=scope_var, values=scope_choices, width=28, state="readonly",
        ).grid(row=1, column=1, sticky="ew", pady=4, padx=(8, 0))

        # 层级：保持 / 设为某级 / 升一级 / 降一级
        ttk.Label(frm, text=t("dlg_batch_layer"), style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        layer_choices = [t("dlg_batch_layer_keep")] + list(layer_display_names()) + [
            t("dlg_batch_shift_up"), t("dlg_batch_shift_down"),
        ]
        layer_var = tk.StringVar(value=layer_choices[0])
        ttk.Combobox(
            frm, textvariable=layer_var, values=layer_choices, width=28, state="readonly",
        ).grid(row=2, column=1, sticky="ew", pady=4, padx=(8, 0))

        # 编号：不改 / 按当前顺序+层级重编号
        ttk.Label(frm, text=t("dlg_batch_renumber"), style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=4)
        renumber_choices = [t("dlg_batch_renumber_off"), t("dlg_batch_renumber_on")]
        renumber_var = tk.StringVar(value=renumber_choices[0])
        ttk.Combobox(
            frm, textvariable=renumber_var, values=renumber_choices, width=28, state="readonly",
        ).grid(row=3, column=1, sticky="ew", pady=4, padx=(8, 0))

        frm.columnconfigure(1, weight=1)
        result: dict = {}

        def _ok() -> None:
            result.update(
                scope=scope_var.get(),
                layer=layer_var.get(),
                renumber=renumber_var.get(),
            )
            dlg.destroy()

        def _cancel() -> None:
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text=t("dlg_cancel"), command=_cancel, style="Action.TButton").pack(
            side="right", padx=(6, 0)
        )
        ttk.Button(btns, text=t("dlg_batch_apply"), command=_ok, style="Primary.TButton").pack(side="right")
        dlg.bind("<Return>", lambda _e: _ok())
        dlg.bind("<Escape>", lambda _e: _cancel())
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.root.wait_window(dlg)
        if not result:
            return
        self._apply_batch_structure(result, selected=sel, all_items=children)

    def _apply_batch_structure(
        self, opts: dict, *, selected: list[str], all_items: list[str],
    ) -> None:
        scope_sel = opts.get("scope") == t("dlg_batch_scope_selected")
        targets = [i for i in selected if self.preview_tree.exists(i)] if scope_sel else list(
            self.preview_tree.get_children()
        )
        if not targets:
            messagebox.showinfo(t("app_title"), t("msg_select_row"))
            return

        layer_opt = str(opts.get("layer") or "")
        renumber_on = opts.get("renumber") == t("dlg_batch_renumber_on")
        changed = 0

        # 1) 改层级
        if layer_opt and layer_opt != t("dlg_batch_layer_keep"):
            for item in targets:
                if item == "_preamble" or self._resolve_item_sid(item) == "_preamble":
                    continue  # 前言层级固定
                vals = list(self.preview_tree.item(item, "values") or ())
                while len(vals) < 8:
                    vals.append("")
                cur_layer = layer_display_to_internal(str(vals[3]) if len(vals) > 3 else "节")
                cur_lv = layer_to_level(cur_layer, self._resolve_item_sid(item))
                if layer_opt == t("dlg_batch_shift_up"):
                    new_lv = max(1, cur_lv - 1)
                elif layer_opt == t("dlg_batch_shift_down"):
                    new_lv = min(MAX_HEADING_LEVEL, cur_lv + 1)
                else:
                    new_lv = layer_to_level(layer_display_to_internal(layer_opt), "")
                    if not (1 <= new_lv <= MAX_HEADING_LEVEL):
                        continue
                new_layer = level_to_layer(new_lv)
                vals[3] = layer_internal_to_display(new_layer)
                self.preview_tree.item(
                    item, values=vals,
                    tags=self._item_display_tags(item, edited=True, hidden=self._is_hidden_item(item)),
                )
                changed += 1

        # 2) 按顺序+层级重编号（栈算法：1 → 1.1 → 1.1.1）
        if renumber_on:
            stack: list[int] = []
            used: set[str] = set()
            renumber_set = set(targets)
            # 全表模式：隐藏行不参与重编号，但预占其 sid 避免冲突
            # 选中模式：未选中行预占 sid
            for item in self.preview_tree.get_children():
                if item in renumber_set:
                    if not scope_sel and self._is_hidden_item(item):
                        pass  # 全表时隐藏行不重编号，下面预占
                    else:
                        continue
                vals = list(self.preview_tree.item(item, "values") or ())
                if len(vals) > 1 and str(vals[1]).strip():
                    used.add(str(vals[1]).strip())

            walk = list(self.preview_tree.get_children())
            for item in walk:
                if not self.preview_tree.exists(item):
                    continue
                if item not in renumber_set:
                    continue
                sid0 = self._resolve_item_sid(item)
                if sid0 == "_preamble" or item == "_preamble":
                    continue
                if self._is_hidden_item(item) and not scope_sel:
                    # 全表重编号时隐藏行保持原号（已预占）
                    continue
                vals = list(self.preview_tree.item(item, "values") or ())
                while len(vals) < 8:
                    vals.append("")
                layer = layer_display_to_internal(str(vals[3]) if len(vals) > 3 else "节")
                lv = layer_to_level(layer, sid0)
                lv = max(1, min(MAX_HEADING_LEVEL, lv))
                # 层级栈：跳级时中间补 1（如 2 后接 level=3 → 2.1.1）
                if len(stack) < lv:
                    while len(stack) < lv:
                        stack.append(0)
                else:
                    stack = stack[:lv]
                stack[-1] += 1
                for i in range(len(stack) - 1):
                    if stack[i] <= 0:
                        stack[i] = 1
                new_sid = ".".join(str(x) for x in stack)
                guard = 0
                while new_sid in used and guard < 500:
                    stack[-1] += 1
                    new_sid = ".".join(str(x) for x in stack)
                    guard += 1
                used.add(new_sid)
                old_sid = str(vals[1]).strip() if len(vals) > 1 else ""
                if new_sid != old_sid:
                    vals[1] = new_sid
                    changed += 1
                self.preview_tree.item(
                    item, values=vals,
                    tags=self._item_display_tags(item, edited=True, hidden=self._is_hidden_item(item)),
                )

        self._renumber_order()
        self.status_var.set(t("msg_batch_done", n=changed if changed else len(targets)))
        self._append_log(
            f"batch structure: scope={'sel' if scope_sel else 'all'} "
            f"layer={opts.get('layer')!r} renumber={renumber_on} changed≈{changed}"
        )

    def _collect_section_order(self) -> list[str]:
        """终版顺序：仅可见（未隐藏）行；用树 iid 以便与 overrides/rename 对齐。"""
        order: list[str] = []
        for item in self.preview_tree.get_children():
            if self._is_hidden_item(item):
                continue
            order.append(item)
        return order

    def _collect_exclude_section_ids(self) -> list[str]:
        """生成时排除：已删除 + 已隐藏（含 iid 与当前/原始编号）。"""
        exclude: set[str] = set(self._deleted_sids)
        for item in self.preview_tree.get_children():
            if not self._is_hidden_item(item):
                continue
            exclude.add(item)
            exclude.add(self._resolve_item_sid(item))
            vals = list(self.preview_tree.item(item, "values") or ())
            if len(vals) > 1 and str(vals[1]).strip():
                exclude.add(str(vals[1]).strip())
            orig = self._preview_original.get(item, {})
            if orig.get("section_id"):
                exclude.add(str(orig["section_id"]).strip())
        return sorted(s for s in exclude if s)

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
            choices = _layer_choices()
            widget: tk.Widget = ttk.Combobox(self.preview_tree, values=choices, state="readonly", font=(UI_FONT, 9))
            cur_disp = layer_internal_to_display(layer_display_to_internal(current))
            widget.set(cur_disp if cur_disp in choices else t("layer_2"))
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
            text_internal = layer_display_to_internal(text)
            if text_internal not in LAYER_NAMES and text_internal != "前言":
                text_internal = layer_display_to_internal(str(orig.get("layer") or "节"))
            if item != "_preamble" and text_internal == "前言":
                text_internal = "章"
            vals[3] = layer_internal_to_display(text_internal)
        changed = (
            str(vals[1]) != str(orig.get("section_id", vals[1]))
            or str(vals[2]) != str(orig.get("title", vals[2]))
            or str(vals[3]) != str(orig.get("layer", vals[3]))
        )
        self.preview_tree.item(
            item, values=vals,
            tags=self._item_display_tags(item, edited=changed or "edited" in (self.preview_tree.item(item, "tags") or ())),
        )

    def _collect_section_overrides(self) -> dict[str, dict]:
        overrides: dict[str, dict] = {}
        for item in self.preview_tree.get_children():
            # 隐藏行终版排除，无需覆盖
            if self._is_hidden_item(item):
                continue
            vals = self.preview_tree.item(item, "values")
            new_sid = str(vals[1]).strip() if len(vals) > 1 else ""
            new_title = str(vals[2]).strip() if len(vals) > 2 else ""
            new_layer = layer_display_to_internal(str(vals[3]).strip() if len(vals) > 3 else "节")
            new_level = layer_to_level(new_layer, new_sid or item)
            orig = self._preview_original.get(item)
            tags = self.preview_tree.item(item, "tags") or ()
            is_manual = (not orig) or ("manual" in tags) or (
                len(vals) > 7 and str(vals[7]) in (t("status_new"), "新增", "New", "Neu")
            )
            if is_manual:
                # 手动新增的空标题：合成时创建无正文 Section
                if not new_sid:
                    continue
                overrides[item] = {
                    "create": True,
                    "section_id": new_sid,
                    "title": new_title or "未命名",
                    "level": new_level,
                }
                continue
            if not orig:
                continue
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
            messagebox.showwarning(t("app_title"), t("msg_add_docx"))
            return
        if self.busy:
            return
        # 若启用模板但尚未加载，尝试加载
        if self.use_template_var.get() and self.format_template is None and self.template_path_var.get().strip():
            self.load_template()
        self._set_busy(True, t("busy_preview"))
        self._append_log(t("log_start_parse"))
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
        path = filedialog.askdirectory(title=t("dlg_pick_outdir"), initialdir=self.vars["output_dir"].get())
        if path:
            self.vars["output_dir"].set(path)

    def _font_config(self) -> FontConfig:
        """完全按界面「格式设置」组装 FontConfig（无模板时使用）。"""
        def yn(level: str, key: str) -> bool:
            return is_yes(self.font_vars[f"{level}_{key}"].get())

        def color_of(level: str) -> str:
            return _color_options().get(self.font_vars[f"{level}_color"].get(), "000000")

        try:
            indent = int(float(self.indent_var.get() or "2"))
        except ValueError:
            indent = 2
        try:
            spacing = float(self.line_spacing_var.get() or "1.5")
        except ValueError:
            spacing = 1.5
        return FontConfig(
            h1_font=self.font_vars["h1_font"].get(), h1_size=float(self.font_vars["h1_size"].get()),
            h1_bold=yn("h1", "bold"), h1_italic=yn("h1", "italic"), h1_underline=yn("h1", "underline"),
            h1_color=color_of("h1"),
            h2_font=self.font_vars["h2_font"].get(), h2_size=float(self.font_vars["h2_size"].get()),
            h2_bold=yn("h2", "bold"), h2_italic=yn("h2", "italic"), h2_underline=yn("h2", "underline"),
            h2_color=color_of("h2"),
            h3_font=self.font_vars["h3_font"].get(), h3_size=float(self.font_vars["h3_size"].get()),
            h3_bold=yn("h3", "bold"), h3_italic=yn("h3", "italic"), h3_underline=yn("h3", "underline"),
            h3_color=color_of("h3"),
            h4_font=self.font_vars["h4_font"].get(), h4_size=float(self.font_vars["h4_size"].get()),
            h4_bold=yn("h4", "bold"), h4_italic=yn("h4", "italic"), h4_underline=yn("h4", "underline"),
            h4_color=color_of("h4"),
            h5_font=self.font_vars["h5_font"].get(), h5_size=float(self.font_vars["h5_size"].get()),
            h5_bold=yn("h5", "bold"), h5_italic=yn("h5", "italic"), h5_underline=yn("h5", "underline"),
            h5_color=color_of("h5"),
            h6_font=self.font_vars["h6_font"].get(), h6_size=float(self.font_vars["h6_size"].get()),
            h6_bold=yn("h6", "bold"), h6_italic=yn("h6", "italic"), h6_underline=yn("h6", "underline"),
            h6_color=color_of("h6"),
            body_font=self.font_vars["body_font"].get(), body_size=float(self.font_vars["body_size"].get()),
            body_bold=yn("body", "bold"), body_italic=yn("body", "italic"), body_underline=yn("body", "underline"),
            body_color=color_of("body"), body_indent=indent, body_line_spacing=spacing,
        )

    def _active_fmt_attrs(self) -> list[str]:
        """当前勾选的属性覆盖字段（font/size/…）。若一个都没勾，视为不覆盖字符属性。"""
        return [a for a, v in self.fmt_attr_override_vars.items() if v.get()]

    def _resolved_font_config(self) -> tuple[FontConfig, str]:
        """生成用版式：模板为底 + 勾选行 × 勾选属性 用界面覆盖；无模板则全手动。

        返回 (font_config, 日志说明)。
        """
        ft = self.format_template if self.use_template_var.get() else None
        if not (ft and ft.font_config):
            return self._font_config(), t("log_layout_manual")

        base = replace(ft.font_config)
        manual = self._font_config()
        overridden: list[str] = []
        level_attrs = ("font", "size", "bold", "italic", "underline", "color")
        active_attrs = self._active_fmt_attrs()

        for level in ("h1", "h2", "h3", "h4", "h5", "h6", "body"):
            if not self.fmt_override_vars.get(level) or not self.fmt_override_vars[level].get():
                continue
            if not active_attrs:
                # 行勾了但属性全关：该级不覆盖任何字符字段
                continue
            applied: list[str] = []
            for attr in level_attrs:
                if attr not in active_attrs:
                    continue  # 未勾属性 → 保留模板
                key = f"{level}_{attr}"
                setattr(base, key, getattr(manual, key))
                applied.append(attr)
            if applied:
                tag = level.upper() if level != "body" else "body"
                overridden.append(f"{tag}({'+'.join(applied)})")

        if self.fmt_override_vars.get("indent") and self.fmt_override_vars["indent"].get():
            base.body_indent = manual.body_indent
            overridden.append("indent")
        if self.fmt_override_vars.get("spacing") and self.fmt_override_vars["spacing"].get():
            base.body_line_spacing = manual.body_line_spacing
            overridden.append("spacing")

        if overridden:
            note = (
                f"版式来源：格式模板 + 界面覆盖 [{', '.join(overridden)}] "
                f"（H1=#{base.h1_color} H2=#{base.h2_color} 正文=#{base.body_color} "
                f"行距={base.body_line_spacing} 缩进={base.body_indent}）"
            )
        else:
            note = (
                f"版式来源：格式模板（无覆盖；H1=#{base.h1_color} H2=#{base.h2_color} "
                f"正文=#{base.body_color} 行距={base.body_line_spacing} 缩进={base.body_indent}）"
            )
        return base, note

    def start_generation(self) -> None:
        if not self.preview_ready or not self.pending_temp_dir:
            messagebox.showwarning(t("app_title"), t("msg_preview_before_gen"))
            return
        if not self.files:
            messagebox.showwarning(t("app_title"), t("msg_need_docx"))
            return
        output_dir = self.vars["output_dir"].get().strip()
        if not output_dir:
            messagebox.showwarning(t("app_title"), t("msg_need_outdir"))
            return
        try:
            os.makedirs(output_dir, exist_ok=True)
            if not os.access(output_dir, os.W_OK):
                raise OSError("目录不可写")
        except OSError as exc:
            messagebox.showerror(t("app_title"), t("msg_outdir_fail", err=exc))
            return

        if self.use_template_var.get():
            if self.format_template is None:
                if self.template_path_var.get().strip():
                    self.load_template()
                if self.format_template is None:
                    messagebox.showwarning(t("app_title"), t("msg_tpl_not_loaded"))
                    return

        paths = build_output_paths(output_dir, self.vars["period"].get())
        demo = os.environ.get("REPORT_DEMO", "").strip() in ("1", "true", "TRUE", "yes")
        existing = [p for p in (paths["docx"], paths["pdf"]) if os.path.exists(p)]
        if existing and not demo:
            if not messagebox.askyesno(
                t("app_title"), t("msg_overwrite", files="\n".join(existing))
            ):
                return

        allow_skip = False
        if self.preview_failures:
            details = "\n".join(f"{getattr(f, 'file_name', '?')} · {getattr(f, 'message', '')}" for f in self.preview_failures)
            if demo:
                allow_skip = True
                self._append_log(f"demo skip failures\n{details}")
            elif not messagebox.askyesno(
                t("app_title"), t("msg_skip_fail", details=details)
            ):
                return
            else:
                allow_skip = True

        # 版式：启用模板 → 默认跟模板；勾选「覆盖」的行用界面设置
        # 未启用模板 → 全部用界面格式设置
        ft = self.format_template if self.use_template_var.get() else None
        font_config, layout_note = self._resolved_font_config()
        self._append_log(layout_note)

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
            "exclude_section_ids": self._collect_exclude_section_ids(),
            "strict_section_order": True,
            "format_template": ft,
            "use_template_titles": self.use_tpl_titles_var.get() if ft else False,
            "staged_root": self.pending_temp_dir,
        }
        self._set_busy(True, t("busy_generate"))
        self._append_log(t("log_start_gen"))
        if ft:
            self._append_log(f"使用格式模板：{ft.source_path}（{len(ft.specs)} 章节）")
        self._append_log(f"章节顺序：{len(options['section_order'])} 项（严格预览顺序）")
        if options["exclude_section_ids"]:
            self._append_log(
                f"排除章节（删除/隐藏）：{len(options['exclude_section_ids'])} 项 → "
                f"{', '.join(options['exclude_section_ids'][:12])}"
                f"{'…' if len(options['exclude_section_ids']) > 12 else ''}"
            )
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
                strict_section_order=bool(options.get("strict_section_order")),
                exclude_section_ids=options.get("exclude_section_ids") or None,
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
                    self._deleted_sids.clear()
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
                        layer = str(row.get("layer", ""))  # 内部中文
                        level = int(row.get("level") or layer_to_level(layer, sid))
                        layer_disp = layer_internal_to_display(layer)
                        iid = sid if not self.preview_tree.exists(sid) else f"{sid}#{row.get('order')}"
                        self._preview_original[iid] = {
                            "section_id": sid, "title": title, "layer": layer, "level": level,
                            "order": row.get("order", 0), "block_count": row.get("block_count", 0),
                            "source_count": row.get("source_count", 0), "confidence": row.get("confidence", ""),
                            "status": status, "tags": (tag,),
                        }
                        self.preview_tree.insert(
                            "", tk.END, iid=iid,
                            values=(row.get("order", ""), sid, title, layer_disp, row.get("block_count", 0),
                                    row.get("source_count", 0), row.get("confidence", ""), status),
                            tags=(tag,),
                        )
                    nsec = len(event.get("rows") or [])
                    self.preview_ready = True
                    self.generate_btn.configure(state="normal")
                    self.status_var.set(t("status_preview_done", n=nsec))
                    self._append_log(f"preview: {event.get('count', 0)} files → {nsec} sections")
                elif kind == "done":
                    self._set_busy(False)
                    self.last_docx = event.get("docx") or None
                    self.last_pdf = event.get("pdf") or None
                    self.open_docx_btn.configure(state="normal" if self.last_docx else "disabled")
                    self.open_pdf_btn.configure(state="normal" if self.last_pdf else "disabled")
                    docx, pdf, pdf_msg = event.get("docx", ""), event.get("pdf", ""), event.get("pdf_msg", "")
                    msg = f"DOCX 已保存：\n{docx}\n\n" + (f"PDF 已保存：\n{pdf}" if pdf else f"PDF 未生成：{pdf_msg}")
                    self.status_var.set(t("status_gen_done"))
                    self._append_log(msg.replace("\n", " | "))
                    # 演示录屏时不弹阻塞对话框，避免打断时间线
                    if os.environ.get("REPORT_DEMO", "").strip() not in ("1", "true", "TRUE", "yes"):
                        messagebox.showinfo(t("app_title"), msg)
                elif kind == "error":
                    self._set_busy(False)
                    self._append_log(event.get("detail") or event.get("message") or "")
                    self.status_var.set(t("status_gen_fail"))
                    if os.environ.get("REPORT_DEMO", "").strip() not in ("1", "true", "TRUE", "yes"):
                        messagebox.showerror(
                            t("app_title"), t("msg_gen_fail", err=event.get("message", ""))
                        )
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
            messagebox.showwarning(t("app_title"), "文件不存在。")
            return
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror(t("app_title"), f"无法打开：{exc}")

    def open_output_dir(self) -> None:
        path = self.vars["output_dir"].get()
        if not os.path.isdir(path):
            messagebox.showwarning(t("app_title"), "输出目录不存在。")
            return
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror(t("app_title"), f"无法打开目录：{exc}")

    def on_close(self) -> None:
        if self.busy:
            messagebox.showwarning(t("app_title"), "任务进行中，请稍候再关闭。")
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
