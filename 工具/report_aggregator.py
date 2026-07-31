"""
report_aggregator.py
====================

将各部门提供的"分板块 docx"按章节编号动态汇总为终版季度调研报告。

核心流程：
  1. 解析每个输入文件，按"编号 + 标题"识别 section 边界
  2. 按 section_id 数字排序，自动推断层级（章/节/小节）
  3. 套用统一格式（含图片、表格原位插入），拼接封面 + 目录，输出终版 docx

默认动态模式——章节结构和标题完全由输入文件决定，按 section_id 数字分量排序。
可选传入 template 参数作为排序参考：模板内章节按模板顺序排前，模板外按数字排后，
模板期望但输入缺失的章节记日志告警（不强制必填、不改变结构，见 F12）。

增强功能：
  - 保留输入文件中的加粗文字格式
  - 自动生成目录（TOC）
  - 图表题注居中 + SEQ 域自动编号 + 书签交叉引用
  - 表格首行突出 + 隔行换色（白 + 浅蓝）

支持三种使用方式：
  - 编程调用 aggregate(input_dir, output_path, cover_kwargs)
  - 命令行 python report_aggregator.py <input_dir> <output_path>
  - 通过 report_app.py 的 Streamlit 界面调用
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import contextvars
from copy import deepcopy
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.document import Document as _Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree


# ---------------------------------------------------------------------------
# F10：损坏/加密/伪装 DOCX 错误隔离
#   解析结果设计为"成功文件、警告、失败文件"三部分，单文件异常不再使整
#   个流程崩溃；调用方据失败文件决定是否禁止生成。
# ---------------------------------------------------------------------------

# 上传/解析阶段检查的文件大小上限（200MB）：超过视为可疑文件，避免耗尽内存
MAX_INPUT_FILE_BYTES = 200 * 1024 * 1024
# OOXML 包必需的核心部件（缺一即判定为非合法 docx / 伪装文件）
_REQUIRED_OOXML_PARTS = ("word/document.xml", "[Content_Types].xml")


@dataclass
class FileFailure:
    """解析失败的文件记录。"""
    file_name: str
    category: str          # bad_zip / missing_parts / encrypted / xml_error / io_error / oversized / other
    message: str
    size: Optional[int] = None  # 文件大小（字节），未知时为 None


@dataclass
class FileWarning:
    """解析成功但带告警的文件记录（例如低置信度标题批量出现）。"""
    file_name: str
    message: str


@dataclass
class ParseResult:
    """parse_input_dir 的结构化返回：成功文件、警告、失败文件三部分。

    与旧调用方兼容：
      - 持有 .sections: Dict[str, Section]（旧 parse_input_dir 的返回）
      - 支持 bool() / len() / 迭代：等价于操作 .sections
      - 失败文件非空时 has_failures() 为真，调用方据此禁止生成
    """
    sections: Dict[str, "Section"] = field(default_factory=dict)
    warnings: List[FileWarning] = field(default_factory=list)
    failures: List[FileFailure] = field(default_factory=list)
    parsed_files: List[str] = field(default_factory=list)  # 成功解析的文件名

    # ---- 旧 API 兼容：让 ParseResult 可直接当 dict 用 ----
    def __bool__(self):
        return bool(self.sections)

    def __len__(self):
        return len(self.sections)

    def __iter__(self):
        return iter(self.sections)

    def __getitem__(self, key):
        return self.sections[key]

    def __contains__(self, key):
        return key in self.sections

    def keys(self):
        return self.sections.keys()

    def values(self):
        return self.sections.values()

    def items(self):
        return self.sections.items()

    def get(self, key, default=None):
        return self.sections.get(key, default)

    def pop(self, key, default=None):
        # dict 兼容：与 keys/values/items/get 一致地代理到 self.sections。
        # F05 在 aggregate 中用 parsed.pop("_preamble") 单独取出前导内容，
        # 取出后 preamble 不参与正常章节的数字排序（避免 int("_preamble") 崩溃）。
        return self.sections.pop(key, default)

    def has_failures(self) -> bool:
        return len(self.failures) > 0


def _classify_parse_error(exc: BaseException, path: str) -> FileFailure:
    """把解析异常归入 F10 的失败类别，返回 FileFailure。"""
    fname = os.path.basename(path)
    size = None
    try:
        size = os.path.getsize(path)
    except OSError:
        pass

    # 加密 docx：python-docx 抛 KeyError('1Table') 或类似 "Encrypted document"
    # OpenXML 加密包特征：根目录有 EncryptedPackage 部件而非 word/document.xml
    msg = str(exc)
    exc_name = type(exc).__name__

    # 显式加密标志
    if exc_name == "KeyError" and ("EncryptedPackage" in msg or "1Table" in msg
                                   or "0Table" in msg):
        return FileFailure(fname, "encrypted", f"文档已加密，需先解密：{msg}", size)

    # BadZipFile：不是 zip（伪装/损坏/普通文本改后缀）
    if isinstance(exc, zipfile.BadZipFile):
        return FileFailure(fname, "bad_zip", f"不是有效的 ZIP/OOXML 包：{msg}", size)

    # 包部件缺失：python-docx 在 Package.open 时抛 KeyError 缺核心部件
    if exc_name == "KeyError":
        return FileFailure(fname, "missing_parts", f"OOXML 包缺少必需部件：{msg}", size)

    # XML 解析错误：OOXML 部件 XML 损坏
    if isinstance(exc, etree.XMLSyntaxError):
        return FileFailure(fname, "xml_error", f"XML 解析失败：{msg}", size)
    # python-docx 内部用 etree，XMLSyntaxError 可能被包装进 ValueError/LookupError
    if exc_name in ("ValueError", "LookupError") and (
        "ParseError" in msg or "mismatched tag" in msg or "syntax error" in msg.lower()
    ):
        return FileFailure(fname, "xml_error", f"XML 解析失败：{msg}", size)

    # I/O 错误
    if isinstance(exc, OSError):
        return FileFailure(fname, "io_error", f"读取文件失败：{msg}", size)

    # 其他
    return FileFailure(fname, "other", f"{exc_name}: {msg}", size)


def _validate_docx_package(path: str) -> Optional[FileFailure]:
    """F10：上传/解析前预检查——ZIP 魔数、必需 OOXML 部件、文件大小。

    返回 None 表示通过；返回 FileFailure 表示问题文件应跳过，不再尝试 Document(path)。
    此函数只做静态检查（不解压全部内容），开销小。
    """
    fname = os.path.basename(path)
    try:
        size = os.path.getsize(path)
    except OSError as e:
        return FileFailure(fname, "io_error", f"无法获取文件大小：{e}")

    # 大小检查
    if size == 0:
        return FileFailure(fname, "io_error", "文件为空（0 字节）", size)
    if size > MAX_INPUT_FILE_BYTES:
        return FileFailure(
            fname, "oversized",
            f"文件过大：{size} 字节 > {MAX_INPUT_FILE_BYTES} 字节上限", size)

    # ZIP 魔数检查（前 4 字节：PK\x03\x04 或空归档 PK\x05\x06）
    try:
        with open(path, "rb") as f:
            magic = f.read(4)
    except OSError as e:
        return FileFailure(fname, "io_error", f"读取文件头失败：{e}", size)

    if magic[:2] != b"PK":
        # 不是 zip：可能是伪装的 docx（实际是 pdf/doc/txt 改后缀）
        # 常见 PDF 魔数 %PDF、旧版 doc 的 D0CF11E0（OLE 复合文档）
        hint = ""
        if magic[:4] == b"%PDF":
            hint = "（检测到 PDF 签名，疑似 PDF 改后缀伪装成 docx）"
        elif magic[:4] == b"\xd0\xcf\x11\xe0":
            hint = "（检测到 OLE 复合文档签名，疑似旧版 .doc 改后缀伪装成 docx）"
        elif magic[:4] == b"PK\x05\x06":
            hint = "（空 ZIP 归档，无任何内容）"
        return FileFailure(
            fname, "bad_zip",
            f"不是有效的 ZIP/OOXML 包（缺少 PK 签名）{hint}", size)

    # 必需 OOXML 部件检查（用 zipfile 不解压全部，只测存在性）
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            # 加密文件特征：根目录有 EncryptedPackage
            if "EncryptedPackage" in names and "word/document.xml" not in names:
                return FileFailure(
                    fname, "encrypted",
                    "文档已加密（检测到 EncryptedPackage 部件，需先解密）", size)
            missing = [p for p in _REQUIRED_OOXML_PARTS if p not in names]
            if missing:
                return FileFailure(
                    fname, "missing_parts",
                    f"OOXML 包缺少必需部件：{', '.join(missing)}", size)
            # 测试核心部件可解压（捕捉"zip 头完好但内部损坏"的情况）
            try:
                # 测试 document.xml 可读
                with zf.open("word/document.xml") as _:
                    pass
            except (zipfile.BadZipFile, OSError, RuntimeError) as e:
                return FileFailure(
                    fname, "bad_zip",
                    f"核心部件 word/document.xml 读取失败：{e}", size)
    except zipfile.BadZipFile as e:
        return FileFailure(fname, "bad_zip", f"ZIP 结构损坏：{e}", size)
    except OSError as e:
        return FileFailure(fname, "io_error", f"打开 ZIP 失败：{e}", size)

    return None  # 静态检查通过


# ---------------------------------------------------------------------------
# 章节识别（无固定模板，结构由输入文件决定）
# ---------------------------------------------------------------------------

@dataclass
class HeadingSpec:
    """标题节点（仅用于可选的 template 参数，默认不使用）。"""
    section_id: str            # 排序用的唯一 ID：'1'、'2.1'、'2.1.1'
    title: str                 # 显示标题
    level: int                 # 1=章, 2=节, 3=小节
    parent: Optional[str] = None


@dataclass
class FormatTemplate:
    """从一份「合成终版 DOCX」抽取的格式/结构参考。

    典型用法：以 Q2 终版报告为模板，指导 Q3 输入材料的合成顺序、标题与字体。
    """
    specs: List[HeadingSpec] = field(default_factory=list)
    font_config: Optional["FontConfig"] = None
    period: str = ""
    title: str = ""
    subtitle: str = ""
    org: str = ""
    date: str = ""
    source_path: str = ""
    use_template_titles: bool = True  # 匹配到的章节优先用模板标题

    def section_order(self) -> List[str]:
        return [s.section_id for s in self.specs]

    def title_map(self) -> Dict[str, str]:
        return {s.section_id: s.title for s in self.specs if s.title}

    def level_map(self) -> Dict[str, int]:
        return {s.section_id: s.level for s in self.specs if s.level}


def _pt_of_run(run) -> Optional[float]:
    try:
        if run.font.size is not None:
            return float(run.font.size.pt)
    except Exception:
        pass
    return None


def _font_name_of_run(run) -> Optional[str]:
    try:
        if run.font.name:
            return run.font.name
    except Exception:
        pass
    try:
        rPr = run._element.rPr
        if rPr is not None and rPr.rFonts is not None:
            # eastAsia / ascii
            ea = rPr.rFonts.get(qn("w:eastAsia"))
            if ea:
                return ea
            ascii_f = rPr.rFonts.get(qn("w:ascii"))
            if ascii_f:
                return ascii_f
    except Exception:
        pass
    return None


def _color_of_run(run, para=None) -> Optional[str]:
    """提取 run（或段落样式）的 RGB 颜色，返回 6 位十六进制（如 \"1F3864\"）。

    主题色/自动色无法解析时返回 None，调用方保留 FontConfig 默认值。
    """
    def _rgb_to_hex(rgb) -> Optional[str]:
        if rgb is None:
            return None
        try:
            s = str(rgb).upper().replace("#", "")
            if len(s) == 6 and all(c in "0123456789ABCDEF" for c in s):
                return s
        except Exception:
            pass
        return None

    try:
        if run is not None and run.font.color is not None:
            hx = _rgb_to_hex(getattr(run.font.color, "rgb", None))
            if hx:
                return hx
    except Exception:
        pass
    # 段落样式字色（Word 常把标题色放在 Heading 样式上，run 本身无色）
    if para is not None and para.style is not None:
        try:
            st_color = para.style.font.color
            if st_color is not None:
                hx = _rgb_to_hex(getattr(st_color, "rgb", None))
                if hx:
                    return hx
        except Exception:
            pass
    return None


def _run_format_sample(
    run, para=None
) -> Tuple[Optional[str], Optional[float], Optional[bool], Optional[bool], Optional[bool], Optional[str]]:
    """从 run（缺省时回退段落样式）采样 font / size / bold / italic / underline / color。"""
    fname = _font_name_of_run(run) if run is not None else None
    fsize = _pt_of_run(run) if run is not None else None
    fbold = fitalic = funder = None
    if run is not None:
        try:
            fbold = run.font.bold
        except Exception:
            fbold = None
        try:
            fitalic = run.font.italic
        except Exception:
            fitalic = None
        try:
            u = run.font.underline
            funder = bool(u) if u is not None else None
        except Exception:
            funder = None
    fcolor = _color_of_run(run, para=para)

    # run 未显式指定时，尝试段落样式
    if para is not None and para.style is not None:
        st = para.style
        try:
            if not fname and st.font.name:
                fname = st.font.name
        except Exception:
            pass
        try:
            if fsize is None and st.font.size is not None:
                fsize = float(st.font.size.pt)
        except Exception:
            pass
        try:
            if fbold is None and st.font.bold is not None:
                fbold = st.font.bold
        except Exception:
            pass
        try:
            if fitalic is None and st.font.italic is not None:
                fitalic = st.font.italic
        except Exception:
            pass
        try:
            if funder is None and st.font.underline is not None:
                funder = bool(st.font.underline)
        except Exception:
            pass
    return fname, fsize, fbold, fitalic, funder, fcolor


def load_format_template(docx_path: str) -> FormatTemplate:
    """从合成终版 DOCX 加载格式模板：章节顺序/标题/层级 + 字体/字号/加粗/颜色 + 封面文案。

    不会复制模板正文内容，只抽取「怎么排、怎么写标题、用什么字体颜色」。
    启用模板合成时，这些版式设定严格覆盖手动格式设置。
    """
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"模板文件不存在：{docx_path}")
    doc = Document(docx_path)
    specs: List[HeadingSpec] = []
    seen: set = set()
    cover_lines: List[str] = []
    first_heading_seen = False

    h1_font = h2_font = h3_font = body_font = None
    h1_size = h2_size = h3_size = body_size = None
    h1_bold = h2_bold = h3_bold = body_bold = None
    h1_italic = h2_italic = h3_italic = body_italic = None
    h1_underline = h2_underline = h3_underline = body_underline = None
    h1_color = h2_color = h3_color = body_color = None
    body_line_spacing = None
    h1_captured = h2_captured = h3_captured = False

    def _plausible_sid(sid: str) -> bool:
        """过滤误识别（如把 2026年… 当成章节 2026）。"""
        if not sid or sid.startswith("_"):
            return False
        # 年份误识别
        if re.fullmatch(r"20\d{2}", sid):
            return False
        parts = sid.split(".")
        try:
            nums = [int(p) for p in parts]
        except ValueError:
            return False
        # 章节号过大基本不是报告结构
        if nums[0] > 30:
            return False
        if any(n > 50 for n in nums):
            return False
        return True

    def _sample_body(para) -> None:
        nonlocal body_font, body_size, body_bold, body_italic, body_underline, body_color, body_line_spacing
        if not para.runs:
            return
        rn = para.runs[0]
        fname, fsize, fbold, fitalic, funder, fcolor = _run_format_sample(rn, para=para)
        if body_font is None and fname:
            body_font = fname
        if body_size is None and fsize is not None:
            body_size = fsize
        if body_bold is None and fbold is not None:
            body_bold = fbold
        if body_italic is None and fitalic is not None:
            body_italic = fitalic
        if body_underline is None and funder is not None:
            body_underline = funder
        if body_color is None and fcolor:
            body_color = fcolor
        if body_line_spacing is None:
            try:
                ls = para.paragraph_format.line_spacing
                if ls is not None:
                    body_line_spacing = float(ls)
            except Exception:
                pass

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        style_name = para.style.name if para.style else ""
        is_heading_style = bool(
            style_name
            and (style_name.startswith("Heading") or style_name.startswith("标题"))
        )

        # 封面：第一个 Heading 样式标题之前的非空段落
        if not first_heading_seen:
            if is_heading_style and text:
                first_heading_seen = True
            else:
                if text and len(cover_lines) < 6:
                    if text not in ("目录", "目 录") and not text.startswith("TOC"):
                        cover_lines.append(text)
                if body_font is None and para.runs and len(text) > 15:
                    _sample_body(para)
                continue

        if not text:
            continue

        # 模板结构：优先 Heading 样式 + 文本编号解析
        info = identify_heading(text, para=para)
        sid = title = None
        level = None
        if info and info.section_id and _plausible_sid(str(info.section_id)):
            sid, title, level = info.section_id, info.title, info.level
        elif is_heading_style:
            info2 = identify_heading(text, para=None)
            if info2 and info2.section_id and _plausible_sid(str(info2.section_id)):
                sid, title, level = info2.section_id, info2.title, info2.level
            else:
                continue
        else:
            if body_font is None and para.runs and len(text) > 20:
                _sample_body(para)
            elif body_color is None and para.runs and len(text) > 20:
                _sample_body(para)
            continue

        if not sid or sid in seen:
            continue
        seen.add(sid)
        parent = sid.rsplit(".", 1)[0] if "." in sid else None
        lvl = level or (sid.count(".") + 1)
        specs.append(HeadingSpec(section_id=sid, title=title or "", level=lvl, parent=parent))

        rn = para.runs[0] if para.runs else None
        fname, fsize, fbold, fitalic, funder, fcolor = _run_format_sample(rn, para=para)
        if lvl == 1 and not h1_captured:
            h1_captured = True
            h1_font, h1_size, h1_bold, h1_italic, h1_underline, h1_color = (
                fname, fsize, fbold, fitalic, funder, fcolor
            )
        elif lvl == 2 and not h2_captured:
            h2_captured = True
            h2_font, h2_size, h2_bold, h2_italic, h2_underline, h2_color = (
                fname, fsize, fbold, fitalic, funder, fcolor
            )
        elif lvl == 3 and not h3_captured:
            h3_captured = True
            h3_font, h3_size, h3_bold, h3_italic, h3_underline, h3_color = (
                fname, fsize, fbold, fitalic, funder, fcolor
            )

    fc = FontConfig()
    if h1_font:
        fc.h1_font = h1_font
    if h1_size:
        fc.h1_size = h1_size
    if h1_bold is not None:
        fc.h1_bold = bool(h1_bold)
    if h1_italic is not None:
        fc.h1_italic = bool(h1_italic)
    if h1_underline is not None:
        fc.h1_underline = bool(h1_underline)
    if h1_color:
        fc.h1_color = h1_color
    if h2_font:
        fc.h2_font = h2_font
    if h2_size:
        fc.h2_size = h2_size
    if h2_bold is not None:
        fc.h2_bold = bool(h2_bold)
    if h2_italic is not None:
        fc.h2_italic = bool(h2_italic)
    if h2_underline is not None:
        fc.h2_underline = bool(h2_underline)
    if h2_color:
        fc.h2_color = h2_color
    if h3_font:
        fc.h3_font = h3_font
    if h3_size:
        fc.h3_size = h3_size
    if h3_bold is not None:
        fc.h3_bold = bool(h3_bold)
    if h3_italic is not None:
        fc.h3_italic = bool(h3_italic)
    if h3_underline is not None:
        fc.h3_underline = bool(h3_underline)
    if h3_color:
        fc.h3_color = h3_color
    if body_font:
        fc.body_font = body_font
    if body_size:
        fc.body_size = body_size
    if body_bold is not None:
        fc.body_bold = bool(body_bold)
    if body_italic is not None:
        fc.body_italic = bool(body_italic)
    if body_underline is not None:
        fc.body_underline = bool(body_underline)
    if body_color:
        fc.body_color = body_color
    if body_line_spacing is not None:
        fc.body_line_spacing = float(body_line_spacing)

    # 封面五行惯例：period, title, subtitle, org, date
    period = cover_lines[0] if len(cover_lines) > 0 else ""
    title = cover_lines[1] if len(cover_lines) > 1 else ""
    subtitle = cover_lines[2] if len(cover_lines) > 2 else ""
    org = cover_lines[3] if len(cover_lines) > 3 else ""
    date = cover_lines[4] if len(cover_lines) > 4 else ""

    return FormatTemplate(
        specs=specs,
        font_config=fc,
        period=period,
        title=title,
        subtitle=subtitle,
        org=org,
        date=date,
        source_path=os.path.abspath(docx_path),
        use_template_titles=True,
    )


# 可选参考模板（默认不使用；aggregate() 的 template 参数传入时才生效）
TEMPLATE: List[HeadingSpec] = [
    HeadingSpec("1", "报告概述", 1),
    HeadingSpec("2", "市场宏观环境分析", 1),
    HeadingSpec("2.1", "经济政策环境", 2, parent="2"),
    HeadingSpec("2.1.1", "宏观经济运行情况", 3, parent="2.1"),
    HeadingSpec("2.1.2", "保险行业政策动态", 3, parent="2.1"),
    HeadingSpec("2.2", "利率与资金面", 2, parent="2"),
    HeadingSpec("2.2.1", "利率走势分析", 3, parent="2.2"),
    HeadingSpec("2.3", "区域市场发展", 2, parent="2"),
    HeadingSpec("2.3.1", "区域保费分布特征", 3, parent="2.3"),
    HeadingSpec("3", "保费收入与增长态势", 1),
    HeadingSpec("3.1", "整体保费规模", 2, parent="3"),
    HeadingSpec("3.1.1", "原保险保费收入", 3, parent="3.1"),
    HeadingSpec("3.1.2", "赔付支出情况", 3, parent="3.1"),
    HeadingSpec("3.2", "分险种保费结构", 2, parent="3"),
    HeadingSpec("3.2.1", "财产险保费分析", 3, parent="3.2"),
    HeadingSpec("3.3", "中小险企市场份额", 2, parent="3"),
    HeadingSpec("4", "产品创新与业务结构", 1),
    HeadingSpec("4.1", "传统保险产品创新", 2, parent="4"),
    HeadingSpec("4.2", "新兴保险产品发展", 2, parent="4"),
    HeadingSpec("4.3", "产品定价策略变化", 2, parent="4"),
    HeadingSpec("4.4", "再保险市场动态", 2, parent="4"),
    HeadingSpec("5", "销售渠道分析", 1),
    HeadingSpec("5.1", "线下渠道", 2, parent="5"),
    HeadingSpec("5.1.1", "个险渠道分析", 3, parent="5.1"),
    HeadingSpec("5.1.2", "银保渠道分析", 3, parent="5.1"),
    HeadingSpec("5.2", "线上渠道", 2, parent="5"),
    HeadingSpec("5.2.1", "互联网保险平台", 3, parent="5.2"),
    HeadingSpec("6", "风险管理与理赔服务", 1),
    HeadingSpec("6.1", "风险评估体系", 2, parent="6"),
    HeadingSpec("6.2", "理赔服务效率", 2, parent="6"),
    HeadingSpec("6.3", "反欺诈技术应用", 2, parent="6"),
    HeadingSpec("7", "行业竞争格局", 1),
    HeadingSpec("8", "趋势展望与策略建议", 1),
]


# ---------------------------------------------------------------------------
# 标题识别（多信号：Word 样式 + 文本模式，任意深度编号 + 全角/中文支持）
# ---------------------------------------------------------------------------

# 文本模式正则
#   - 任意深度点分编号："2.1.1.1  标题"（旧版只支持 1~2 个点，现在不限）
#   - 一级阿拉伯编号："1  标题" / "1. 标题"（旧版要求至少一个点，会漏掉一级标题）
#   - 中文数字章节："第一章  标题" / "第十二章 标题"
#   - 阿拉伯章节："第1章  标题" / "第12章 标题"
#   - 全角点号兼容："２．１  标题"（全角数字 + 全角句点）
# 编号与标题之间允许：空格 / 制表符 / 全角空格 / 半角点 / 全角点 / 顿号 / 冒号 / 全角冒号 / 短横
_SEP = r"[\s\u3000\.\．\、:：\-]*"
RE_NUMERIC_DEEP = re.compile(
    r"^\s*(\d+(?:\.\d+)+)" + _SEP + r"(\S.*)$"
)
# 一级阿拉伯编号：编号与标题之间必须有分隔符（空格/点/顿号等），
# 禁止 "10年期国债…" / "7月保费…" / "2026年第三季度…" 这类正文被当成一级标题。
RE_NUMERIC_TOP = re.compile(
    r"^\s*(\d+)[\s\u3000\.\．\、:：\-]+(\S.*)$"
)
RE_CHAPTER = re.compile(r"^\s*第([一二三四五六七八九十百]+)章" + _SEP + r"(\S.*)$")
RE_CHAPTER_ARAB = re.compile(r"^\s*第(\d+)章" + _SEP + r"(\S.*)$")

# 正文时间单位开头：编号被剥离后标题若以这些开头，多半是 "N年/N月…" 正文。
# 注意不要用单字「分/时/秒」——会误伤「分险种保费结构」等合法标题。
_RE_TITLE_TIME_UNIT = re.compile(r"^(年|月|日|季度)")
# 过长且含句读的一级标题 → 正文段落误识别
_RE_PROSE_PUNCT = re.compile(r"[，。；、]")


def _plausible_heading_match(sid: str, title: str) -> bool:
    """过滤编号标题的常见误识别。

    典型误判：
      - "10年期国债收益率…" → sid=10, title=年期国债…
      - "7月保费收入…"     → sid=7,  title=月保费…
      - "2026年第三季度…"  → sid=2026, title=年第三季度…
      - "1年期LPR…"        → sid=1,  title=年期LPR…（污染真正的第一章）
    """
    if not sid or not title:
        return False
    # 年份（19xx/20xx）几乎不可能是报告章节号
    if re.fullmatch(r"(19|20)\d{2}", sid):
        return False
    parts = sid.split(".")
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return False
    # 章节号过大基本不是报告结构
    if nums[0] > 30:
        return False
    if any(n > 50 for n in nums):
        return False
    # "N年/N月/N日…" 正文：编号与单位粘连被切开后标题以时间单位开头
    if _RE_TITLE_TIME_UNIT.match(title):
        return False
    # 一级标题过长且带句读 → 多半是整段正文被误切
    if len(parts) == 1 and len(title) > 40 and _RE_PROSE_PUNCT.search(title):
        return False
    return True

# 全角数字 → 半角，用于兼容 "２．１ 标题" 这类输入
_FULLWIDTH_DIGIT = str.maketrans("０１２３４５６７８９", "0123456789")
# 全角点号 → 半角点，便于统一走 RE_NUMERIC_DEEP
_FULLWIDTH_DOT = str.maketrans("．。・", "...")

CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6,
          "七": 7, "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12,
          "十三": 13, "十四": 14, "十五": 15, "十六": 16, "十七": 17,
          "十八": 18, "十九": 19, "二十": 20, "二十一": 21, "二十二": 22,
          "二十三": 23, "二十四": 24, "二十五": 25}


def _cn_to_int(cn: str) -> Optional[int]:
    """中文数字 → int，支持 1-99（含'十'的组合，如'二十三'）。"""
    if cn in CN_NUM:
        return CN_NUM[cn]
    # 解析"二十三"这类组合
    if "十" in cn:
        parts = cn.split("十")
        tens = CN_NUM.get(parts[0], 1) if parts[0] else 1
        ones = CN_NUM.get(parts[1], 0) if len(parts) > 1 and parts[1] else 0
        return tens * 10 + ones if ones else tens * 10
    return None


# 保留旧正则别名，避免外部引用断裂（如 test 脚本）
RE_NUMERIC = RE_NUMERIC_DEEP


@dataclass
class HeadingInfo:
    """标题识别结果。

    confidence 取值：
      - "high"  : 文本编号 + Word 标题样式/大纲级别双重确认
      - "medium": 仅文本编号命中（无样式信号）
      - "low"   : 仅样式信号命中但文本无编号（可能是无编号标题，需人工确认）
    source_signal 记录命中的信号，用于日志审计。
    """
    section_id: str
    title: str
    level: int
    confidence: str = "medium"
    source_signal: str = "text"


def _style_is_heading(style_name: str) -> Tuple[bool, Optional[int]]:
    """判断样式名是否为 Word 内置标题样式，返回 (是否标题, 层级)。

    支持 "Heading 1" / "标题 1" / "Heading 2" 等中英文写法（1–9 级）。
    """
    if not style_name:
        return False, None
    s = style_name.strip().lower()
    for lvl in range(1, 10):
        for pat in (f"heading {lvl}", f"标题 {lvl}", f"标题{lvl}"):
            if s == pat:
                return True, min(lvl, 6)  # 终版格式最多用到六级
    return False, None


def _outline_level(para) -> Optional[int]:
    """从段落 XML 读取 w:outlineLvl，返回 1-based 层级（0→1），无则 None。"""
    pPr = para._element.find(qn('w:pPr')) if hasattr(para, '_element') else None
    if pPr is None:
        return None
    lvl_elem = pPr.find(qn('w:outlineLvl'))
    if lvl_elem is None:
        return None
    val = lvl_elem.get(qn('w:val'))
    if val is None:
        return None
    try:
        return int(val) + 1  # outlineLvl 是 0-based
    except ValueError:
        return None


def identify_heading(text: str, para: Optional[Paragraph] = None) -> Optional[HeadingInfo]:
    """多信号标题识别。返回 HeadingInfo 或 None。

    与旧版兼容：调用方可继续把它当 truthy 用；如需 sid 字符串取 .section_id。

    识别策略（综合文本 + Word 样式）：
      文本信号（任一命中即得 section_id）：
        - 第N章 / 第NN章              → 'N'           （中文/阿拉伯，N 1-99）
        - 2.1 / 2.1.1.1 / 3.2.1        → '2.1' 等       （任意深度点分）
        - 1  标题 / 1. 标题            → '1'            （一级阿拉伯，旧版会漏）
      样式信号：
        - Heading N / 标题 N / w:outlineLvl
      置信度：
        - 文本+样式双命中 → high
        - 仅文本命中      → medium
        - 仅样式命中      → low（无编号标题，需人工确认，不参与自动排序）
    """
    if not text:
        return None
    # 全角→半角归一化，统一后续匹配
    t = text.strip().translate(_FULLWIDTH_DIGIT).translate(_FULLWIDTH_DOT)
    if not t:
        return None

    sid: Optional[str] = None
    title: Optional[str] = None
    text_signal = ""

    # 1) 第N章（中文）
    m = RE_CHAPTER.match(t)
    if m:
        cn, raw_title = m.groups()
        n = _cn_to_int(cn)
        if n is not None and 1 <= n <= 99:
            cand_sid, cand_title = str(n), raw_title.strip()
            if _plausible_heading_match(cand_sid, cand_title):
                sid, title = cand_sid, cand_title
                text_signal = f"第N章(中文,第{cn}章)"
    # 2) 第N章（阿拉伯）
    if sid is None:
        m = RE_CHAPTER_ARAB.match(t)
        if m:
            n_str, raw_title = m.groups()
            n = int(n_str)
            if 1 <= n <= 99:
                cand_sid, cand_title = str(n), raw_title.strip()
                if _plausible_heading_match(cand_sid, cand_title):
                    sid, title = cand_sid, cand_title
                    text_signal = f"第N章(阿拉伯,第{n}章)"
    # 3) 多级点分编号（2.1 / 2.1.1.1 ...）—— 放在一级之前，避免被一级吞掉
    if sid is None:
        m = RE_NUMERIC_DEEP.match(t)
        if m:
            cand_sid = m.group(1)
            cand_title = m.group(2).strip()
            if _plausible_heading_match(cand_sid, cand_title):
                sid, title = cand_sid, cand_title
                text_signal = f"点分编号({sid})"
    # 4) 一级阿拉伯（1  标题 / 1. 标题）—— 必须有分隔符，避免 "10年期…" 误判
    if sid is None:
        m = RE_NUMERIC_TOP.match(t)
        if m:
            cand_sid = m.group(1)
            cand_title = m.group(2).strip()
            # 防止 "3.2  标题" 被切成 sid=3 / title="2  标题"
            # （点分多级应由 RE_NUMERIC_DEEP 处理；此处若落到一级说明 deep 未采纳）
            rest = t.lstrip()[len(cand_sid):]
            if re.match(r"[\.．]\d", rest):
                pass
            elif _plausible_heading_match(cand_sid, cand_title):
                sid, title = cand_sid, cand_title
                text_signal = f"一级编号({sid})"

    # 样式信号
    style_name = ""
    style_level: Optional[int] = None
    outline_level: Optional[int] = None
    if para is not None:
        style_name = para.style.name if para.style else ""
        is_h, sl = _style_is_heading(style_name)
        if is_h:
            style_level = sl
        outline_level = _outline_level(para)
    style_hit_level = style_level or outline_level

    # 综合置信度
    if sid is not None and title:
        level = min(sid.count('.') + 1, 6)
        if style_hit_level is not None:
            confidence = "high"
            signal = f"{text_signal}+样式(H{style_hit_level})"
            # 样式层级与编号层级不一致时仍以编号为准，但在信号里标注
            if style_hit_level != level:
                signal += f"[样式层级{style_hit_level}≠编号层级{level}]"
        else:
            confidence = "medium"
            signal = text_signal
        return HeadingInfo(
            section_id=sid, title=title, level=level,
            confidence=confidence, source_signal=signal,
        )

    # 仅样式命中（无编号标题）——返回 low，section_id 用占位，调用方需人工确认
    if sid is None and style_hit_level is not None:
        return HeadingInfo(
            section_id=f"_styled_{style_hit_level}",
            title=t, level=style_hit_level,
            confidence="low",
            source_signal=f"仅样式(H{style_hit_level},{style_name})",
        )

    return None


def _level_from_sid(sid: str) -> int:
    """从 section_id 推断层级：'1'→1, '2.1'→2, '2.1.1'→3, … 最深 6。"""
    if not sid or sid.startswith("_"):
        return 1
    return min(sid.count(".") + 1, 6)


# ---------------------------------------------------------------------------
# 解析输入文件
# ---------------------------------------------------------------------------

@dataclass
class RunInfo:
    """单个 run 的完整格式属性（F06：取代旧版 (text, is_bold) 元组）。

    保留 italic/underline/color/size/font_name，避免输出时丢失字符格式。
    color 为 6 位十六进制字符串（如 "1F3864"），None 表示不指定（继承）。
    F07：bold/italic/underline 改 Optional[bool]，None=源未指定（与 False 区分），
    让 _apply_run_info 能用 FontConfig 默认值兜底"未指定"的格式。
    """
    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color: Optional[str] = None      # "RRGGBB"
    size: Optional[float] = None     # pt
    font_name: Optional[str] = None  # None 表示继承


@dataclass
class ParaFormat:
    """段落格式属性（F06：保留对齐/缩进/行距/段前段后）。"""
    alignment: Optional[str] = None   # "left"/"center"/"right"/"justify"/None
    first_line_indent: Optional[float] = None  # 字符数
    left_indent: Optional[float] = None        # 字符数
    line_spacing: Optional[float] = None       # 倍数
    space_before: Optional[float] = None       # pt
    space_after: Optional[float] = None        # pt


@dataclass
class Block:
    """section 内的一个内容块。

    kind 取值：
      - "p"        : 普通文本段落
      - "bullet"   : 列表项（无序，原 List Bullet）
      - "number"   : 编号列表项（F06：w:numPr 识别，含 num_id/list_level）
      - "table"    : 表格（table 字段持有源 Table 对象）
      - "image"    : 含图片的段落（src_para_elem 持有源段落 XML，src_doc 持有源 Document）
      - "rich"     : 含超链接/外链图的段落（F03，走保真复制+关系迁移，同 image 路径）

    F06：runs 字段从 (text, is_bold) 升级为 RunInfo 列表，保留 italic/underline/
    color/size/font；para_format 保留段落对齐/缩进/行距；num_id/list_level 记录
    编号列表语义（取代仅认 "List Bullet" 英文样式名）。
    source_file 记录该块来自哪个输入文件，用于合并审计与守恒校验。
    """
    kind: str
    text: str = ""
    table: Optional[Table] = None
    style_name: str = "Normal"
    src_para_elem: Optional[object] = None   # lxml element
    src_doc: Optional[_Document] = None      # 源 Document 对象（用于查找图片关系）
    runs: List[RunInfo] = field(default_factory=list)  # F06：RunInfo 替代 (text, is_bold)
    source_file: str = ""                    # 来源文件名（审计用）
    section_id: str = ""                     # 所属 section_id（F04 题注作用域用）
    # F06：段落格式与列表语义
    para_format: Optional[ParaFormat] = None
    num_id: Optional[int] = None             # w:numPr 的 numId（编号列表）
    list_level: int = 0                      # w:numPr 的 ilvl（列表层级，0=顶层）


@dataclass
class Section:
    """从输入文件抽取出来的一个 section 单元。

    多个输入文件中 section_id 相同的 section 会被合并到同一个 Section 实例
    （在 parse_input_dir 中完成），不再静默丢弃：
      - 同标题：按文件顺序追加 blocks，merge_count 累加
      - 不同标题：仍追加 blocks，但置 title_conflict=True，titles_seen 记录全部标题
    source_files 记录所有来源文件名；merge_count=1 表示未发生合并。
    level/confidence/source_signal 来自首个识别该 section 的 HeadingInfo，
    供 aggregate 写入循环与 UI 预警使用。
    """
    section_id: str
    title: str
    blocks: List[Block] = field(default_factory=list)
    source_files: List[str] = field(default_factory=list)
    merge_count: int = 1
    title_conflict: bool = False
    titles_seen: List[str] = field(default_factory=list)
    level: int = 1
    confidence: str = "medium"
    source_signal: str = "text"
    is_preamble: bool = False       # F05：首标题前的前导内容（前言/摘要/免责声明等）


def _para_has_drawing(para_elem) -> bool:
    """检查段落 XML 元素是否包含 w:drawing（即嵌入图片）。"""
    return len(para_elem.findall('.//' + qn('w:drawing'))) > 0


def _para_has_relations(para_elem) -> bool:
    """检查段落是否含 OOXML 关系引用（超链接、外链图等），需走保真复制路径。

    F03：含 w:hyperlink/@r:id 或 a:blip/@r:link 的段落不能简单用 runs 文本重建，
    否则超链接和外链图关系会丢失。这类段落改走深拷贝+关系迁移。
    """
    # 超链接
    if para_elem.findall('.//' + qn('w:hyperlink')):
        return True
    # 外链图（a:blip 的 r:link，区别于嵌入图的 r:embed）
    for blip in para_elem.findall('.//' + qn('a:blip')):
        if blip.get(qn('r:link')) is not None:
            return True
    return False


def _extract_run_info(run) -> RunInfo:
    """从 python-docx run 提取完整格式属性（F06/F07）。

    F07：bold/italic/underline 保留 python-docx 原值（None=未指定，True/False=显式），
    不再用 bool() 压平，让 _apply_run_info 能区分"源未指定"与"源显式不加粗"。
    """
    font = run.font
    color_hex = None
    try:
        if font.color and font.color.rgb is not None:
            color_hex = str(font.color.rgb)
    except Exception:
        pass
    size = None
    try:
        if font.size is not None:
            size = font.size.pt
    except Exception:
        pass
    return RunInfo(
        text=run.text or "",
        bold=font.bold,          # None/True/False 保留
        italic=font.italic,
        underline=font.underline,
        color=color_hex,
        size=size,
        font_name=font.name,
    )


def _extract_num_pr(para) -> Tuple[Optional[int], int, bool]:
    """从段落提取 w:numPr（F06：列表编号属性）。

    返回 (num_id, list_level, is_numbered)。
    is_numbered=True 表示该段落是编号/项目符号列表项（通过 numPr 判定，
    取代旧版仅认 "List Bullet" 英文样式名的窄规则）。
    """
    pPr = para._element.find(qn('w:pPr')) if hasattr(para, '_element') else None
    if pPr is None:
        return None, 0, False
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        return None, 0, False
    num_id = None
    ilvl = 0
    num_id_elem = numPr.find(qn('w:numId'))
    if num_id_elem is not None and num_id_elem.get(qn('w:val')):
        num_id = int(num_id_elem.get(qn('w:val')))
    ilvl_elem = numPr.find(qn('w:ilvl'))
    if ilvl_elem is not None and ilvl_elem.get(qn('w:val')):
        ilvl = int(ilvl_elem.get(qn('w:val')))
    return num_id, ilvl, True


def _extract_para_format(para) -> ParaFormat:
    """提取段落格式属性（F06：对齐/缩进/行距/段前段后）。"""
    pf = para.paragraph_format
    align_map = {
        None: None,
    }
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _A
        align_map = {None: None, _A.LEFT: "left", _A.CENTER: "center",
                     _A.RIGHT: "right", _A.JUSTIFY: "justify"}
    except Exception:
        pass
    alignment = align_map.get(pf.alignment, None)

    first_line = None
    left_indent = None
    try:
        if pf.first_line_indent is not None:
            first_line = pf.first_line_indent.pt / 12.0  # 粗→字符近似
        if pf.left_indent is not None:
            left_indent = pf.left_indent.pt / 12.0
    except Exception:
        pass

    line_spacing = None
    try:
        if pf.line_spacing is not None:
            line_spacing = float(pf.line_spacing)
    except Exception:
        pass

    space_before = space_after = None
    try:
        if pf.space_before is not None:
            space_before = pf.space_before.pt
        if pf.space_after is not None:
            space_after = pf.space_after.pt
    except Exception:
        pass

    return ParaFormat(
        alignment=alignment, first_line_indent=first_line,
        left_indent=left_indent, line_spacing=line_spacing,
        space_before=space_before, space_after=space_after,
    )


def _iter_body_items(doc: _Document):
    """按文档顺序产出段落和表格元素。"""
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def parse_input_file(path: str, source_file: str = "") -> List[Section]:
    """从一个输入 docx 中解析出全部 section（按出现顺序）。

    source_file 用于给每个 Block/Section 打上来源文件标记，便于合并审计与
    守恒校验；省略时取 path 的 basename。

    F10：本函数假定文件已通过 _validate_docx_package 静态检查。若调用方
    未先校验，此处仍会因 Document(path) 抛出异常——请改用 parse_input_dir，
    它对每个文件做了"预检查 + 异常分类"的隔离。
    """
    if not source_file:
        source_file = os.path.basename(path)
    doc = Document(path)
    sections: List[Section] = []
    current: Optional[Section] = None
    # F05：首标题前的前导内容（摘要/引言/免责声明/前言等）单独收集，不再丢弃。
    # 用特殊 section_id=_preamble 的 Section 承载，aggregate 会把它作为"前言"输出。
    preamble: Optional[Section] = None

    def _ensure_preamble():
        """惰性创建 preamble Section（仅当确有前导内容时才创建）。"""
        nonlocal preamble
        if preamble is None:
            preamble = Section(
                section_id="_preamble", title="前言",
                level=1, is_preamble=True,
            )
            preamble.source_files = [source_file]
            preamble.titles_seen = ["前言"]
            sections.append(preamble)
        return preamble

    for item in _iter_body_items(doc):
        if isinstance(item, Paragraph):
            text = item.text
            info = identify_heading(text, para=item)
            style = item.style.name if item.style else "Normal"

            if info is not None:
                # 低置信度标题（仅样式命中、无编号）：不新建 section，内容并入当前
                # section，避免丢数据；标记该 block 以便日志/UI 提示人工确认。
                if info.confidence == "low":
                    target = current if current is not None else _ensure_preamble()
                    target.blocks.append(Block(
                        kind="p", text=text.strip(), style_name=style,
                        source_file=source_file,
                        section_id=target.section_id,
                    ))
                    continue
                # 新 section 开始。直接用识别器给出的 title，不再二次正则切，
                # 避免"识别命中但标题切错"的不一致（旧版 F02 旁的隐患）。
                current = Section(
                    section_id=info.section_id, title=info.title,
                    level=info.level, confidence=info.confidence,
                    source_signal=info.source_signal,
                )
                current.source_files = [source_file]
                current.titles_seen = [info.title]
                sections.append(current)
                continue

            stripped = text.strip()

            # F05：首标题之前的内容（current is None）不再 continue 丢弃，
            # 而是并入 preamble Section，保证摘要/引言/免责声明等前导内容可见。
            if current is None:
                # 检查段落是否包含图片
                if _para_has_drawing(item._element):
                    psec = _ensure_preamble()
                    psec.blocks.append(Block(
                        kind="image", text=stripped, style_name=style,
                        src_para_elem=item._element, src_doc=doc,
                        source_file=source_file, section_id="_preamble",
                    ))
                    continue
                if not stripped:
                    continue
                if _para_has_relations(item._element):
                    psec = _ensure_preamble()
                    psec.blocks.append(Block(
                        kind="rich", text=stripped, style_name=style,
                        src_para_elem=item._element, src_doc=doc,
                        source_file=source_file, section_id="_preamble",
                    ))
                    continue
                # F06：提取完整 run 属性 + 段落格式 + 列表语义
                runs = [_extract_run_info(r) for r in item.runs if r.text]
                num_id, list_level, is_list = _extract_num_pr(item)
                pfmt = _extract_para_format(item)
                if is_list:
                    kind = "number" if num_id is not None else "bullet"
                else:
                    kind = "bullet" if style == "List Bullet" else "p"
                psec = _ensure_preamble()
                psec.blocks.append(Block(
                    kind=kind, text=stripped, style_name=style, runs=runs,
                    source_file=source_file, section_id="_preamble",
                    para_format=pfmt, num_id=num_id, list_level=list_level,
                ))
                continue

            # 以下为已有 current（正常章节内）的处理
            # 检查段落是否包含图片
            if _para_has_drawing(item._element):
                # 图片段落：保存整个段落 XML + 源文档引用
                current.blocks.append(Block(
                    kind="image",
                    text=stripped,
                    style_name=style,
                    src_para_elem=item._element,
                    src_doc=doc,
                    source_file=source_file,
                    section_id=current.section_id,
                ))
                continue

            if not stripped:
                continue

            # F03：含超链接/外链图的段落走保真复制路径，避免关系丢失
            if _para_has_relations(item._element):
                current.blocks.append(Block(
                    kind="rich",
                    text=stripped,
                    style_name=style,
                    src_para_elem=item._element,
                    src_doc=doc,
                    source_file=source_file,
                    section_id=current.section_id,
                ))
                continue

            # F06：捕获完整 run 格式（italic/underline/color/size/font）+ 段落格式 + 列表语义
            runs = [_extract_run_info(r) for r in item.runs if r.text]
            num_id, list_level, is_list = _extract_num_pr(item)
            pfmt = _extract_para_format(item)
            # F06：用 w:numPr 判定列表（取代仅认 "List Bullet" 英文样式名）
            if is_list:
                kind = "number" if num_id is not None else "bullet"
            else:
                kind = "bullet" if style == "List Bullet" else "p"
            current.blocks.append(Block(
                kind=kind, text=stripped, style_name=style, runs=runs,
                source_file=source_file,
                section_id=current.section_id,
                para_format=pfmt, num_id=num_id, list_level=list_level,
            ))

        elif isinstance(item, Table):
            # F05：首标题前的表格也进 preamble，不再丢弃
            target = current if current is not None else _ensure_preamble()
            target.blocks.append(Block(
                kind="table", table=item, source_file=source_file,
                section_id=target.section_id,
            ))
    return sections


def parse_input_dir(input_dir: str) -> ParseResult:
    """合并多个输入文件中识别出的 section。

    同 section_id 的 section 会被合并到同一个 Section 实例（按文件顺序追加
    blocks），不再静默丢弃：
      - 同标题：直接追加 blocks，merge_count 累加
      - 不同标题：仍追加 blocks，但置 title_conflict=True，titles_seen 记录
        全部标题，供 aggregate 日志告警与人工核对

    合并详情全部挂在返回的 ParseResult.sections 中各 Section 实例上
    （source_files / merge_count / title_conflict / titles_seen），调用方
    可直接读取审计，无需额外通道。

    F10：返回 ParseResult（成功文件、警告、失败文件三部分）。坏文件不再使
    整个流程崩溃——每个文件独立预检查 + 异常分类，失败文件进
    result.failures，成功文件继续参与合并。调用方据 result.has_failures()
    决定是否禁止生成。

    向后兼容：ParseResult 实现 __iter__/__getitem__/keys()/values()/items()/
    get()/__contains__/__bool__/__len__，旧代码 `for sec in parsed.values()`
    或 `parsed[sid]` 无需修改即可工作。
    """
    merged: Dict[str, Section] = {}
    warnings: List[FileWarning] = []
    failures: List[FileFailure] = []
    parsed_files: List[str] = []

    file_list = sorted(
        f for f in os.listdir(input_dir) if f.lower().endswith(".docx")
    )
    for fname in file_list:
        path = os.path.join(input_dir, fname)

        # ---- F10 预检查：ZIP 魔数、必需 OOXML 部件、文件大小 ----
        pre_fail = _validate_docx_package(path)
        if pre_fail is not None:
            failures.append(pre_fail)
            continue

        # ---- F10 异常分类：Document(path) / 解析过程按文件捕获 ----
        try:
            secs = parse_input_file(path, source_file=fname)
        except (zipfile.BadZipFile, KeyError, OSError,
                etree.XMLSyntaxError, ValueError, LookupError, Exception) as exc:
            # 兜底 Exception：python-docx 内部可能抛各种子类，统一归类
            failures.append(_classify_parse_error(exc, path))
            continue

        parsed_files.append(fname)
        for sec in secs:
            if sec.section_id not in merged:
                merged[sec.section_id] = sec
            else:
                existing = merged[sec.section_id]
                # 按文件顺序追加内容，绝不丢弃
                existing.blocks.extend(sec.blocks)
                existing.source_files.extend(sec.source_files)
                existing.merge_count += 1
                # 标题冲突检测：同编号出现不同标题时标记，便于人工核对
                if sec.title not in existing.titles_seen:
                    existing.titles_seen.append(sec.title)
                    existing.title_conflict = True

    return ParseResult(
        sections=merged, warnings=warnings,
        failures=failures, parsed_files=parsed_files,
    )


# ---------------------------------------------------------------------------
# 排版常量
# ---------------------------------------------------------------------------

FONT_HEADING = "黑体"
FONT_BODY = "宋体"
COLOR_HEADING = RGBColor(0x1F, 0x38, 0x64)    # 深蓝（H2 标题用）
COLOR_BLACK = RGBColor(0x00, 0x00, 0x00)     # 黑色（H1/H3 标题用）
COLOR_SUBTITLE = RGBColor(0x66, 0x66, 0x66)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SIZE_COVER_TITLE = Pt(26)
SIZE_COVER_SUB = Pt(14)
SIZE_H1 = Pt(18)     # 小二
SIZE_H2 = Pt(15)     # 小三
SIZE_H3 = Pt(14)     # 四号
SIZE_BODY = Pt(12)   # 小四
SIZE_CAPTION = Pt(10)
SIZE_HEADER = Pt(9)            # 页眉页脚字号

# 页眉页脚
COLOR_HEADER = RGBColor(0x99, 0x99, 0x99)   # 页眉文字浅灰
COLOR_HEADER_BORDER = "1F3864"              # 页眉下划线深蓝

# 表格颜色
TABLE_HEADER_FILL = "1F3864"   # 深蓝（首行）
TABLE_ALT_FILL = "DCE6F2"      # 浅蓝（隔行）


# ---------------------------------------------------------------------------
# 字体格式配置（可通过 App UI 自定义）
# ---------------------------------------------------------------------------

@dataclass
class FontConfig:
    """各级标题和正文的字体格式配置。

    所有颜色用 6 位十六进制字符串表示，如 "1F3864"。
    字号为磅值（pt），如 18 = 小二，15 = 小三，14 = 四号，12 = 小四。
    """

    # 一级标题
    h1_font: str = "黑体"
    h1_size: float = 18.0       # 小二
    h1_bold: bool = True
    h1_italic: bool = False
    h1_underline: bool = False
    h1_color: str = "000000"   # 黑色

    # 二级标题
    h2_font: str = "黑体"
    h2_size: float = 15.0       # 小三
    h2_bold: bool = True
    h2_italic: bool = False
    h2_underline: bool = False
    h2_color: str = "1F3864"   # 深蓝

    # 三级标题
    h3_font: str = "宋体"
    h3_size: float = 14.0       # 四号
    h3_bold: bool = True
    h3_italic: bool = False
    h3_underline: bool = False
    h3_color: str = "000000"    # 黑色

    # 四级标题
    h4_font: str = "宋体"
    h4_size: float = 12.0       # 小四
    h4_bold: bool = True
    h4_italic: bool = False
    h4_underline: bool = False
    h4_color: str = "000000"

    # 五级标题
    h5_font: str = "宋体"
    h5_size: float = 12.0
    h5_bold: bool = True
    h5_italic: bool = False
    h5_underline: bool = False
    h5_color: str = "000000"

    # 六级标题
    h6_font: str = "宋体"
    h6_size: float = 11.0
    h6_bold: bool = True
    h6_italic: bool = False
    h6_underline: bool = False
    h6_color: str = "000000"

    # 正文
    body_font: str = "宋体"
    body_size: float = 12.0      # 小四
    body_bold: bool = False
    body_italic: bool = False
    body_underline: bool = False
    body_color: str = "000000"
    body_indent: int = 2        # 首行缩进字符数
    body_line_spacing: float = 1.5  # 行距倍数


def _color_from_hex(hex_str: str) -> RGBColor:
    """将 6 位十六进制颜色字符串转为 RGBColor。"""
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


def _apply_font_config(fc: FontConfig):
    """将 FontConfig 应用到模块级排版常量（F13 已收敛危害）。

    所有格式函数在调用时引用模块级变量（而非默认参数），因此更新这些变量
    后，后续调用会自动使用新值。F13 之后：
      - 默认参数（_set_run_font / _ensure_style / _add_field_simple）已改为
        None + 运行时读取当前全局，消除"定义时绑定旧全局"的陈旧值危害；
        显式传参（如 _set_run_font(r, FONT_BODY, SIZE_H1)）本就在调用时
        求值，自动反映最新设置。
      - 书签 ID 计数器已改为 contextvars.ContextVar，保证多次调用 / 并行
        测试的可重入与线程安全。
    本函数仍以 `global` 写模块级常量——适用于 Streamlit 单进程单次聚合模型。
    若未来引入多线程并发聚合，需在此处加锁或将渲染配置一并显式化。
    """
    global FONT_HEADING, FONT_BODY
    global COLOR_HEADING, COLOR_BLACK
    global SIZE_H1, SIZE_H2, SIZE_H3, SIZE_BODY

    # 字体
    FONT_HEADING = fc.h1_font       # H1/H2 共用标题字体
    FONT_BODY = fc.body_font        # 正文字体（H3/H4 也用 body 字体族）

    # 字号
    SIZE_H1 = Pt(fc.h1_size)
    SIZE_H2 = Pt(fc.h2_size)
    SIZE_H3 = Pt(fc.h3_size)
    SIZE_BODY = Pt(fc.body_size)

    # 颜色
    COLOR_HEADING = _color_from_hex(fc.h2_color)   # 深蓝（H2 用）
    COLOR_BLACK = _color_from_hex(fc.h1_color)     # 黑色（H1/H3 用）


# ---------------------------------------------------------------------------
# 基础工具函数
# ---------------------------------------------------------------------------

def _set_run_font(run, name=None, size=None, bold=False, color=None, italic=False, underline=False):
    """设置 run 字体。

    F13：name/size 默认 None（不再在定义时绑定旧全局，避免 _apply_font_config
    更新模块级常量后默认参数仍用陈旧值）。运行时读取当前全局，反映本次聚合
    的 FontConfig 设置。
    """
    if name is None:
        name = FONT_BODY
    if size is None:
        size = SIZE_BODY
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    run.font.size = size
    run.font.bold = bool(bold)
    run.font.italic = bool(italic)
    run.font.underline = bool(underline)
    if color is not None:
        run.font.color.rgb = color


def _apply_run_info(run, ri: RunInfo, fc: FontConfig = None, default_bold: bool = None):
    """把 RunInfo 文本写入 run，版式严格按 FontConfig（模板或手动设置）。

    版式优先级（模板启用时 FontConfig 来自模板；否则来自 UI 手动设置）：
      - 字体 / 字号 / 颜色：一律用 FontConfig.body_*（统一终版风格）
      - 加粗 / 斜体 / 下划线：源 run 显式为 True 时保留强调；否则用 body_* 设定
    default_bold 参数：旧调用方兼容，若显式传入则覆盖 fc.body_bold。
    """
    if fc is None:
        fc = FontConfig()
    body_font = fc.body_font
    body_size = fc.body_size
    run.font.name = body_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), body_font)
    rfonts.set(qn("w:ascii"), body_font)
    rfonts.set(qn("w:hAnsi"), body_font)
    run.font.size = Pt(body_size)
    # 加粗：源显式加粗保留强调；否则用 default_bold / body_bold
    if ri.bold is True:
        run.font.bold = True
    elif default_bold is not None:
        run.font.bold = bool(default_bold)
    else:
        run.font.bold = bool(fc.body_bold)
    # 斜体 / 下划线：源显式开启时保留强调；否则用 body 设定
    if ri.italic is True:
        run.font.italic = True
    else:
        run.font.italic = bool(fc.body_italic)
    src_under = False
    if ri.underline is True:
        src_under = True
    elif ri.underline not in (None, False, 0):
        # WD_UNDERLINE 枚举等真值
        src_under = True
    run.font.underline = True if src_under else bool(fc.body_underline)
    # 颜色严格按 FontConfig（模板或手动「正文颜色」）
    color_to_apply = fc.body_color
    if color_to_apply:
        try:
            run.font.color.rgb = RGBColor.from_string(color_to_apply)
        except Exception:
            pass


def _apply_para_format(para, pfmt: Optional[ParaFormat]):
    """F06：把段落格式（对齐/缩进/行距/段前段后）应用到段落。"""
    if pfmt is None:
        return
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _A
    align_map = {"left": _A.LEFT, "center": _A.CENTER, "right": _A.RIGHT, "justify": _A.JUSTIFY}
    if pfmt.alignment and pfmt.alignment in align_map:
        para.alignment = align_map[pfmt.alignment]
    pf = para.paragraph_format
    if pfmt.first_line_indent is not None:
        _set_first_line_indent(para, chars=int(pfmt.first_line_indent) if pfmt.first_line_indent >= 1 else None)
    if pfmt.left_indent is not None:
        pf.left_indent = Pt(pfmt.left_indent * 12)
    if pfmt.line_spacing is not None:
        pf.line_spacing = pfmt.line_spacing
    if pfmt.space_before is not None:
        pf.space_before = Pt(pfmt.space_before)
    if pfmt.space_after is not None:
        pf.space_after = Pt(pfmt.space_after)


def _ensure_style(doc, name, base="Normal", font_name=None, size=None, bold=False, color=None):
    """确保文档中存在指定样式的样式定义。"""
    # F13：默认参数改为 None，运行时读取当前全局（避免定义时绑定陈旧值）
    if font_name is None:
        font_name = FONT_BODY
    if size is None:
        size = SIZE_BODY
    from docx.enum.style import WD_STYLE_TYPE
    styles = doc.styles
    try:
        st = styles[name]
    except KeyError:
        st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = styles[base]
    st.font.name = font_name
    st.font.size = size
    st.font.bold = bool(bold)
    if color is not None:
        st.font.color.rgb = color
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    return st


def _get_sectpr(doc):
    """获取 body 中的 sectPr 元素（文档末尾的 section 属性）。"""
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    return sectPr


def _insert_before_sectpr(doc, element):
    """在 sectPr 之前插入一个 XML 元素（保证内容顺序正确）。"""
    body = doc.element.body
    sectPr = _get_sectpr(doc)
    if sectPr is not None:
        sectPr.addprevious(element)
    else:
        body.append(element)


def _set_first_line_indent(paragraph, chars: int = 2, body_size_pt: float = None):
    """设置段落首行缩进 N 个中文字符。

    同时设置 w:firstLineChars（字符单位，Word 优先使用）和
    w:firstLine（磅值，兼容性回退）。
    body_size_pt 用于计算磅值回退，默认用当前 SIZE_BODY。
    """
    if body_size_pt is None:
        body_size_pt = SIZE_BODY.pt
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        _insert_pPr_child(pPr, ind)
    # 字符单位：100 = 1 字符
    ind.set(qn('w:firstLineChars'), str(chars * 100))
    # 磅值回退：chars * 字号 * 20 twips/pt
    twips = int(chars * body_size_pt * 20)
    ind.set(qn('w:firstLine'), str(twips))


# OOXML schema 中 w:pPr 子元素的正确顺序
_PPR_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
    'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
    'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
    'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'jc',
    'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
    'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
]


def _insert_pPr_child(pPr, element):
    """在 w:pPr 中按 OOXML schema 顺序插入子元素。

    避免 append 导致元素顺序违反 schema（如 pBdr 在 jc 之后），
    这会导致 Word 报错无法打开文件。
    """
    tag = etree.QName(element).localname
    if tag not in _PPR_ORDER:
        pPr.append(element)
        return
    target_idx = _PPR_ORDER.index(tag)
    for i, child in enumerate(list(pPr)):
        child_tag = etree.QName(child).localname
        if child_tag in _PPR_ORDER:
            child_idx = _PPR_ORDER.index(child_tag)
            if child_idx > target_idx:
                pPr.insert(i, element)
                return
    pPr.append(element)


# ---------------------------------------------------------------------------
# 书签 & 域（SEQ / REF / TOC）
# ---------------------------------------------------------------------------

# F13：书签 ID 计数器改为 contextvars.ContextVar——每个执行上下文（线程 /
# asyncio task / 测试并行）拥有独立副本，保证可重入和线程安全，杜绝多次
# aggregate 调用或并行测试间的计数器互相污染。替代旧版模块级可变列表计数器。
_bookmark_counter_var: "contextvars.ContextVar[int]" = contextvars.ContextVar(
    "bookmark_counter", default=0)


def _reset_bookmark_counter():
    """每次 aggregate() 调用前重置书签 ID 计数器（在当前上下文内）。"""
    _bookmark_counter_var.set(0)


def _add_bookmark_start(paragraph, name: str) -> int:
    """在段落中（pPr 之后）插入 bookmarkStart，返回书签 ID。"""
    bm_id = _bookmark_counter_var.get()
    _bookmark_counter_var.set(bm_id + 1)
    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), str(bm_id))
    bm_start.set(qn('w:name'), name)
    paragraph._element.append(bm_start)
    return bm_id


def _add_bookmark_end(paragraph, bm_id: int):
    """在段落末尾插入 bookmarkEnd。"""
    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), str(bm_id))
    paragraph._element.append(bm_end)


def _run_rpr_xml(size=None, bold: bool = False, font_name: Optional[str] = None):
    """构造 w:rPr（字号/字体/加粗），用于域指令与结果 run，保证编号与题注同字号。"""
    if size is None:
        size = SIZE_BODY
    if font_name is None:
        font_name = FONT_BODY
    try:
        half_pt = str(int(round(float(size.pt) * 2)))
    except Exception:
        half_pt = "20"  # 10pt
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rpr.append(rfonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), half_pt)
    rpr.append(sz)
    sz_cs = OxmlElement("w:szCs")
    sz_cs.set(qn("w:val"), half_pt)
    rpr.append(sz_cs)
    if bold:
        rpr.append(OxmlElement("w:b"))
    return rpr


def _add_field_simple(paragraph, instr: str, placeholder: str = "",
                      size=None, bold=False, font_name: Optional[str] = None):
    """在段落末尾追加复杂域（fldChar begin/instr/separate/result/end）。

    使用复杂域而非 w:fldSimple，便于在结果 run 上写入字号/字体，
    避免 Word 更新域后编号变成正文默认字号（比题注大一号）。

    instr       : 域指令，如 'SEQ Figure \\* ARABIC'
    placeholder : 域未更新时的占位文本（及初始结果）
    """
    if size is None:
        size = SIZE_BODY
    if font_name is None:
        font_name = FONT_BODY

    def _append_run(build_fn):
        r = paragraph.add_run()
        build_fn(r)
        # 每个域相关 run 都带同一 rPr，更新域后字号仍一致
        rpr = r._element.get_or_add_rPr()
        # 清空后写入标准 rPr 子节点
        for child in list(rpr):
            rpr.remove(child)
        new_rpr = _run_rpr_xml(size=size, bold=bold, font_name=font_name)
        for child in list(new_rpr):
            rpr.append(child)
        return r

    # begin
    r1 = _append_run(lambda r: None)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r1._element.append(fld_begin)

    # instrText
    r2 = _append_run(lambda r: None)
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    r2._element.append(instr_el)

    # separate
    r3 = _append_run(lambda r: None)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r3._element.append(fld_sep)

    # result / placeholder
    r4 = paragraph.add_run(placeholder or "")
    rpr4 = r4._element.get_or_add_rPr()
    for child in list(rpr4):
        rpr4.remove(child)
    new_rpr4 = _run_rpr_xml(size=size, bold=bold, font_name=font_name)
    for child in list(new_rpr4):
        rpr4.append(child)
    r4.font.size = size
    r4.font.name = font_name
    r4.font.bold = bool(bold)

    # end
    r5 = _append_run(lambda r: None)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r5._element.append(fld_end)
    return r4


# ---------------------------------------------------------------------------
# 封面 & 目录
# ---------------------------------------------------------------------------

def add_cover(doc, *, period: str, title: str, subtitle: str, org: str, date: str):
    """追加封面页（之后插入分页符）。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(120)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(period)
    _set_run_font(r, FONT_HEADING, SIZE_COVER_TITLE, bold=True, color=COLOR_HEADING)
    p.paragraph_format.space_after = Pt(8)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    _set_run_font(r, FONT_HEADING, SIZE_COVER_TITLE, bold=True, color=COLOR_HEADING)
    p.paragraph_format.space_after = Pt(36)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle)
    _set_run_font(r, FONT_BODY, SIZE_COVER_SUB, color=COLOR_SUBTITLE)
    p.paragraph_format.space_after = Pt(180)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(org)
    _set_run_font(r, FONT_BODY, SIZE_COVER_SUB)
    p.paragraph_format.space_after = Pt(6)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(date)
    _set_run_font(r, FONT_BODY, SIZE_COVER_SUB)
    p.paragraph_format.space_after = Pt(12)

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def add_toc(doc):
    """在封面后插入目录页：标题 + TOC 域 + 分页符。

    TOC 域使用 fldChar（begin / instrText / separate / placeholder / end）结构，
    用户在 Word 中按 Ctrl+A → F9 即可自动生成完整目录。
    """
    # 目录标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("目  录")
    _set_run_font(r, FONT_HEADING, SIZE_H1, bold=True, color=COLOR_BLACK)
    p.paragraph_format.space_after = Pt(12)

    # TOC 域 —— 使用独立 w:r 元素避免嵌套
    p2 = doc.add_paragraph()

    # Run 1: fldChar begin
    r1 = p2.add_run()
    fldBegin = OxmlElement('w:fldChar')
    fldBegin.set(qn('w:fldCharType'), 'begin')
    r1._element.append(fldBegin)
    _set_run_font(r1, FONT_BODY, SIZE_BODY)

    # Run 2: instrText
    r2 = p2.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    r2._element.append(instr)
    _set_run_font(r2, FONT_BODY, SIZE_BODY)

    # Run 3: fldChar separate
    r3 = p2.add_run()
    fldSep = OxmlElement('w:fldChar')
    fldSep.set(qn('w:fldCharType'), 'separate')
    r3._element.append(fldSep)
    _set_run_font(r3, FONT_BODY, SIZE_BODY)

    # Run 4: 占位文本
    r4 = p2.add_run('请在 Word 中按 Ctrl+A → F9 更新目录')
    _set_run_font(r4, FONT_BODY, SIZE_BODY)

    # Run 5: fldChar end
    r5 = p2.add_run()
    fldEnd = OxmlElement('w:fldChar')
    fldEnd.set(qn('w:fldCharType'), 'end')
    r5._element.append(fldEnd)
    _set_run_font(r5, FONT_BODY, SIZE_BODY)

    # 分页
    p3 = doc.add_paragraph()
    p3.add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------------------
# 页眉 & 页脚
# ---------------------------------------------------------------------------

# 页眉可选内容项
HEADER_OPTIONS = {
    "title":   "报告标题",
    "org":     "编制单位",
    "date":    "发布日期",
    "period":  "报告周期",
}


def _add_page_field(paragraph):
    """在段落中追加 PAGE 域（自动页码）。

    使用独立的 w:r 元素分别承载 fldChar / instrText / 占位文本，
    避免 w:r 嵌套导致 Word 重复渲染占位文字。
    """
    # Run 1: fldChar begin
    r1 = paragraph.add_run()
    fldBegin = OxmlElement('w:fldChar')
    fldBegin.set(qn('w:fldCharType'), 'begin')
    r1._element.append(fldBegin)
    _set_run_font(r1, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)

    # Run 2: instrText
    r2 = paragraph.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    r2._element.append(instr)
    _set_run_font(r2, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)

    # Run 3: fldChar separate
    r3 = paragraph.add_run()
    fldSep = OxmlElement('w:fldChar')
    fldSep.set(qn('w:fldCharType'), 'separate')
    r3._element.append(fldSep)
    _set_run_font(r3, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)

    # Run 4: 占位文本（域未更新时显示）
    r4 = paragraph.add_run('1')
    _set_run_font(r4, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)

    # Run 5: fldChar end
    r5 = paragraph.add_run()
    fldEnd = OxmlElement('w:fldChar')
    fldEnd.set(qn('w:fldCharType'), 'end')
    r5._element.append(fldEnd)
    _set_run_font(r5, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)


def add_header_footer(doc, header_items: List[str], header_data: Dict[str, str]):
    """设置页眉和页脚。

    header_items  : 选择显示在页眉中的信息项，如 ["title", "org"]
    header_data   : 各信息项的文本，如 {"title": "...", "org": "..."}

    效果：
      - 首页（封面）不显示页眉页脚
      - 页眉：所选信息居中显示，底部加深蓝细线
      - 页脚：居中页码 "第 X 页"
    """
    section = doc.sections[0]
    # 封面页不显示页眉页脚
    section.different_first_page_header_footer = True

    # 显式创建 first-page header/footer 空件（否则 Word 可能无法打开）
    first_header = section.first_page_header
    first_header.is_linked_to_previous = False
    for fp in first_header.paragraphs:
        fp.clear()
    first_footer = section.first_page_footer
    first_footer.is_linked_to_previous = False
    for fp in first_footer.paragraphs:
        fp.clear()

    # ---- 页眉 ----
    header = section.header
    header.is_linked_to_previous = False

    # 清空已有内容
    for p in header.paragraphs:
        p.clear()

    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)

    # 拼接页眉文本
    parts = []
    for item in header_items:
        val = header_data.get(item, "")
        if val:
            parts.append(val)
    header_text = "  |  ".join(parts)

    if header_text:
        r = p.add_run(header_text)
        _set_run_font(r, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)

    # 页眉底部边框（必须按 schema 顺序插入，否则 Word 无法打开）
    pPr = p._element.get_or_add_pPr()
    existing_bdr = pPr.find(qn('w:pBdr'))
    if existing_bdr is not None:
        pPr.remove(existing_bdr)
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')      # 0.5pt
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), COLOR_HEADER_BORDER)
    pBdr.append(bottom)
    _insert_pPr_child(pPr, pBdr)

    # ---- 页脚 ----
    footer = section.footer
    footer.is_linked_to_previous = False

    for fp in footer.paragraphs:
        fp.clear()

    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    r = fp.add_run("第 ")
    _set_run_font(r, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)

    _add_page_field(fp)

    r = fp.add_run(" 页")
    _set_run_font(r, FONT_BODY, SIZE_HEADER, color=COLOR_HEADER)


# ---------------------------------------------------------------------------
# 标题 & 正文
# ---------------------------------------------------------------------------

def add_heading(doc, level: int, text: str, fc: FontConfig = None):
    """按指定级别写入标题（统一格式 + 应用 Word 原生 Heading 样式）。

    支持 1–6 级：章 / 节 / 小节 / 四级 / 五级 / 六级。
    """
    if fc is None:
        fc = FontConfig()
    try:
        level = int(level)
    except (TypeError, ValueError):
        level = 1
    level = min(max(level, 1), 6)

    _cfg = {
        1: (fc.h1_font, fc.h1_size, fc.h1_bold, fc.h1_italic, fc.h1_underline, fc.h1_color),
        2: (fc.h2_font, fc.h2_size, fc.h2_bold, fc.h2_italic, fc.h2_underline, fc.h2_color),
        3: (fc.h3_font, fc.h3_size, fc.h3_bold, fc.h3_italic, fc.h3_underline, fc.h3_color),
        4: (fc.h4_font, fc.h4_size, fc.h4_bold, fc.h4_italic, fc.h4_underline, fc.h4_color),
        5: (fc.h5_font, fc.h5_size, fc.h5_bold, fc.h5_italic, fc.h5_underline, fc.h5_color),
        6: (fc.h6_font, fc.h6_size, fc.h6_bold, fc.h6_italic, fc.h6_underline, fc.h6_color),
    }
    font, size_pt, bold, italic, under, hex_color = _cfg[level]
    style_name = f"Heading {level}"

    color = _color_from_hex(hex_color)
    try:
        p = doc.add_paragraph(style=style_name)
    except KeyError:
        # 个别模板缺 Heading 4–6 时回退到已 ensure 的样式
        p = doc.add_paragraph(style="Heading 3" if level >= 3 else f"Heading {level}")
    p.paragraph_format.space_before = Pt(18 if level == 1 else (12 if level <= 3 else 8))
    p.paragraph_format.space_after = Pt(6 if level <= 3 else 4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    _set_run_font(r, font, Pt(size_pt), bold=bold, color=color, italic=italic, underline=under)
    return p


def add_body_paragraph(doc, block: Block, fc: FontConfig = None):
    """添加正文段落，按 FontConfig 统一字体/字号/颜色/缩进/行距。"""
    if fc is None:
        fc = FontConfig()
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = float(fc.body_line_spacing or 1.5)
    _set_first_line_indent(p, chars=fc.body_indent, body_size_pt=fc.body_size)
    if block.para_format:
        _apply_para_format(p, block.para_format)
        # 行距以格式设置为准（模板或手动）
        p.paragraph_format.line_spacing = float(fc.body_line_spacing or 1.5)

    if block.runs:
        for ri in block.runs:
            r = p.add_run(ri.text)
            _apply_run_info(r, ri, fc)
    else:
        r = p.add_run(block.text)
        _apply_run_info(r, RunInfo(text=block.text), fc)
    return p


def add_bullet(doc, block: Block, fc: FontConfig = None):
    """F06：写入列表项，保留 run 完整格式 + 段落格式 + 列表层级。

    旧版只接受 text 字符串、固定用 List Bullet 样式、丢失所有字符格式。
    现在用 block.runs（RunInfo）+ block.para_format + block.list_level。
    编号列表用 numPr 重建（num_id 在跨文档合并时无法直接迁移，改用缩进+符号表达层级）。
    """
    if fc is None:
        fc = FontConfig()
    # 用缩进表达层级（每层缩进 2 字符），避免跨文档 numId 冲突
    indent_chars = block.list_level * 2
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = float(fc.body_line_spacing or 1.5)
    if indent_chars > 0:
        p.paragraph_format.left_indent = Pt(indent_chars * 12)
    # 应用保留的段落格式（若有）
    if block.para_format:
        _apply_para_format(p, block.para_format)
    # 写入 run，保留 italic/underline/color/size
    for ri in (block.runs if block.runs else [RunInfo(text=block.text)]):
        if not ri.text:
            continue
        r = p.add_run(ri.text)
        _apply_run_info(r, ri, fc)
    return p


# ---------------------------------------------------------------------------
# 题注检测 & SEQ 域 & 交叉引用
# ---------------------------------------------------------------------------

# 题注正则：匹配 "图1：xxx" / "表 2 — xxx" 等
RE_FIG_CAPTION = re.compile(r"^图\s*(\d+)\s*[:：.\-\s]*(.*)$")
RE_TBL_CAPTION = re.compile(r"^表\s*(\d+)\s*[:：.\-\s]*(.*)$")
# 正文引用正则：匹配 "如图1所示" / "见表2" 等
RE_FIG_TBL_REF = re.compile(r"([图表])\s*(\d+)")


@dataclass
class Caption:
    """题注实体（F04：实体 ID 与显示编号分离，支持作用域查找）。

    entity_id: 全局唯一书签名（_Ref_图_<seq>），用于 REF 域引用
    caption_type: "图" 或 "表"
    orig_num: 源文件中的原始编号（各部门文件常各自从 1 开始）
    seq_num: 输出文档中的统一编号（全局递增）
    section_id: 该题注所属 section（作用域键之一）
    source_file: 该题注所属来源文件（作用域键之一）
    """
    entity_id: str
    caption_type: str
    orig_num: int
    seq_num: int
    section_id: str
    source_file: str


@dataclass
class _RefResolveResult:
    """引用解析结果。"""
    bookmark_name: Optional[str]   # 命中题注的书签名；None 表示未命中
    ambiguous: bool = False        # 多匹配（歧义）
    note: str = ""                 # 歧义/无匹配说明


class CaptionIndex:
    """题注作用域索引（F04：替代旧的全局单值 caption_map）。

    支持按 (类型, 原编号) 在多级作用域内查找题注实体：
      1. 当前 section + 当前来源文件（最精确）
      2. 当前 section（跨文件，但同章节）
      3. 当前来源文件（跨章节，但同文件）
      4. 全局（兜底，仅当唯一时才采纳）
    零匹配或多匹配时返回 ambiguous=True，调用方保留原文并报告。
    """

    def __init__(self):
        # 按 (caption_type, orig_num) 分组的题注列表
        self._by_key: Dict[Tuple[str, int], List[Caption]] = {}

    def add(self, cap: Caption):
        key = (cap.caption_type, cap.orig_num)
        self._by_key.setdefault(key, []).append(cap)

    def resolve(self, caption_type: str, orig_num: int,
                section_id: str, source_file: str) -> _RefResolveResult:
        """按作用域优先级解析引用。"""
        key = (caption_type, orig_num)
        candidates = self._by_key.get(key, [])
        if not candidates:
            return _RefResolveResult(None, note=f"无匹配题注 {caption_type}{orig_num}")

        # 作用域 1：当前 section + 当前文件
        hits = [c for c in candidates
                if c.section_id == section_id and c.source_file == source_file]
        if len(hits) == 1:
            return _RefResolveResult(hits[0].entity_id)
        if len(hits) > 1:
            return _RefResolveResult(None, ambiguous=True,
                                     note=f"{caption_type}{orig_num} 在当前章节+文件有 {len(hits)} 个题注")

        # 作用域 2：当前 section
        hits = [c for c in candidates if c.section_id == section_id]
        if len(hits) == 1:
            return _RefResolveResult(hits[0].entity_id)
        if len(hits) > 1:
            return _RefResolveResult(None, ambiguous=True,
                                     note=f"{caption_type}{orig_num} 在当前章节有 {len(hits)} 个题注")

        # 作用域 3：当前来源文件
        hits = [c for c in candidates if c.source_file == source_file]
        if len(hits) == 1:
            return _RefResolveResult(hits[0].entity_id)
        if len(hits) > 1:
            return _RefResolveResult(None, ambiguous=True,
                                     note=f"{caption_type}{orig_num} 在当前文件有 {len(hits)} 个题注")

        # 作用域 4：全局兜底（仅当全局唯一时才采纳）
        if len(candidates) == 1:
            return _RefResolveResult(candidates[0].entity_id)
        return _RefResolveResult(None, ambiguous=True,
                                 note=f"{caption_type}{orig_num} 全局有 {len(candidates)} 个题注，无法确定")


def is_fig_caption(text: str) -> bool:
    return bool(RE_FIG_CAPTION.match(text.strip()))


def is_tbl_caption(text: str) -> bool:
    return bool(RE_TBL_CAPTION.match(text.strip()))


def is_table_ref(text: str) -> bool:
    """旧接口保留：判断是否为表引用行。"""
    t = text.strip()
    return bool(re.match(r"^表\s*\d+", t)) or t.startswith("表：")


def is_fig_ref(text: str) -> bool:
    """旧接口保留：判断是否为图引用行。"""
    t = text.strip()
    return bool(re.match(r"^图\s*\d+", t)) or t.startswith("图：")


def add_caption(doc, block: Block, caption_type: str, seq_num: int):
    """添加居中的图表题注，使用 SEQ 域自动编号 + 书签。

    结构：[bookmarkStart] "图" [SEQ域] "：" [bookmarkEnd] "描述文字"
    显示效果：图1：描述…（全角冒号；编号与标签同为题注字号 SIZE_CAPTION）

    书签包裹 SEQ 域结果，使正文 REF 域可引用题注编号。
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.0

    # 提取描述文字
    if caption_type == "图":
        m = RE_FIG_CAPTION.match(block.text.strip())
    else:
        m = RE_TBL_CAPTION.match(block.text.strip())
    description = m.group(2).strip() if m and m.group(2) else ""
    # 描述若仍以冒号开头（源文残留），去掉以免「图1：：xxx」
    if description.startswith(("：", ":", "—", "-", ".")):
        description = description.lstrip("：:.—-–—．. ").strip()

    seq_name = "Figure" if caption_type == "图" else "Table"
    bookmark_name = f"_Ref_{caption_type}_{seq_num}"

    # 1. 题注标签（"图" / "表"，与编号紧挨）
    r = p.add_run(caption_type)
    _set_run_font(r, FONT_BODY, SIZE_CAPTION)

    # 2. 书签开始
    bm_id = _add_bookmark_start(p, bookmark_name)

    # 3. SEQ 域（自动编号，字号与题注一致）
    _add_field_simple(
        p,
        f"SEQ {seq_name} \\* ARABIC",
        placeholder=str(seq_num),
        size=SIZE_CAPTION,
        font_name=FONT_BODY,
    )

    # 4. 书签结束
    _add_bookmark_end(p, bm_id)

    # 5. 全角冒号 + 描述
    r = p.add_run("：")
    _set_run_font(r, FONT_BODY, SIZE_CAPTION)
    if description:
        r = p.add_run(description)
        _set_run_font(r, FONT_BODY, SIZE_CAPTION)

    return p


def add_body_with_xref(doc, block: Block, caption_index: CaptionIndex,
                       fc: FontConfig = None, resolve_log: Optional[List[str]] = None):
    """添加正文段落：保留完整 run 格式 + 段落格式 + 将图N/表N引用替换为 REF 域。

    F04：引用解析改用 CaptionIndex 作用域查找，取代旧的全局单值 caption_map。
    F06：runs 从 (text, is_bold) 升级为 RunInfo，保留 italic/underline/color/size；
         应用 block.para_format（对齐/缩进/行距/段前段后）。
    block.section_id / block.source_file 决定当前引用的作用域。
    resolve_log：若提供，歧义/无匹配会追加到该列表，供 aggregate 审计。
    """
    if fc is None:
        fc = FontConfig()
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = float(fc.body_line_spacing or 1.5)
    _set_first_line_indent(p, chars=fc.body_indent, body_size_pt=fc.body_size)
    # F06：应用保留的段落格式（覆盖默认），但行距以格式设置为准
    if block.para_format:
        _apply_para_format(p, block.para_format)
        p.paragraph_format.line_spacing = float(fc.body_line_spacing or 1.5)

    for ri in (block.runs if block.runs else [RunInfo(text=block.text)]):
        _add_run_with_xref(p, ri, caption_index,
                           block.section_id, block.source_file, fc, resolve_log)

    return p


def _add_run_with_xref(paragraph, ri: RunInfo,
                       caption_index: CaptionIndex, section_id: str, source_file: str,
                       fc: FontConfig = None, resolve_log: Optional[List[str]] = None):
    """将 RunInfo 文本写入段落，图N/表N部分按作用域解析并替换为 REF 域。

    F06：用 RunInfo 保留的 italic/underline/color/size 写入，而非只保 bold。
    """
    if fc is None:
        fc = FontConfig()
    body_size = Pt(fc.body_size)
    text = ri.text
    last_end = 0

    for m in RE_FIG_TBL_REF.finditer(text):
        # 前面的普通文本
        if m.start() > last_end:
            prefix = text[last_end:m.start()]
            r = paragraph.add_run(prefix)
            _apply_run_info(r, ri, fc)

        label = m.group(1)   # "图" or "表"
        num = int(m.group(2))

        result = caption_index.resolve(label, num, section_id, source_file)
        if result.bookmark_name and not result.ambiguous:
            # 命中唯一题注 → 用 REF 域
            r = paragraph.add_run(label)
            _apply_run_info(r, ri, fc)
            _add_field_simple(
                paragraph,
                f'REF {result.bookmark_name} \\h',
                placeholder=str(num),
                size=body_size,
                bold=ri.bold if ri.bold is not None else fc.body_bold,
            )
        else:
            # 无匹配或歧义 → 保留原文，记录供审计
            r = paragraph.add_run(m.group(0))
            _apply_run_info(r, ri, fc)
            if resolve_log is not None:
                if result.ambiguous:
                    resolve_log.append(
                        f"⚠️ 引用歧义：'{label}{num}'（当前 {source_file}/{section_id}）—— {result.note}，已保留原文")
                elif result.bookmark_name is None:
                    resolve_log.append(
                        f"ℹ️ 无匹配题注：'{label}{num}'（当前 {source_file}/{section_id}）—— {result.note}")

        last_end = m.end()

    # 剩余文本
    if last_end < len(text):
        suffix = text[last_end:]
        r = paragraph.add_run(suffix)
        _apply_run_info(r, ri, fc)


# ---------------------------------------------------------------------------
# 表格 & 图片原位插入
# ---------------------------------------------------------------------------

def _set_cell_shading(cell, fill_color: str):
    """设置单元格背景色（w:shd）。"""
    tc_pr = cell._element.get_or_add_tcPr()
    # 移除已有的 shd
    existing = tc_pr.find(qn('w:shd'))
    if existing is not None:
        tc_pr.remove(existing)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tc_pr.append(shd)


def _set_header_row_repeat(row):
    """设置首行在每页重复显示（w:tblHeader）。"""
    trPr = row._element.get_or_add_trPr()
    existing = trPr.find(qn('w:tblHeader'))
    if existing is not None:
        trPr.remove(existing)
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), 'true')
    trPr.append(tblHeader)


def _style_table(table: Table):
    """应用表格样式：首行突出（深蓝底白字加粗）+ 隔行换色（白+浅蓝）。"""
    # 确保 Table Grid 样式（带边框）
    try:
        table.style = "Table Grid"
    except KeyError:
        pass

    for i, row in enumerate(table.rows):
        for cell in row.cells:
            # 背景色
            if i == 0:
                _set_cell_shading(cell, TABLE_HEADER_FILL)
            elif i % 2 == 0:
                _set_cell_shading(cell, TABLE_ALT_FILL)
            else:
                _set_cell_shading(cell, "auto")

            # 字体
            for para in cell.paragraphs:
                for run in para.runs:
                    if i == 0:
                        _set_run_font(run, FONT_BODY, Pt(10), bold=True, color=COLOR_WHITE)
                    else:
                        _set_run_font(run, FONT_BODY, Pt(10))
                # 处理没有 run 但有文本的情况
                if not para.runs and para.text:
                    r = para.add_run(para.text)
                    if i == 0:
                        _set_run_font(r, FONT_BODY, Pt(10), bold=True, color=COLOR_WHITE)
                    else:
                        _set_run_font(r, FONT_BODY, Pt(10))

    # 首行在每页重复
    if len(table.rows) > 0:
        _set_header_row_repeat(table.rows[0])


def _copy_table_to_doc(src_table: Table, dest_doc) -> Table:
    """深拷贝源表格，插入到 dest_doc，并应用表格样式。

    F03 修复：deepcopy 后调用 _migrate_relationships 迁移表格内可能引用的
    关系（超链接、嵌图等），避免关系断裂。
    """
    src_xml = src_table._element
    new_xml = deepcopy(src_xml)
    # 表格内可能含超链接(w:hyperlink/@r:id)、嵌图(a:blip/@r:embed)等关系引用
    _migrate_relationships(new_xml, src_table.part, dest_doc)
    _insert_before_sectpr(dest_doc, new_xml)
    new_tbl = Table(new_xml, dest_doc)

    # 应用表格样式：首行突出 + 隔行换色
    _style_table(new_tbl)

    return new_tbl


# ---------------------------------------------------------------------------
# 通用 OOXML 关系迁移（F03 修复）
# ---------------------------------------------------------------------------

# 需要迁移的关系类型 → 内部/外链
# 内部关系：复制 target_part（blob + content_type）到目标包，重建 rId
# 外链关系：按 reltype + target_ref（URL）重建，重建 rId
_REL_INTERNAL = {
    RT.IMAGE, RT.HYPERLINK,  # HYPERLINK 也可能是内部的（锚点），但通常外链，下面按 is_external 判
    RT.CHART, RT.OLE_OBJECT, RT.FOOTNOTES, RT.ENDNOTES,
}
# 副本 XML 中承载关系引用的属性名（namespace-qualified）
_REL_ATTRS = [qn('r:id'), qn('r:embed'), qn('r:link')]


def _migrate_relationships(elem, src_part, dest_doc) -> List[str]:
    """把 elem（已 deepcopy 的 XML 子树）中所有关系引用迁移到 dest_doc。

    src_part: 源内容所在的 Part（如 src_doc.part 或 table.part），用于查 .rels。
    dest_doc: 目标 Document。

    遍历 elem 中所有 r:id / r:embed / r:link 属性，对每个旧 rId：
      - 内部关系（is_external=False）：复制 target_part 的 blob/content_type 到
        目标包（用统一 partname 分配避免冲突），建立新关系，更新属性为新 rId
      - 外链关系（is_external=True，如外链超链接/外链图）：按 reltype + target_ref
        重建外链关系，更新属性为新 rId
    返回迁移日志列表（供 aggregate 审计）。
    """
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    dest_part = dest_doc.part
    dest_pkg = dest_part.package
    migrated: List[str] = []

    # 统一 partname 分配：扫描目标包已有 partname，按类型递增编号避免冲突
    if not hasattr(dest_doc, '_migrated_part_counter'):
        dest_doc._migrated_part_counter = {}
    counter = dest_doc._migrated_part_counter

    # 收集 elem 中所有带关系属性的节点
    rel_refs = []  # (node, attr_qname, old_rid)
    for attr_qn in _REL_ATTRS:
        for node in elem.iter():
            if node.get(attr_qn) is not None:
                rel_refs.append((node, attr_qn, node.get(attr_qn)))

    for node, attr_qn, old_rid in rel_refs:
        src_rel = src_part.rels.get(old_rid)
        if src_rel is None:
            migrated.append(f"⚠️ 未找到源关系 rId={old_rid}（{attr_qn}），引用可能断裂")
            continue

        reltype = src_rel.reltype

        if src_rel.is_external:
            # 外链关系（外链超链接、外链图片 r:link 等）：按 URL 重建
            target_ref = src_rel.target_ref  # URL 字符串
            new_rid = dest_part.relate_to(target_ref, reltype, is_external=True)
            node.set(attr_qn, new_rid)
            migrated.append(f"外链关系 {reltype.split('/')[-1]}: {old_rid}→{new_rid} ({target_ref})")
        else:
            # 内部关系：复制 target_part
            src_target = src_rel.target_part
            blob = src_target.blob
            content_type = src_target.content_type

            # 推断 partname 扩展名
            ext = _ext_for_content_type(content_type, src_target.partname)

            # 按 reltype 选 partname 前缀，统一计数避免冲突
            kind = reltype.split('/')[-1]  # image / hyperlink / chart / ...
            counter[kind] = counter.get(kind, 0) + 1
            idx = counter[kind]
            partname = PackURI('/word/media/%s%d.%s' % (kind, idx, ext))

            new_part = Part(partname, content_type, blob, dest_pkg)
            new_rid = dest_part.relate_to(new_part, reltype)
            node.set(attr_qn, new_rid)
            migrated.append(f"内部关系 {kind}: {old_rid}→{new_rid} ({content_type})")

    return migrated


def _ext_for_content_type(content_type: str, src_partname) -> str:
    """根据 content_type 推断扩展名，回退到源 partname 的扩展名或 bin。"""
    ext_map = {
        'image/png': 'png',
        'image/jpeg': 'jpg',
        'image/gif': 'gif',
        'image/bmp': 'bmp',
        'image/tiff': 'tiff',
        'image/svg+xml': 'svg',
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml': 'xml',
        'application/vnd.openxmlformats-officedocument.oleObject': 'bin',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml': 'xml',
    }
    if content_type in ext_map:
        return ext_map[content_type]
    # 回退：从源 partname 取扩展名
    if src_partname and hasattr(src_partname, 'ext'):
        return src_partname.ext
    return 'bin'


def _copy_image_paragraph_to_doc(src_para_elem, src_doc: _Document, dest_doc: _Document):
    """深拷贝含图片的段落 XML，并把其中所有 OOXML 关系迁移到目标文档。

    F03 修复：原版只处理 a:blip 的 r:embed（嵌入图片），漏掉超链接、外链图片
    (r:link)、图表、OLE 对象等。现在改用通用关系迁移器 _migrate_relationships
    统一处理副本中所有 r:id / r:embed / r:link 属性。
    """
    new_para = deepcopy(src_para_elem)
    _migrate_relationships(new_para, src_doc.part, dest_doc)
    _insert_before_sectpr(dest_doc, new_para)
    return new_para


def _verify_docx_integrity(path: str) -> Tuple[bool, str]:
    """F03：校验生成文件的 OOXML 包完整性。

    重新打开 docx，遍历主文档 part 的所有关系，确认每个内部关系的目标 part
    可达；外链关系跳过（无法离线验证 URL）。返回 (ok, message)。
    """
    try:
        from docx import Document as _OpenDoc
        chk = _OpenDoc(path)
        main_part = chk.part
        broken = []
        n_internal = 0
        n_external = 0
        for rid, rel in main_part.rels.items():
            if rel.is_external:
                n_external += 1
                continue
            n_internal += 1
            try:
                _ = rel.target_part  # 触发解析，目标不存在会抛
            except Exception as e:
                broken.append(f"{rid}({rel.reltype.split('/')[-1]}): {e}")
        # 同时确认 ZIP 部件数 > 0（防止空包）
        try:
            parts = list(main_part.package.iter_parts())
        except Exception:
            parts = []
        if broken:
            return False, f"{len(broken)} 个关系断裂：{'; '.join(broken[:3])}"
        if not parts:
            return False, "包内无任何 part"
        return True, f"内部关系 {n_internal} 个、外链 {n_external} 个、parts {len(parts)} 个，全部有效"
    except Exception as e:
        return False, f"无法���新打开文件：{e}"


# ---------------------------------------------------------------------------
# 章节排序（预览 / 聚合共用）
# ---------------------------------------------------------------------------

def sort_section_ids(parsed, template: Optional[List[HeadingSpec]] = None) -> List[str]:
    """按最终报告规则返回 section_id 顺序。"""
    template_order = {
        spec.section_id: i for i, spec in enumerate(template or [])
    }

    def sort_key(sid: str):
        if sid == "_preamble":
            return (-1, 0)
        if sid in template_order:
            return (0, template_order[sid])
        try:
            return (1, tuple(int(x) for x in sid.split(".")))
        except ValueError:
            return (2, (0,))

    return sorted(parsed.keys(), key=sort_key)


def _update_word_fields(doc) -> None:
    try:
        doc.Fields.Update()
    except Exception:
        pass
    try:
        for story in doc.StoryRanges:
            try:
                story.Fields.Update()
            except Exception:
                pass
    except Exception:
        pass
    try:
        for toc in doc.TablesOfContents:
            toc.Update()
    except Exception:
        pass


def _convert_docx_to_pdf_win_word(docx_path: str, pdf_path: str):
    """Windows：用本机 Word (COM) 更新域并导出 PDF。"""
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False, "未装 pywin32，无法用 Word 生成 PDF"
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(docx_path))
        _update_word_fields(doc)
        doc.Save()
        doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)
        doc.Close(0)
        doc = None
        if not os.path.isfile(pdf_path) or os.path.getsize(pdf_path) == 0:
            return False, "Word 未生成有效 PDF 文件"
        return True, "PDF 已生成（Word COM：目录/题注/页码已更新）"
    except Exception as e:
        return False, f"Word COM 导出失败：{e}"
    finally:
        try:
            if doc is not None:
                doc.Close(0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def _convert_docx_to_pdf_mac_word(docx_path: str, pdf_path: str):
    """macOS：通过 AppleScript 调用 Microsoft Word 导出 PDF。"""
    docx_abs = os.path.abspath(docx_path)
    pdf_abs = os.path.abspath(pdf_path)
    # 转义 AppleScript 字符串中的反斜杠与引号
    def _as_escape(s: str) -> str:
        return s.replace("\\", "\\\\").replace('"', '\\"')

    docx_esc = _as_escape(docx_abs)
    pdf_esc = _as_escape(pdf_abs)
    script = f'''
set docxPath to POSIX file "{docx_esc}"
set pdfPath to POSIX file "{pdf_esc}"
tell application "Microsoft Word"
    activate
    set theDoc to open file name docxPath
    try
        update fields of theDoc
    end try
    try
        save as theDoc file name pdfPath file format format PDF
    on error errMsg number errNum
        close theDoc saving no
        error errMsg number errNum
    end try
    close theDoc saving no
end tell
'''
    try:
        # 先确认 Word 是否存在
        check = subprocess.run(
            ["osascript", "-e", 'id of application "Microsoft Word"'],
            capture_output=True, text=True, timeout=15,
        )
        if check.returncode != 0:
            return False, "未安装 Microsoft Word（Mac）"
        r = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=180,
        )
        if r.returncode != 0:
            err = (r.stderr or r.stdout or "").strip() or f"exit {r.returncode}"
            return False, f"Word AppleScript 失败：{err}"
        if not os.path.isfile(pdf_abs) or os.path.getsize(pdf_abs) == 0:
            return False, "Word 未写出有效 PDF"
        return True, "PDF 已生成（macOS Word）"
    except FileNotFoundError:
        return False, "系统无 osascript，无法调用 Word"
    except subprocess.TimeoutExpired:
        return False, "Word 导出 PDF 超时"
    except Exception as e:
        return False, f"macOS Word 导出失败：{e}"


def _convert_docx_to_pdf_libreoffice(docx_path: str, pdf_path: str):
    """跨平台：LibreOffice / soffice 无界面导出 PDF（目录域可能不会像 Word 一样完整更新）。"""
    candidates = [
        "soffice",
        "libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    bin_path = None
    for c in candidates:
        if os.path.isabs(c) and os.path.isfile(c):
            bin_path = c
            break
        found = shutil.which(c)
        if found:
            bin_path = found
            break
    if not bin_path:
        return False, "未找到 LibreOffice (soffice)"

    docx_abs = os.path.abspath(docx_path)
    pdf_abs = os.path.abspath(pdf_path)
    out_dir = tempfile.mkdtemp(prefix="report-pdf-")
    try:
        r = subprocess.run(
            [bin_path, "--headless", "--nologo", "--nofirststartwizard",
             "--convert-to", "pdf", "--outdir", out_dir, docx_abs],
            capture_output=True, text=True, timeout=180,
        )
        if r.returncode != 0:
            err = (r.stderr or r.stdout or "").strip() or f"exit {r.returncode}"
            return False, f"LibreOffice 导出失败：{err}"
        # 输出名通常与源文件 stem 相同
        stem = os.path.splitext(os.path.basename(docx_abs))[0] + ".pdf"
        produced = os.path.join(out_dir, stem)
        if not os.path.isfile(produced):
            # 兜底：取目录内第一个 pdf
            pdfs = [f for f in os.listdir(out_dir) if f.lower().endswith(".pdf")]
            if not pdfs:
                return False, "LibreOffice 未生成 PDF"
            produced = os.path.join(out_dir, pdfs[0])
        shutil.copy2(produced, pdf_abs)
        if not os.path.isfile(pdf_abs) or os.path.getsize(pdf_abs) == 0:
            return False, "LibreOffice PDF 无效"
        return True, "PDF 已生成（LibreOffice；目录域请在 Word 中再更新一次更佳）"
    except subprocess.TimeoutExpired:
        return False, "LibreOffice 导出超时"
    except Exception as e:
        return False, f"LibreOffice 导出失败：{e}"
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)


def _convert_docx_to_pdf_with_word(docx_path: str, pdf_path: str):
    """跨平台导出 PDF：优先本机 Word，其次 LibreOffice。

    - Windows：Word COM（pywin32）
    - macOS：Microsoft Word + AppleScript
    - 通用回退：LibreOffice soffice
    失败时仍保留 DOCX，不中断主流程。
    """
    errors: list[str] = []
    if sys.platform.startswith("win"):
        ok, msg = _convert_docx_to_pdf_win_word(docx_path, pdf_path)
        if ok:
            return True, msg
        errors.append(msg)
    elif sys.platform == "darwin":
        ok, msg = _convert_docx_to_pdf_mac_word(docx_path, pdf_path)
        if ok:
            return True, msg
        errors.append(msg)

    ok, msg = _convert_docx_to_pdf_libreoffice(docx_path, pdf_path)
    if ok:
        return True, msg
    errors.append(msg)

    detail = "；".join(errors)
    return False, f"PDF 未生成：{detail}。DOCX 仍可正常使用，可在 Word/WPS 中打开后另存为 PDF"


# ---------------------------------------------------------------------------
# 主流程：组装报告
# ---------------------------------------------------------------------------

def aggregate(
    input_dir: str,
    output_path: str,
    *,
    period: str = "2026年第二季度",
    title: str = "中国保险行业调研报告",
    subtitle: str = "市场环境 · 保费增长 · 渠道变革 · 趋势展望",
    org: str = "保险行业研究中心",
    date: str = "2026年7月",
    header_items: Optional[List[str]] = None,
    template: Optional[List[HeadingSpec]] = None,
    font_config: Optional[FontConfig] = None,
    silent: bool = False,
    allow_skip_failures: bool = False,
    title_overrides: Optional[Dict[str, str]] = None,
    section_overrides: Optional[Dict[str, Dict]] = None,
    section_order: Optional[List[str]] = None,
    strict_section_order: bool = False,
    exclude_section_ids: Optional[List[str]] = None,
    format_template: Optional[FormatTemplate] = None,
    format_template_path: Optional[str] = None,
    use_template_titles: Optional[bool] = None,
) -> Dict:
    """读取 input_dir 下的所有 docx，按 section_id 排序汇总，输出到 output_path。

    默认动态模式：章节结构由输入决定。
    可选 format_template / format_template_path：以某次「合成终版 DOCX」为
    格式参考（章节顺序、标题、层级、字体），内容仍来自本次输入材料。

    典型场景：用 Q2 终版作为模板，合成 Q3 输入 → 结构/版式对齐 Q2，内容用 Q3。
    """
    header_items = header_items if header_items is not None else ["title"]

    # 加载格式模板
    ft: Optional[FormatTemplate] = format_template
    if ft is None and format_template_path:
        ft = load_format_template(format_template_path)
    if ft is not None:
        if template is None and ft.specs:
            template = list(ft.specs)
        # 启用格式模板时：版式（字体/字号/加粗/颜色）严格取自模板；
        # 仅当调用方未传入 font_config 时用模板；若传入则以其为准（桌面端可保留 body_indent）。
        if font_config is None and ft.font_config is not None:
            font_config = ft.font_config
        if section_order is None and ft.specs:
            section_order = ft.section_order()
        prefer_tpl_titles = (
            ft.use_template_titles if use_template_titles is None else use_template_titles
        )
    else:
        prefer_tpl_titles = False

    fc = font_config or FontConfig()
    _apply_font_config(fc)
    log = []

    def log_msg(s: str):
        log.append(s)
        if not silent:
            print(s)

    _reset_bookmark_counter()

    if ft is not None:
        log_msg(f"[0/5] 格式模板：{ft.source_path or '(内存)'}")
        log_msg(
            f"      模板章节 {len(ft.specs)} 个；版式严格按模板 "
            f"H1={fc.h1_font}/{fc.h1_size}/#{fc.h1_color} "
            f"H2={fc.h2_font}/{fc.h2_size}/#{fc.h2_color} "
            f"H3={fc.h3_font}/{fc.h3_size}/#{fc.h3_color} "
            f"Body={fc.body_font}/{fc.body_size}/#{fc.body_color}"
        )

    log_msg(f"[1/5] 扫描输入目录：{input_dir}")
    parsed = parse_input_dir(input_dir)
    log_msg(f"      共识别 section：{len(parsed)} 个"
            f"（成功文件 {len(parsed.parsed_files)} 份）")

    # ---- F10：失败文件处理 ----
    if parsed.failures:
        log_msg(f"      ⚠️ 发现 {len(parsed.failures)} 份失败文件：")
        for fl in parsed.failures:
            size_str = f"，{fl.size} 字节" if fl.size is not None else ""
            log_msg(f"        ✗ {fl.file_name}  [{fl.category}]{size_str}")
            log_msg(f"            {fl.message}")
        if not allow_skip_failures:
            # 默认禁止生成：列出失败文件供用户决策
            failure_lines = [
                f"{fl.file_name}({fl.category}): {fl.message}"
                for fl in parsed.failures
            ]
            raise ValueError(
                f"检测到 {len(parsed.failures)} 份无法解析的文件，已中止生成。"
                f"请修正或移除这些文件后重试；若确认跳过失败文件继续生成，"
                f"需显式设置 allow_skip_failures=True。\n失败文件：\n"
                + "\n".join(failure_lines)
            )
        else:
            log_msg(f"      ℹ️ 已设置 allow_skip_failures=True，跳过上述失败文件继续生成（报告中将留痕）")

    # 无有效 section 时一律阻止生成：即便允许跳过失败文件，其前提也是
    # "跳过坏文件后仍有有效文件可生成"。全部文件都失败时若继续生成，只会
    # 得到一个只有封面/目录的空壳报告，对用户产生误导，违背 F10 让用户
    # 知情决策的精神。
    if not parsed.sections:
        raise ValueError("输入目录中没有可解析的 docx 文件，无法生成报告")

    # ---- 格式模板：用模板标题/层级对齐输入中已有的 section_id ----
    rename_map: Dict[str, str] = {}
    if ft is not None and prefer_tpl_titles and ft.specs:
        n_title = n_level = 0
        for spec in ft.specs:
            sec = parsed.get(spec.section_id)
            if sec is None:
                continue
            # 避免把误识别的超长正文标题强行改成模板章节名
            input_title = (sec.title or "").strip()
            tpl_title = (spec.title or "").strip()
            suspicious = (
                getattr(sec, "title_conflict", False)
                or getattr(sec, "confidence", "") == "low"
                or len(input_title) > 36
                or (tpl_title and len(tpl_title) > 36)
            )
            if (
                not suspicious
                and tpl_title
                and input_title != tpl_title
            ):
                log_msg(f"      模板标题：{spec.section_id}  {input_title!r} → {tpl_title!r}")
                sec.title = tpl_title
                n_title += 1
            # 层级：仅在输入层级与编号深度大致吻合时才用模板层级校正
            if spec.level and getattr(sec, "level", 0) != spec.level:
                expected = spec.section_id.count(".") + 1
                if spec.level == expected:
                    sec.level = spec.level
                    n_level += 1
        if n_title or n_level:
            log_msg(f"      模板对齐：更新标题 {n_title} 处、层级 {n_level} 处")

    # ---- 用户手动覆盖（预览编辑），可覆盖模板标题；支持 create 空标题壳 ----
    if section_overrides or title_overrides:
        unified: Dict[str, Dict] = {}
        if title_overrides:
            for sid, nt in title_overrides.items():
                tt = str(nt or "").strip()
                if tt:
                    unified.setdefault(sid, {})["title"] = tt
        if section_overrides:
            for sid, ov in section_overrides.items():
                if not isinstance(ov, dict):
                    continue
                bucket = unified.setdefault(sid, {})
                if ov.get("title") is not None and str(ov.get("title")).strip():
                    bucket["title"] = str(ov["title"]).strip()
                if ov.get("section_id") is not None and str(ov.get("section_id")).strip():
                    bucket["section_id"] = str(ov["section_id"]).strip()
                if ov.get("level") is not None:
                    try:
                        lv = int(ov["level"])
                        if 1 <= lv <= 6:
                            bucket["level"] = lv
                    except (TypeError, ValueError):
                        pass
                if ov.get("create") or ov.get("empty"):
                    bucket["create"] = True
        renames = []
        created_n = 0
        for old_sid, ov in unified.items():
            sec = parsed.get(old_sid)
            # 预览中「新增子标题」：无正文，仅写入标题
            if sec is None and ov.get("create"):
                target_sid = str(ov.get("section_id") or old_sid).strip()
                if not target_sid or target_sid.startswith("_"):
                    log_msg(f"      ⚠️ 跳过无效新增章节编号：{target_sid!r}")
                    continue
                if target_sid in parsed:
                    log_msg(f"      ⚠️ 新增章节编号冲突：{target_sid} 已存在，跳过")
                    continue
                # 注意：不可用 title= 赋值，会遮蔽 aggregate 参数「报告主标题」，
                # 导致封面主标题被最后一次新增章节标题覆盖。
                new_title = str(ov.get("title") or "未命名").strip() or "未命名"
                level = int(ov.get("level") or (target_sid.count(".") + 1))
                if not (1 <= level <= 6):
                    level = min(6, max(1, target_sid.count(".") + 1))
                new_sec = Section(
                    section_id=target_sid,
                    title=new_title,
                    blocks=[],
                    source_files=["(手动新增)"],
                    merge_count=1,
                    level=level,
                    confidence="high",
                    source_signal="manual",
                )
                if hasattr(parsed, "sections"):
                    parsed.sections[target_sid] = new_sec
                else:
                    parsed[target_sid] = new_sec
                if old_sid != target_sid:
                    rename_map[old_sid] = target_sid
                created_n += 1
                log_msg(f"      新增空标题：{target_sid}  {new_title!r}（level={level}，无正文）")
                continue
            if sec is None:
                continue
            if "title" in ov and sec.title != ov["title"]:
                log_msg(f"      标题覆盖：{old_sid}  {sec.title!r} → {ov['title']!r}")
                sec.title = ov["title"]
            if "level" in ov and getattr(sec, "level", None) != ov["level"]:
                log_msg(f"      层级覆盖：{old_sid}  level {getattr(sec, 'level', '?')} → {ov['level']}")
                sec.level = ov["level"]
            new_sid = ov.get("section_id")
            if new_sid and new_sid != old_sid:
                renames.append((old_sid, new_sid))
        if created_n:
            log_msg(f"      手动新增空标题 {created_n} 个")
        for old_sid, new_sid in renames:
            if new_sid in parsed and new_sid != old_sid:
                log_msg(f"      ⚠️ 编号覆盖冲突：{old_sid} → {new_sid} 已存在，跳过重命名")
                continue
            sec = parsed.get(old_sid)
            if sec is None:
                continue
            log_msg(f"      编号覆盖：{old_sid} → {new_sid}")
            sec.section_id = new_sid
            rename_map[old_sid] = new_sid
            if hasattr(parsed, "sections"):
                parsed.sections[new_sid] = sec
                parsed.sections.pop(old_sid, None)
            else:
                parsed[new_sid] = sec
                parsed.pop(old_sid, None)

    # ---- 排除章节（预览中删除/隐藏）----
    _exclude: set = set()
    if exclude_section_ids:
        for sid in exclude_section_ids:
            s = str(sid or "").strip()
            if not s:
                continue
            _exclude.add(s)
            # 也排除重命名后的目标/源
            _exclude.add(rename_map.get(s, s))
        # 反向：rename 源若映射到排除目标
        for old, new in list(rename_map.items()):
            if new in _exclude or old in _exclude:
                _exclude.add(old)
                _exclude.add(new)

    skipped_blocks = 0
    if _exclude:
        for sid in list(_exclude):
            sec = parsed.get(sid)
            if sec is None:
                continue
            nblk = len(sec.blocks)
            skipped_blocks += nblk
            log_msg(f"      排除章节：{sid}  {sec.title!r}（{nblk} 块，不写入终版）")
            if hasattr(parsed, "sections"):
                parsed.sections.pop(sid, None)
            else:
                parsed.pop(sid, None)
        if skipped_blocks:
            log_msg(f"      已排除 {len([s for s in _exclude if s not in parsed])} 个章节，"
                    f"合计跳过内容块 {skipped_blocks} 个")

    # ---- 手动/模板章节顺序 ----
    _user_section_order: Optional[List[str]] = None
    if section_order:
        mapped: List[str] = []
        seen_o: set = set()
        for sid in section_order:
            final = rename_map.get(sid, sid)
            if final in _exclude:
                continue
            if final in parsed and final not in seen_o:
                mapped.append(final)
                seen_o.add(final)
        # 非严格模式：未出现在顺序中的输入章节追加到末尾（桌面预览用严格模式，避免已删章节回流）
        if not strict_section_order:
            for sid in sort_section_ids(parsed, template):
                if sid not in seen_o and sid not in _exclude:
                    mapped.append(sid)
                    seen_o.add(sid)
        _user_section_order = mapped
        log_msg(
            f"      使用章节顺序：{len(mapped)} 项"
            f"{'（严格按预览表，不自动补缺）' if strict_section_order else ''}"
            f"{'（来自格式模板/手动调序）' if ft or section_order else ''}"
        )

    # 合并审计：输入 section 总数 = 各 Section.merge_count 之和；
    # 若大于去重后数量，说明发生了同编号合并，必须显式记录，杜绝静默丢数据。
    total_input_sections = sum(sec.merge_count for sec in parsed.values())
    merged_secs = [sec for sec in parsed.values() if sec.merge_count > 1]
    conflict_secs = [sec for sec in parsed.values() if sec.title_conflict]
    if total_input_sections != len(parsed):
        log_msg(f"      合并：输入 {total_input_sections} 个 section → 去重后 {len(parsed)} 个")
        if merged_secs:
            log_msg(f"      发生合并的 section：{len(merged_secs)} 个")
            for sec in merged_secs:
                flag = "  ⚠️标题冲突" if sec.title_conflict else ""
                log_msg(f"        {sec.section_id}  来自 {len(sec.source_files)} 份文件 → {len(sec.blocks)} 块{flag}")
                for sf in sec.source_files:
                    log_msg(f"            来源：{sf}")
                if sec.title_conflict:
                    for t in sec.titles_seen:
                        log_msg(f"            标题：{t}")
    if conflict_secs:
        log_msg(f"      ⚠️ 发现 {len(conflict_secs)} 个同编号不同标题的冲突 section，已按追加方式合并，请人工核对标题")
    total_input_blocks = sum(len(sec.blocks) for sec in parsed.values())

    # F05：preamble（首标题前的前导内容）单独取出，不参与正常章节排序
    preamble_sec: Optional[Section] = parsed.pop("_preamble", None)
    if preamble_sec and preamble_sec.blocks:
        log_msg(f"      前导内容（前言）：{len(preamble_sec.blocks)} 块，来自 {preamble_sec.source_files}，"
                f"将作为'前言'置于正文最前")

    # F12：template 参与排序——template 列出的 section 按模板顺序排前，未列入
    # 的按 section_id 数字分量排后；template 为 None 时纯数字分量排序。
    if template:
        _tpl_order = {spec.section_id: i for i, spec in enumerate(template)}
        _tpl_map = {spec.section_id: spec for spec in template}
    else:
        _tpl_order = {}
        _tpl_map = {}

    def _section_sort_key(sid: str):
        if sid == "_preamble":
            return (-1, 0)
        if sid in _tpl_order:
            return (0, _tpl_order[sid])                      # 模板内：按模板顺序
        try:
            return (1, [int(x) for x in sid.split('.')])     # 模板外：数字分量排后
        except ValueError:
            return (2, [0])                                  # 非数字：排最后

    if _user_section_order is not None:
        sorted_sids = [s for s in _user_section_order if s in parsed]
        if not strict_section_order:
            for s in sorted(parsed.keys(), key=_section_sort_key):
                if s not in sorted_sids and s not in _exclude:
                    sorted_sids.append(s)
    else:
        sorted_sids = [s for s in sorted(parsed.keys(), key=_section_sort_key) if s not in _exclude]

    # F12：缺失章节告警——模板期望但输入中没有的章节，按模板顺序列出供人工补齐
    # （仅告警，不强制阻断，与动态模式哲学兼容）
    if template:
        _missing_tpl = [sp.section_id for sp in template if sp.section_id not in parsed]
        if _missing_tpl:
            log_msg(f"      ℹ️ 模板参考：{len(_missing_tpl)} 个模板章节在输入中缺失（补齐可获完整结构）：")
            for sid in _missing_tpl:
                log_msg(f"        {sid}  {_tpl_map[sid].title}")

    low_conf_secs: List[Section] = []
    for sid in sorted_sids:
        sec = parsed[sid]
        flag = ""
        if sec.confidence == "low":
            flag = "  ⚠️低置信度(仅样式,需人工确认)"
            low_conf_secs.append(sec)
        elif sec.confidence == "high":
            flag = f"  [{sec.source_signal}]"
        elif sec.confidence == "medium":
            flag = f"  [{sec.source_signal}]"
        log_msg(f"        {sid}  {sec.title}{flag}")
    if low_conf_secs:
        log_msg(f"      ⚠️ 发现 {len(low_conf_secs)} 个仅靠样式识别的低置信度标题，"
                f"其内容已并入相邻章节，请人工确认是否应独立成章")

    # 统计图片和表格
    image_count = sum(
        1 for sec in parsed.values() for blk in sec.blocks if blk.kind == "image"
    )
    table_count = sum(
        1 for sec in parsed.values() for blk in sec.blocks if blk.kind == "table"
    )
    log_msg(f"      包含图片：{image_count} 张，表格：{table_count} 个")

    bold_count = sum(
        1 for sec in parsed.values() for blk in sec.blocks
        if blk.kind == "p" and any(ri.bold for ri in blk.runs)
    )
    log_msg(f"      含加粗文字的段落：{bold_count} 个")

    # -----------------------------------------------------------------------
    # 预扫描：建立题注实体索引（F04：作用域映射，取代旧全局单值 caption_map）
    #   每个题注分配唯一实体 ID（_Ref_图_<seq>），按 (section_id, source_file,
    #   类型, 原编号) 建索引，引用解析时按多级作用域查找，杜绝跨文件串线。
    # -----------------------------------------------------------------------
    caption_index = CaptionIndex()
    fig_seq = 0
    tbl_seq = 0
    # F05：preamble 里的题注也纳入索引（section_id 用 _preamble）
    _scan_sections = ([preamble_sec] if preamble_sec and preamble_sec.blocks else []) \
        + [parsed[sid] for sid in sorted_sids]
    for sec in _scan_sections:
        for blk in sec.blocks:
            if blk.kind != "p":
                continue
            stripped = blk.text.strip()
            if is_fig_caption(stripped):
                fig_seq += 1
                m = RE_FIG_CAPTION.match(stripped)
                orig_num = int(m.group(1)) if m else fig_seq
                caption_index.add(Caption(
                    entity_id=f"_Ref_图_{fig_seq}", caption_type="图",
                    orig_num=orig_num, seq_num=fig_seq,
                    section_id=sec.section_id, source_file=blk.source_file,
                ))
            elif is_tbl_caption(stripped):
                tbl_seq += 1
                m = RE_TBL_CAPTION.match(stripped)
                orig_num = int(m.group(1)) if m else tbl_seq
                caption_index.add(Caption(
                    entity_id=f"_Ref_表_{tbl_seq}", caption_type="表",
                    orig_num=orig_num, seq_num=tbl_seq,
                    section_id=sec.section_id, source_file=blk.source_file,
                ))
    log_msg(f"      题注：图 {fig_seq} 个，表 {tbl_seq} 个")
    # 统计跨文件重复原编号（F04 风险提示）
    dup_keys = [k for k, v in caption_index._by_key.items() if len(v) > 1]
    if dup_keys:
        log_msg(f"      ℹ️ 发现 {len(dup_keys)} 组跨文件/跨章节重复编号："
                + "、".join(f"{t}{n}({len(caption_index._by_key[(t,n)])}个)" for t, n in dup_keys))

    # -----------------------------------------------------------------------
    # 创建文档 & 封面 & 目录
    # -----------------------------------------------------------------------
    log_msg(f"[2/5] 加载统一格式")
    doc = Document()
    s = doc.sections[0]
    s.page_height = Cm(29.7)
    s.page_width = Cm(21.0)
    s.top_margin = Cm(2.54)
    s.bottom_margin = Cm(2.54)
    s.left_margin = Cm(3.18)
    s.right_margin = Cm(3.18)
    _ensure_style(doc, "Normal", base="Normal", font_name=fc.body_font, size=Pt(fc.body_size))
    _ensure_style(doc, "Heading 1", base="Normal", font_name=fc.h1_font, size=Pt(fc.h1_size), bold=fc.h1_bold, color=_color_from_hex(fc.h1_color))
    _ensure_style(doc, "Heading 2", base="Normal", font_name=fc.h2_font, size=Pt(fc.h2_size), bold=fc.h2_bold, color=_color_from_hex(fc.h2_color))
    _ensure_style(doc, "Heading 3", base="Normal", font_name=fc.h3_font, size=Pt(fc.h3_size), bold=fc.h3_bold, color=_color_from_hex(fc.h3_color))
    _ensure_style(doc, "Heading 4", base="Normal", font_name=fc.h4_font, size=Pt(fc.h4_size), bold=fc.h4_bold, color=_color_from_hex(fc.h4_color))
    _ensure_style(doc, "Heading 5", base="Normal", font_name=fc.h5_font, size=Pt(fc.h5_size), bold=fc.h5_bold, color=_color_from_hex(fc.h5_color))
    _ensure_style(doc, "Heading 6", base="Normal", font_name=fc.h6_font, size=Pt(fc.h6_size), bold=fc.h6_bold, color=_color_from_hex(fc.h6_color))

    header_data = {"title": title, "org": org, "date": date, "period": period}
    add_header_footer(doc, header_items, header_data)
    hf_desc = " | ".join(header_data.get(k, "") for k in header_items) if header_items else "无"
    log_msg(f"      页眉：{hf_desc}")
    log_msg(f"      页脚：页码（第 X 页）")

    log_msg(f"[3/5] 拼装封面")
    log_msg(f"      封面主标题（报告主标题）：{title!r}")
    add_cover(doc, period=period, title=title, subtitle=subtitle, org=org, date=date)

    log_msg(f"[4/5] 插入目录与正文")
    add_toc(doc)

    # -----------------------------------------------------------------------
    # 按 sorted_sids 顺序写入正文（完全由输入文件决定结构）
    # -----------------------------------------------------------------------
    fig_seq = 0
    tbl_seq = 0
    written_blocks = 0
    resolve_log: List[str] = []  # F04：引用歧义/无匹配审计

    # F05：先写前导内容（前言）——首标题前的摘要/引言/免责声明等，不再丢弃
    if preamble_sec and preamble_sec.blocks:
        add_heading(doc, 1, "前言", fc)
        for blk in preamble_sec.blocks:
            if blk.kind == "p":
                stripped = blk.text.strip()
                if is_fig_caption(stripped):
                    fig_seq += 1
                    add_caption(doc, blk, "图", fig_seq)
                elif is_tbl_caption(stripped):
                    tbl_seq += 1
                    add_caption(doc, blk, "表", tbl_seq)
                else:
                    add_body_with_xref(doc, blk, caption_index, fc, resolve_log)
            elif blk.kind in ("bullet", "number"):
                # F06：列表项（含编号列表）走 add_bullet，保留 run 格式+层级
                add_bullet(doc, blk, fc)
            elif blk.kind == "table":
                _copy_table_to_doc(blk.table, doc)
                doc.add_paragraph()
            elif blk.kind == "image":
                _copy_image_paragraph_to_doc(blk.src_para_elem, blk.src_doc, doc)
            elif blk.kind == "rich":
                _copy_image_paragraph_to_doc(blk.src_para_elem, blk.src_doc, doc)
            written_blocks += 1

    for sid in sorted_sids:
        sec = parsed[sid]
        # 优先用识别器给出的 level（可能来自 Word 标题样式/大纲级别，比纯点数推断更准）；
        level = sec.level if sec.level > 0 else _level_from_sid(sid)

        if level == 1:
            heading_text = f"第{_cn(sid)}章  {sec.title}"
        else:
            heading_text = f"{sid}  {sec.title}"
        add_heading(doc, level, heading_text, fc)

        for blk in sec.blocks:
            if blk.kind == "p":
                stripped = blk.text.strip()
                if is_fig_caption(stripped):
                    fig_seq += 1
                    add_caption(doc, blk, "图", fig_seq)
                elif is_tbl_caption(stripped):
                    tbl_seq += 1
                    add_caption(doc, blk, "表", tbl_seq)
                else:
                    add_body_with_xref(doc, blk, caption_index, fc, resolve_log)
            elif blk.kind in ("bullet", "number"):
                add_bullet(doc, blk, fc)
            elif blk.kind == "table":
                _copy_table_to_doc(blk.table, doc)
                doc.add_paragraph()
            elif blk.kind == "image":
                _copy_image_paragraph_to_doc(blk.src_para_elem, blk.src_doc, doc)
            elif blk.kind == "rich":
                # F03：含超链接/外链图的段落走保真复制+关系迁移，避免关系丢失
                _copy_image_paragraph_to_doc(blk.src_para_elem, blk.src_doc, doc)
            written_blocks += 1

    # F04：引用解析审计
    if resolve_log:
        ambig = [s for s in resolve_log if "歧义" in s]
        nomatch = [s for s in resolve_log if "无匹配" in s]
        if ambig:
            log_msg(f"      ⚠️ 发现 {len(ambig)} 处引用歧义（已保留原文，需人工核对）：")
            for s in ambig:
                log_msg(f"        {s}")
        if nomatch:
            log_msg(f"      ℹ️ {len(nomatch)} 处引用无对应题注（已保留原文）")

    log_msg(f"[5/5] 保存：{output_path}")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    doc.save(output_path)

    # 守恒校验：写入块数应等于仍在 parsed 中的输入块
    # （删除/隐藏章节已在此前 pop，skipped_blocks 仅用于审计）
    if written_blocks != total_input_blocks:
        log_msg(
            f"      ⚠️ 守恒校验失败：期望写入 {total_input_blocks} 块，"
            f"实际写入 {written_blocks} 块"
            f"{f'（另已排除 {skipped_blocks} 块）' if skipped_blocks else ''}，请排查数据丢失"
        )
    else:
        extra = f"（另已排除 {skipped_blocks} 块）" if skipped_blocks else ""
        log_msg(f"      守恒校验通过：输入/输出块数一致（{written_blocks} 块）{extra}")

    # F03：输出包完整性校验——重新打开生成文件，确认 ZIP 部件与内部关系目标有效。
    integrity_ok, integrity_msg = _verify_docx_integrity(output_path)
    if integrity_ok:
        log_msg(f"      包完整性校验通过：{integrity_msg}")
    else:
        log_msg(f"      ⚠️ 包完整性校验失败：{integrity_msg}")
    log_msg("完成 ✓")

    return {
        "input_dir": input_dir,
        "output_path": output_path,
        "sections_found": sorted_sids,
        "image_count": image_count,
        "table_count": table_count,
        "bold_paragraphs": bold_count,
        "fig_captions": fig_seq,
        "tbl_captions": tbl_seq,
        "integrity_check": {"ok": integrity_ok, "message": integrity_msg},
        # F10：失败文件留痕（无论是否跳过，都记录在结果中供调用方/用户审计）
        "parsed_files": list(parsed.parsed_files),
        "skipped_files": [
            {"file": fl.file_name, "category": fl.category, "message": fl.message}
            for fl in parsed.failures
        ],
        "log": log,
    }


def _cn(n: str) -> str:
    """阿拉伯数字 -> 中文（支持 1-20）。"""
    table = {
        "1": "一", "2": "二", "3": "三", "4": "四", "5": "五",
        "6": "六", "7": "七", "8": "八", "9": "九", "10": "十",
        "11": "十一", "12": "十二", "13": "十三", "14": "十四", "15": "十五",
        "16": "十六", "17": "十七", "18": "十八", "19": "十九", "20": "二十",
    }
    return table.get(n, n)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None):
    import argparse
    parser = argparse.ArgumentParser(description="保险行业季度调研报告自动汇总工具")
    parser.add_argument("input_dir", help="输入目录（含各部门的 docx 文件）")
    parser.add_argument("output_path", help="输出 docx 路径")
    parser.add_argument("--period", default="2026年第二季度")
    parser.add_argument("--title", default="中国保险行业调研报告")
    parser.add_argument("--subtitle", default="市场环境 · 保费增长 · 渠道变革 · 趋势展望")
    parser.add_argument("--org", default="保险行业研究中心")
    parser.add_argument("--date", default="2026年7月")
    parser.add_argument(
        "--header-items", nargs="*", default=["title"],
        choices=["title", "org", "date", "period"],
        help="页眉显示项，可选 title/org/date/period，空列表传 --header-items 不带值",
    )
    args = parser.parse_args(argv)
    aggregate(
        args.input_dir,
        args.output_path,
        period=args.period,
        title=args.title,
        subtitle=args.subtitle,
        org=args.org,
        date=args.date,
        header_items=args.header_items,
    )


if __name__ == "__main__":
    main()
