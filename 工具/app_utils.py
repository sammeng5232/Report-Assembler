"""Web 与桌面界面共用的安全文件处理工具（跨平台）。"""

from __future__ import annotations

import hashlib
import os
import platform
import re
import shutil
import subprocess
import sys
from typing import Dict, List, Sequence, Tuple


def is_windows() -> bool:
    return sys.platform.startswith("win")


def is_macos() -> bool:
    return sys.platform == "darwin"


def ui_font_family() -> str:
    """界面默认中文字体：按系统选择可用家族名。"""
    if is_macos():
        return "PingFang SC"
    if is_windows():
        return "微软雅黑"
    return "Noto Sans CJK SC"


def default_doc_fonts() -> Dict[str, str]:
    """报告正文/标题默认字体（写入 DOCX 的 eastAsia 字体名）。"""
    if is_macos():
        # macOS 常见中文字体；Word/Pages 均可识别
        return {
            "ui": "PingFang SC",
            "heading": "PingFang SC",
            "body": "Songti SC",
            "options": [
                "PingFang SC", "Heiti SC", "Songti SC", "STSong",
                "STHeiti", "Hiragino Sans GB", "Arial Unicode MS",
                "微软雅黑", "宋体", "黑体", "楷体", "仿宋",
            ],
        }
    if is_windows():
        return {
            "ui": "微软雅黑",
            "heading": "微软雅黑",
            "body": "宋体",
            "options": ["微软雅黑", "宋体", "黑体", "楷体", "仿宋", "PingFang SC", "Songti SC"],
        }
    return {
        "ui": "Noto Sans CJK SC",
        "heading": "Noto Sans CJK SC",
        "body": "Noto Serif CJK SC",
        "options": [
            "Noto Sans CJK SC", "Noto Serif CJK SC", "WenQuanYi Micro Hei",
            "微软雅黑", "宋体", "黑体",
        ],
    }


def default_desktop_dir() -> str:
    """用户桌面目录（兼容中文「桌面」与英文 Desktop）。"""
    home = os.path.expanduser("~")
    for name in ("Desktop", "桌面"):
        p = os.path.join(home, name)
        if os.path.isdir(p):
            return p
    return home


def open_path(path: str) -> None:
    """用系统默认方式打开文件或文件夹（Win / macOS / Linux）。"""
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    if is_windows():
        os.startfile(path)  # type: ignore[attr-defined]
        return
    if is_macos():
        subprocess.run(["open", path], check=False)
        return
    # Linux / 其它
    opener = shutil.which("xdg-open") or shutil.which("gio")
    if opener:
        subprocess.run([opener, path], check=False)
    else:
        raise OSError("当前系统无法打开文件（缺少 xdg-open）")

_WIN_RESERVED = frozenset({
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
})


def sanitize_display_name(text, default: str = "未命名周期", max_len: int = 50) -> str:
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    cleaned = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9_-]", "", text).strip("-_")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip("-_")
    if not cleaned or cleaned.upper() in _WIN_RESERVED:
        return default
    return cleaned


def build_output_paths(output_dir: str, period: str) -> Dict[str, str]:
    safe = sanitize_display_name(period)
    base = f"保险行业季度调研报告_终版_{safe}"
    root = os.path.abspath(output_dir)
    return {
        "docx": os.path.join(root, base + ".docx"),
        "pdf": os.path.join(root, base + ".pdf"),
        "base": base,
        "dir": root,
    }


def copy_selected_files(files: Sequence[str], target_dir: str) -> Tuple[int, List[dict], List[dict]]:
    os.makedirs(target_dir, exist_ok=True)
    basename_count: Dict[str, int] = {}
    records: List[dict] = []
    for idx, src in enumerate(files, start=1):
        src = os.path.abspath(src)
        name = os.path.basename(src)
        basename_count[name] = basename_count.get(name, 0) + 1
        stem, ext = os.path.splitext(name)
        disk_name = f"{idx:03d}_{stem}{ext or '.docx'}"
        dst = os.path.join(target_dir, disk_name)
        shutil.copy2(src, dst)
        h = hashlib.sha256()
        with open(dst, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        records.append({
            "display_name": name, "disk_name": disk_name, "path": dst,
            "sha256": h.hexdigest()[:16],
        })
    conflicts = [{"name": n, "count": c} for n, c in basename_count.items() if c > 1]
    return len(records), conflicts, records


def build_section_preview_rows(parsed) -> List[dict]:
    from report_aggregator import sort_section_ids

    rows: List[dict] = []
    for order, sid in enumerate(sort_section_ids(parsed), start=1):
        sec = parsed[sid]
        level = sec.level if getattr(sec, "level", 0) > 0 else (
            0 if sid == "_preamble" else sid.count(".") + 1
        )
        if sid == "_preamble":
            layer = "前言"
        else:
            layer = "章" if level == 1 else ("节" if level == 2 else "小节")
        if getattr(sec, "title_conflict", False):
            status = "标题冲突"
        elif getattr(sec, "merge_count", 1) > 1:
            status = f"合并{sec.merge_count}次"
        elif getattr(sec, "is_preamble", False) or sid == "_preamble":
            status = "前言"
        elif getattr(sec, "confidence", "") == "low":
            status = "低置信度"
        else:
            status = "正常"
        conf_map = {"high": "高", "medium": "中", "low": "低"}
        conf = conf_map.get(getattr(sec, "confidence", ""), getattr(sec, "confidence", ""))
        if level not in (1, 2, 3):
            level = 1 if layer == "章" else 2 if layer == "节" else 3
        rows.append({
            "order": order, "section_id": sid, "title": sec.title, "layer": layer,
            "level": level, "block_count": len(sec.blocks),
            "source_count": len(getattr(sec, "source_files", []) or []),
            "confidence": conf, "status": status,
        })
    return rows


def layer_to_level(layer: str, sid: str = "") -> int:
    layer = (layer or "").strip()
    if layer in ("章", "一级", "1", "前言"):
        return 1
    if layer in ("节", "二级", "2"):
        return 2
    if layer in ("小节", "三级", "3"):
        return 3
    try:
        lv = int(layer)
        if lv in (1, 2, 3):
            return lv
    except ValueError:
        pass
    return min(sid.count(".") + 1, 3) if sid else 2
