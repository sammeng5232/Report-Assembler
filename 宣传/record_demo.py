# -*- coding: utf-8 -*-
"""
真实操作演示录制（默认后台，不抢桌面）。

默认模式 --background：
  - 启动真实 GUI 演示流程
  - 窗口移到屏外，不置顶、不杀其它程序、不抓全桌面
  - 用 Win32 PrintWindow 抓取该窗口画面 → ffmpeg 编码 MP4
  - 你可继续正常使用电脑

前台全桌面模式（会打扰，不推荐）：
  python record_demo.py --desktop
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import subprocess
import shutil
import struct
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "工具"
OUT_DIR = ROOT / "宣传"
DEMO_OUT = ROOT / "演示输出"
MARKER = DEMO_OUT / "_demo_done.txt"
MP4 = OUT_DIR / "保险报告生成器_操作演示.mp4"
APP_TITLE = "保险行业季度调研报告生成器"


def find_ffmpeg() -> str:
    w = shutil.which("ffmpeg")
    if w:
        return w
    base = Path.home() / "AppData/Local/Microsoft/WinGet/Packages"
    if base.is_dir():
        for p in base.rglob("ffmpeg.exe"):
            return str(p)
    raise SystemExit("未找到 ffmpeg，请先安装：winget install Gyan.FFmpeg")


def find_hwnd(title_substr: str, timeout: float = 30.0):
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    found = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def enum_proc(hwnd, _lp):
        if not user32.IsWindowVisible(hwnd):
            # 屏外窗口有时仍 visible；也检查有标题的
            pass
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        if title_substr in buf.value:
            found.append(hwnd)
        return True

    deadline = time.time() + timeout
    while time.time() < deadline:
        found.clear()
        user32.EnumWindows(enum_proc, 0)
        if found:
            return found[0]
        time.sleep(0.3)
    return None


def capture_window_bgr(hwnd):
    """PrintWindow 抓取窗口客户区，返回 (w, h, bgr_bytes) 或 None。"""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    rect = RECT()
    if not user32.GetClientRect(hwnd, ctypes.byref(rect)):
        return None
    w = int(rect.right - rect.left)
    h = int(rect.bottom - rect.top)
    if w < 80 or h < 80:
        return None

    hwnd_dc = user32.GetDC(hwnd)
    if not hwnd_dc:
        return None
    mem_dc = gdi32.CreateCompatibleDC(hwnd_dc)
    bmp = gdi32.CreateCompatibleBitmap(hwnd_dc, w, h)
    old = gdi32.SelectObject(mem_dc, bmp)

    # PW_CLIENTONLY=1 | PW_RENDERFULLCONTENT=2 → 3
    ok = user32.PrintWindow(hwnd, mem_dc, 3)
    if not ok:
        # 回退：BitBlt（屏外窗口可能失败）
        gdi32.BitBlt(mem_dc, 0, 0, w, h, hwnd_dc, 0, 0, 0x00CC0020)

    # BITMAPINFOHEADER
    bi = struct.pack(
        "<IiiHHIIiiII",
        40, w, -h, 1, 32, 0, 0, 0, 0, 0, 0,
    )
    buf_size = w * h * 4
    buf = (ctypes.c_char * buf_size)()
    gdi32.GetDIBits(mem_dc, bmp, 0, h, buf, bi, 0)

    gdi32.SelectObject(mem_dc, old)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(hwnd, hwnd_dc)

    # BGRA → BGR
    raw = bytes(buf)
    bgr = bytearray(w * h * 3)
    si = 0
    di = 0
    for _ in range(w * h):
        bgr[di] = raw[si]
        bgr[di + 1] = raw[si + 1]
        bgr[di + 2] = raw[si + 2]
        si += 4
        di += 3
    return w, h, bytes(bgr)


def record_window_background(ffmpeg: str, hwnd, mp4: Path, duration_sec: float, fps: int = 8) -> int:
    """后台抓窗编码，不占用全桌面 gdigrab。"""
    first = None
    for _ in range(50):
        first = capture_window_bgr(hwnd)
        if first:
            break
        time.sleep(0.2)
    if not first:
        print("FAIL: cannot capture window")
        return 1
    w, h, _ = first
    # 偶数尺寸（yuv420）
    w -= w % 2
    h -= h % 2
    log_path = OUT_DIR / "record_ffmpeg.log"
    cmd = [
        ffmpeg, "-y",
        "-f", "rawvideo",
        "-pix_fmt", "bgr24",
        "-s", f"{w}x{h}",
        "-r", str(fps),
        "-i", "-",
        "-an",
        "-c:v", "libx264",
        "-preset", "veryfast",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        str(mp4),
    ]
    print(f"window capture {w}x{h} @ {fps}fps for ~{duration_sec}s → {mp4}")
    log_f = open(log_path, "w", encoding="utf-8", errors="replace")
    rec = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=log_f,
    )
    assert rec.stdin is not None
    t0 = time.time()
    frame_interval = 1.0 / fps
    frames = 0
    try:
        while time.time() - t0 < duration_sec:
            t_frame = time.time()
            cap = capture_window_bgr(hwnd)
            if cap:
                cw, ch, data = cap
                # 尺寸变化时简单裁切/填充到固定 w,h
                if cw != w or ch != h:
                    # 跳过异常帧，写上一帧黑屏占位太麻烦，直接 pad/crop 最简：跳过
                    pass
                else:
                    rec.stdin.write(data)
                    frames += 1
            # 演示完成后再录约 4 秒尾部即可结束（不必空等到 duration）
            if MARKER.is_file() and frames > fps * 15:
                if time.time() - MARKER.stat().st_mtime >= 4.0:
                    print("demo done + 4s tail, stop capture")
                    break
            elapsed = time.time() - t_frame
            sleep = frame_interval - elapsed
            if sleep > 0:
                time.sleep(sleep)
    finally:
        try:
            rec.stdin.close()
        except Exception:
            pass
        try:
            rec.wait(timeout=60)
        except subprocess.TimeoutExpired:
            rec.kill()
        log_f.close()
    print(f"wrote {frames} frames, ffmpeg code={rec.returncode}")
    return 0 if rec.returncode == 0 else 1


def record_desktop_foreground(ffmpeg: str, mp4: Path, duration_sec: int) -> subprocess.Popen:
    """全桌面录制（会录到你的其它窗口，不推荐）。"""
    log_path = OUT_DIR / "record_ffmpeg.log"
    cmd = [
        ffmpeg, "-y",
        "-f", "gdigrab",
        "-framerate", "12",
        "-t", str(duration_sec),
        "-i", "desktop",
        "-vf", "scale=1600:-2",
        "-c:v", "libx264",
        "-preset", "veryfast",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        str(mp4),
    ]
    log_f = open(log_path, "w", encoding="utf-8", errors="replace")
    return subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=log_f)


def main() -> int:
    parser = argparse.ArgumentParser(description="录制报告生成器演示（默认后台）")
    parser.add_argument(
        "--desktop",
        action="store_true",
        help="前台全桌面 gdigrab（会干扰正在使用的电脑，不推荐）",
    )
    parser.add_argument("--duration", type=int, default=70, help="最长录制秒数")
    parser.add_argument("--fps", type=int, default=8, help="后台抓窗帧率")
    args = parser.parse_args()

    ffmpeg = find_ffmpeg()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    DEMO_OUT.mkdir(parents=True, exist_ok=True)
    if MARKER.exists():
        MARKER.unlink()
    if MP4.exists():
        MP4.unlink()

    # 只结束本演示相关的旧 app，不杀浏览器等
    try:
        subprocess.run(
            ["taskkill", "/F", "/FI", f"WINDOWTITLE eq {APP_TITLE}*"],
            capture_output=True,
        )
    except Exception:
        pass

    env = os.environ.copy()
    env["REPORT_DEMO"] = "1"
    env["PYTHONUTF8"] = "1"
    if not args.desktop:
        env["REPORT_DEMO_BG"] = "1"
    else:
        env.pop("REPORT_DEMO_BG", None)

    py = sys.executable
    app_cmd = [py, str(TOOLS / "desktop_app.py")]
    mode = "DESKTOP(会打扰)" if args.desktop else "BACKGROUND(不抢桌面)"
    print(f"mode: {mode}")
    print("ffmpeg:", ffmpeg)
    print("output:", MP4)
    print("start app…")
    app = subprocess.Popen(app_cmd, cwd=str(TOOLS), env=env)

    if args.desktop:
        time.sleep(2.0)
        rec = record_desktop_foreground(ffmpeg, MP4, args.duration)
        deadline = time.time() + args.duration + 40
        done = False
        while time.time() < deadline:
            if MARKER.is_file() and not done:
                print("demo marker found")
                done = True
            if rec.poll() is not None:
                break
            time.sleep(0.5)
        try:
            rec.wait(timeout=30)
        except subprocess.TimeoutExpired:
            rec.kill()
        rc_ffmpeg = rec.returncode or 0
    else:
        hwnd = find_hwnd(APP_TITLE, timeout=25)
        if not hwnd:
            print("FAIL: app window not found")
            app.terminate()
            return 1
        print(f"hwnd={hwnd}, capturing off-screen window…")
        # 稍等界面布局稳定
        time.sleep(1.0)
        rc_ffmpeg = record_window_background(ffmpeg, hwnd, MP4, args.duration, fps=args.fps)
        done = MARKER.is_file()

    try:
        app.terminate()
        app.wait(timeout=5)
    except Exception:
        try:
            app.kill()
        except Exception:
            pass

    if not MP4.is_file() or MP4.stat().st_size < 30_000:
        print("FAIL: mp4 missing/small")
        log = OUT_DIR / "record_ffmpeg.log"
        if log.is_file():
            print(log.read_text(encoding="utf-8", errors="replace")[-2000:])
        return 1

    ffprobe = str(Path(ffmpeg).with_name("ffprobe.exe"))
    if Path(ffprobe).is_file():
        pr = subprocess.run(
            [ffprobe, "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", str(MP4)],
            capture_output=True, text=True,
        )
        if pr.returncode != 0:
            print("FAIL: mp4 unreadable", pr.stderr)
            return 1
        print("duration_sec≈", pr.stdout.strip())

    print("OK:", MP4, f"({MP4.stat().st_size / 1024 / 1024:.1f} MB)")
    if not done:
        print("WARNING: demo marker not seen")
    if MARKER.is_file():
        print("demo outputs:\n" + MARKER.read_text(encoding="utf-8"))
    return 0 if rc_ffmpeg == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
