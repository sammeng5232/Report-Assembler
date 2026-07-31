#!/bin/bash
# 在 MacBook 本机执行，打包成 .app（不能在 Windows 上交叉编译 macOS 应用）
set -e
cd "$(dirname "$0")"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "请在 macOS 上运行本脚本。"
  exit 1
fi

python3 -m venv .venv 2>/dev/null || true
# shellcheck disable=SC1091
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

# onedir .app
python -m PyInstaller \
  --noconfirm --clean --windowed \
  --name "保险报告生成器" \
  --osx-bundle-identifier "com.insurance.report.generator" \
  desktop_app.py

OUT="dist/保险报告生成器.app"
echo ""
echo "打包完成：$OUT"
echo "可拖到「应用程序」或任意文件夹；首次打开若被拦截："
echo "  系统设置 → 隐私与安全性 → 仍要打开"
echo "或：xattr -cr \"$OUT\""
