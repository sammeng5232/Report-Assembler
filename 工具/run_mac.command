#!/bin/bash
# 双击或在终端执行：在 Mac 上启动「保险报告生成器」
set -e
cd "$(dirname "$0")"

echo "=== 保险报告生成器 (macOS) ==="

# 优先用 python3
if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo "未找到 Python3。请先安装："
  echo "  1) 打开 https://www.python.org/downloads/macos/ 安装官方包，或"
  echo "  2) brew install python"
  read -r -p "按回车退出…"
  exit 1
fi

echo "Python: $($PY --version 2>&1)"

# 虚拟环境（首次自动创建）
if [ ! -d ".venv" ]; then
  echo "首次运行：创建虚拟环境并安装依赖…"
  $PY -m venv .venv
  # shellcheck disable=SC1091
  source .venv/bin/activate
  python -m pip install --upgrade pip
  python -m pip install -r requirements.txt
else
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi

# Tk 可用性检查
python - <<'PY'
import sys
try:
    import tkinter
except Exception as e:
    print("Tkinter 不可用：", e)
    print("若使用 Homebrew Python，可尝试：brew install python-tk")
    sys.exit(1)
print("Tkinter OK")
PY

echo "启动界面…"
exec python desktop_app.py
