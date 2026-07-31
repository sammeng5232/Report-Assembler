@echo off
chcp 65001 >nul
cd /d "%~dp0工具"

if exist ".venv_build\Scripts\pythonw.exe" (
  start "" ".venv_build\Scripts\pythonw.exe" desktop_app.py
  exit /b 0
)
if exist ".venv_build\Scripts\python.exe" (
  ".venv_build\Scripts\python.exe" desktop_app.py
  if errorlevel 1 pause
  exit /b %errorlevel%
)

where pythonw >nul 2>&1
if %errorlevel%==0 (
  start "" pythonw desktop_app.py
  exit /b 0
)
where python >nul 2>&1
if %errorlevel%==0 (
  python desktop_app.py
  if errorlevel 1 pause
  exit /b %errorlevel%
)

echo 未找到 Python。请先安装 Python 3.10–3.13（勾选 tcl/tk），
echo 或在「工具」目录执行：
echo   python -m venv .venv_build
echo   .venv_build\Scripts\pip install -r requirements.txt
echo   .venv_build\Scripts\python desktop_app.py
echo.
echo 若只想用 exe：请右键以管理员运行「解除智能应用控制拦截.ps1」
pause
exit /b 1
