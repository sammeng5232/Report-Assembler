@echo off
chcp 65001 >nul
set "ROOT=%~dp0"
cd /d "%ROOT%"

rem 优先源码启动（不受「智能应用控制」拦截 PyInstaller exe 的影响）
if exist "%ROOT%工具\.venv_build\Scripts\pythonw.exe" (
  start "" "%ROOT%工具\.venv_build\Scripts\pythonw.exe" "%ROOT%工具\desktop_app.py"
  exit /b 0
)
if exist "%ROOT%工具\.venv_build\Scripts\python.exe" (
  start "" "%ROOT%工具\.venv_build\Scripts\python.exe" "%ROOT%工具\desktop_app.py"
  exit /b 0
)

rem 回退：尝试直接运行已签名 exe
if exist "%ROOT%保险报告生成器.exe" (
  rem 先清封锁标记
  powershell -NoProfile -ExecutionPolicy Bypass -Command "try{Unblock-File -LiteralPath '%ROOT%保险报告生成器.exe' -EA SilentlyContinue}catch{}"
  start "" "%ROOT%保险报告生成器.exe"
  exit /b 0
)

echo 找不到启动入口。请运行「解除智能应用控制拦截.ps1」或安装 Python。
pause
exit /b 1
