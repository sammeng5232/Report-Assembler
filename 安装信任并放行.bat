@echo off
chcp 65001 >nul
:: 自动请求管理员权限，执行完整信任/签名/Defender排除/关闭智能应用控制
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo 正在请求管理员权限...
  powershell -NoProfile -Command "Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0安装信任并放行.ps1\"' -Verb RunAs"
  exit /b 0
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0安装信任并放行.ps1"
pause
