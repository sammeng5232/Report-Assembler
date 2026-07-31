#Requires -Version 5.1
<#
.SYNOPSIS
  处理 Windows「智能应用控制 / Smart App Control」拦截本工具 exe 的问题。

.DESCRIPTION
  1) 解除 Zone.Identifier 封锁
  2) 本机自签 + 安装到当前用户/本机信任存储（管理员时更彻底）
  3) 尝试将智能应用控制设为关闭（需管理员；强制模式常无法仅靠签名绕过）
  4) 打开设置页；并提供源码启动（推荐，立刻可用）

  用法（建议右键「使用 PowerShell 运行」；若仍拦 exe 请「以管理员身份运行」）：
  powershell -ExecutionPolicy Bypass -File ".\解除智能应用控制拦截.ps1"
#>
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$exe = Join-Path $root "保险报告生成器.exe"
$tool = Join-Path $root "工具"
$signScript = Join-Path $tool "sign_and_trust.ps1"
$desktopApp = Join-Path $tool "desktop_app.py"

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Unblock-Tree([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch {}
            # 直接删 Zone.Identifier ADS
            try {
                $zi = $_.FullName + ":Zone.Identifier"
                if (Test-Path -LiteralPath $zi) { Remove-Item -LiteralPath $zi -Force -ErrorAction SilentlyContinue }
            } catch {}
        }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " 保险报告生成器 · 解除智能应用控制拦截" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
$admin = Test-IsAdmin
if ($admin) {
    Write-Host "当前：管理员权限" -ForegroundColor Green
} else {
    Write-Host "当前：普通用户（若仍被拦，请右键本脚本 → 以管理员身份运行）" -ForegroundColor Yellow
}
Write-Host ""

# ---- 1. 解除封锁 ----
Write-Host "[1/5] 解除文件封锁标记..." -ForegroundColor Yellow
Unblock-Tree $root
if (Test-Path -LiteralPath $exe) {
    try { Unblock-File -LiteralPath $exe -ErrorAction SilentlyContinue } catch {}
}
Write-Host "      完成。" -ForegroundColor Green

# ---- 2. 签名 + 信任证书 ----
Write-Host "[2/5] 本机代码签名并信任证书..." -ForegroundColor Yellow
if (Test-Path -LiteralPath $signScript) {
    try {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $signScript
    } catch {
        Write-Host "      签名脚本异常：$($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "      未找到 sign_and_trust.ps1，跳过。" -ForegroundColor DarkYellow
}

# 管理员：额外装到 LocalMachine 的 Root / TrustedPublisher（对部分策略更有效）
if ($admin -and (Test-Path -LiteralPath $exe)) {
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $exe
        $cert = $sig.SignerCertificate
        if ($cert) {
            $tmpCer = Join-Path $env:TEMP "report_assembler_local_lm.cer"
            Export-Certificate -Cert $cert -FilePath $tmpCer | Out-Null
            foreach ($store in @("Root", "TrustedPublisher")) {
                $o = & certutil -addstore $store $tmpCer 2>&1 | Out-String
                Write-Host "      LocalMachine\$store : $($o.Trim() -replace '\s+', ' ')" -ForegroundColor DarkGray
            }
            Remove-Item $tmpCer -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "      本机存储导入跳过：$($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

if (Test-Path -LiteralPath $exe) {
    $st = Get-AuthenticodeSignature -LiteralPath $exe
    Write-Host "      exe 签名状态: $($st.Status)" -ForegroundColor $(if ($st.Status -eq 'Valid') { 'Green' } else { 'Yellow' })
}

# ---- 3. 尝试关闭 SAC（仅管理员；自签无法在强制模式下保证放行）----
Write-Host "[3/5] 尝试调整智能应用控制策略..." -ForegroundColor Yellow
$sacKey = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
$sacState = $null
try {
    if (Test-Path $sacKey) {
        $sacState = (Get-ItemProperty -Path $sacKey -Name "VerifiedAndReputablePolicyState" -ErrorAction SilentlyContinue).VerifiedAndReputablePolicyState
    }
} catch {}
$stateName = switch ("$sacState") {
    "0" { "关闭(0)" }
    "1" { "评估(1)" }
    "2" { "强制(2)" }
    default { "未知($sacState)" }
}
Write-Host "      当前策略状态: $stateName" -ForegroundColor White

if ($admin) {
    # 0=Off 1=Evaluation 2=Enforcement
    # 强制模式往往只能改为关闭；改为评估在部分版本也可用
    try {
        if (-not (Test-Path $sacKey)) {
            New-Item -Path $sacKey -Force | Out-Null
        }
        # 优先关到 0；若策略锁死会失败
        Set-ItemProperty -Path $sacKey -Name "VerifiedAndReputablePolicyState" -Value 0 -Type DWord -Force -ErrorAction Stop
        Write-Host "      已写入 VerifiedAndReputablePolicyState=0（关闭）。" -ForegroundColor Green
        Write-Host "      若资源管理器仍拦截，请重启电脑后再试 exe。" -ForegroundColor Yellow
    } catch {
        Write-Host "      无法通过注册表修改（可能被策略锁定）：$($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-Host "      请手动：Windows 安全中心 → 应用和浏览器控制 → 智能应用控制 → 关闭" -ForegroundColor White
    }
} else {
    Write-Host "      非管理员，跳过注册表修改。" -ForegroundColor DarkYellow
}

# ---- 4. 打开设置 ----
Write-Host "[4/5] 打开 Windows 安全中心相关设置..." -ForegroundColor Yellow
try { Start-Process "windowsdefender://appbrowser" } catch {
    try { Start-Process "ms-settings:windowsdefender" } catch {}
}
Write-Host ""
Write-Host "请在设置中确认：" -ForegroundColor White
Write-Host "  Windows 安全中心 → 应用和浏览器控制 → 智能应用控制设置 → 关闭" -ForegroundColor White
Write-Host "  （自签证书无法在「强制」模式下保证放行 PyInstaller 程序）" -ForegroundColor DarkGray
Write-Host ""

# ---- 5. 源码启动（立刻可用）----
Write-Host "[5/5] 准备「源码启动」（推荐，不受 exe 拦截影响）..." -ForegroundColor Yellow

$pyCandidates = @(
    (Join-Path $tool ".venv_build\Scripts\pythonw.exe"),
    (Join-Path $tool ".venv_build\Scripts\python.exe"),
    "C:\Users\mengz\.workbuddy\binaries\python\versions\3.13.12\python.exe",
    (Get-Command pythonw -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source),
    (Get-Command python -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
) | Where-Object { $_ -and (Test-Path $_) }

$py = $null
foreach ($c in $pyCandidates) {
    try {
        $probe = $c
        if ($c -like "*pythonw.exe") {
            $probe = ($c -replace "pythonw\.exe$", "python.exe")
        }
        if (-not (Test-Path $probe)) { $probe = $c }
        & $probe -c "import tkinter" 2>$null
        if ($LASTEXITCODE -eq 0) { $py = $c; break }
    } catch {}
}

$launchBat = Join-Path $root "启动报告生成器.bat"
$launchSrc = Join-Path $root "启动报告生成器(源码).bat"
# 确保 bat 存在（可能被旧版破坏）
@"
@echo off
chcp 65001 >nul
set "ROOT=%~dp0"
cd /d "%ROOT%"
if exist "%ROOT%工具\.venv_build\Scripts\pythonw.exe" (
  start "" "%ROOT%工具\.venv_build\Scripts\pythonw.exe" "%ROOT%工具\desktop_app.py"
  exit /b 0
)
if exist "%ROOT%工具\.venv_build\Scripts\python.exe" (
  start "" "%ROOT%工具\.venv_build\Scripts\python.exe" "%ROOT%工具\desktop_app.py"
  exit /b 0
)
if exist "%ROOT%保险报告生成器.exe" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "try{Unblock-File -LiteralPath '%ROOT%保险报告生成器.exe' -EA SilentlyContinue}catch{}"
  start "" "%ROOT%保险报告生成器.exe"
  exit /b 0
)
echo 无法启动。请安装 Python 或运行解除智能应用控制拦截.ps1
pause
"@ | Set-Content -Path $launchBat -Encoding ASCII

@"
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
where python >nul 2>&1
if %errorlevel%==0 (
  python desktop_app.py
  if errorlevel 1 pause
  exit /b %errorlevel%
)
echo 未找到 Python。
pause
exit /b 1
"@ | Set-Content -Path $launchSrc -Encoding ASCII

if ($py -and (Test-Path -LiteralPath $desktopApp)) {
    Write-Host "      已就绪源码启动：$py" -ForegroundColor Green
    Write-Host "      也可双击：启动报告生成器.bat" -ForegroundColor Green
    Write-Host ""
    # 无交互环境自动启动；有控制台则询问
    $auto = $env:REPORT_SAC_AUTO
    if ($auto -eq "1" -or -not [Environment]::UserInteractive) {
        $ans = "Y"
    } else {
        try {
            $ans = Read-Host "是否现在用 Python 源码启动界面（推荐，立刻可用）？(Y/N)"
        } catch {
            $ans = "Y"
        }
    }
    if ($ans -match '^[Yy]' -or [string]::IsNullOrWhiteSpace($ans)) {
        $arg = "`"$desktopApp`""
        if ($py -like "*pythonw.exe") {
            Start-Process -FilePath $py -ArgumentList $desktopApp -WorkingDirectory $tool
        } else {
            Start-Process -FilePath $py -ArgumentList $desktopApp -WorkingDirectory $tool
        }
        Write-Host "      已启动界面。" -ForegroundColor Green
    }
} else {
    Write-Host "      未检测到可用 Python + tkinter。" -ForegroundColor DarkYellow
    Write-Host "      请关闭智能应用控制后双击 保险报告生成器.exe" -ForegroundColor White
    Write-Host "      或在「工具」目录创建 venv 后 pip install -r requirements.txt" -ForegroundColor White
}

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor Cyan
Write-Host "说明：" -ForegroundColor Cyan
Write-Host " • 智能应用控制「强制」模式会拦无微软信誉的本地 exe；自签通常不够。" -ForegroundColor Gray
Write-Host " • 本机自用推荐：双击「启动报告生成器.bat」（走 Python，不被当未知 exe 拦）。" -ForegroundColor Gray
Write-Host " • 或关闭智能应用控制后再用 保险报告生成器.exe。" -ForegroundColor Gray
Write-Host "----------------------------------------" -ForegroundColor Cyan
Write-Host ""
if ([Environment]::UserInteractive) {
    Write-Host "按回车退出..."
    try { Read-Host | Out-Null } catch {}
}
