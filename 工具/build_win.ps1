# 在 Windows 本机打包「保险报告生成器.exe」（onedir：exe + _internal）
# 用法：在 工具 目录执行  powershell -ExecutionPolicy Bypass -File .\build_win.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$outRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$pyCandidates = @(
    "C:\Users\mengz\.workbuddy\binaries\python\versions\3.13.12\python.exe",
    "C:\Users\mengz\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe",
    (Get-Command python -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
) | Where-Object { $_ -and (Test-Path $_) }

$py = $null
foreach ($c in $pyCandidates) {
    try {
        & $c -c "import tkinter" 2>$null
        if ($LASTEXITCODE -eq 0) { $py = $c; break }
    } catch {}
}
if (-not $py) {
    throw "未找到带 tkinter 的 Python。请安装 Python 3.10–3.13（勾选 tcl/tk）。"
}
Write-Host "Using Python: $py"

if (-not (Test-Path ".venv_build")) {
    & $py -m venv .venv_build
}
& ".\.venv_build\Scripts\python.exe" -m pip install -q --upgrade pip
& ".\.venv_build\Scripts\python.exe" -m pip install -q -r requirements.txt pyinstaller pywin32

if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }

$verFile = Join-Path $PSScriptRoot "version_info.txt"
$piArgs = @(
    "-m", "PyInstaller",
    "--noconfirm", "--clean", "--windowed", "--onedir",
    "--name", "保险报告生成器",
    "--paths", ".",
    "--hidden-import", "app_utils",
    "--hidden-import", "report_aggregator",
    "--hidden-import", "i18n",
    "--hidden-import", "docx",
    "--hidden-import", "lxml",
    "--hidden-import", "win32com",
    "--hidden-import", "win32com.client"
)
if (Test-Path $verFile) {
    $piArgs += @("--version-file", $verFile)
}
$piArgs += "desktop_app.py"
& ".\.venv_build\Scripts\python.exe" @piArgs

$dist = Join-Path $PSScriptRoot "dist\保险报告生成器"
if (-not (Test-Path (Join-Path $dist "保险报告生成器.exe"))) {
    throw "打包失败：未生成 exe"
}
if (-not (Test-Path (Join-Path $dist "_internal\_tk_data"))) {
    throw "打包失败：未包含 tkinter 运行时"
}

Copy-Item (Join-Path $dist "保险报告生成器.exe") (Join-Path $outRoot "保险报告生成器.exe") -Force
if (Test-Path (Join-Path $outRoot "_internal")) {
    Remove-Item (Join-Path $outRoot "_internal") -Recurse -Force
}
Copy-Item (Join-Path $dist "_internal") (Join-Path $outRoot "_internal") -Recurse -Force

# 解除封锁 + 本机代码签名（含 _internal 全部 PE，降低本机策略拦截）
$sign = Join-Path $PSScriptRoot "sign_and_trust.ps1"
if (Test-Path $sign) {
    Write-Host "签名并解除封锁（SIGN_ALL_PE=1）..."
    $env:SIGN_ALL_PE = "1"
    & powershell -NoProfile -ExecutionPolicy Bypass -File $sign -All
}

Write-Host ""
Write-Host "打包完成："
Write-Host "  $(Join-Path $outRoot '保险报告生成器.exe')"
Write-Host "  $(Join-Path $outRoot '_internal\')"
Write-Host "请保持 exe 与 _internal 在同一目录。"
Write-Host "若被智能应用控制拦截，请运行上级目录：解除智能应用控制拦截.ps1"
