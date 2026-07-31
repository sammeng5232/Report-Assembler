#Requires -Version 5.1
<#
.SYNOPSIS
  管理员一键：信任本机签名证书 + 全量签名 + Defender 排除 + 关闭/放宽智能应用控制。

.DESCRIPTION
  自签无法获得「微软云信誉」，但在本机可通过：
  - 证书装入本机「受信任的根 / 受信任的发布者」
  - 主程序与 _internal 内 PE 全部 Authenticode 签名
  - Windows Defender 排除本目录（避免 PUA 误杀 PyInstaller）
  - 将智能应用控制从强制改为关闭（若策略允许）

  双击「安装信任并放行.bat」会自动请求管理员权限。
#>
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$tool = Join-Path $root "工具"
$exe = Join-Path $root "保险报告生成器.exe"
$signScript = Join-Path $tool "sign_and_trust.ps1"

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# 非管理员：自我提权
if (-not (Test-IsAdmin)) {
    Write-Host "需要管理员权限以写入本机证书存储 / Defender 排除 / 智能应用控制策略..." -ForegroundColor Yellow
    $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    try {
        Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $arg -Wait
    } catch {
        Write-Host "用户取消了管理员授权，或提权失败：$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "请右键本脚本 → 以管理员身份运行。" -ForegroundColor Yellow
        if ([Environment]::UserInteractive) { Read-Host "按回车退出"; }
        exit 1
    }
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " 安装信任并放行（管理员）" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "目录: $root"
Write-Host ""

# 1) 全量签名
Write-Host "[1/4] 代码签名（主程序 + _internal 全部 PE）..." -ForegroundColor Yellow
$env:SIGN_ALL_PE = "1"
if (Test-Path $signScript) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $signScript -All
} else {
    Write-Host "  未找到 sign_and_trust.ps1" -ForegroundColor Red
}

# 2) Defender 排除（PUA 常误杀 PyInstaller 包）
Write-Host "[2/4] 添加 Windows Defender 排除项..." -ForegroundColor Yellow
try {
    if (Get-Command Add-MpPreference -ErrorAction SilentlyContinue) {
        Add-MpPreference -ExclusionPath $root -ErrorAction Stop
        if (Test-Path $exe) {
            Add-MpPreference -ExclusionProcess $exe -ErrorAction SilentlyContinue
            Add-MpPreference -ExclusionPath $exe -ErrorAction SilentlyContinue
        }
        # 关闭对本机用户的部分激进阻断（若策略允许）
        try {
            Set-MpPreference -PUAProtection Enabled -ErrorAction SilentlyContinue
            # 不强制关 PUA 全局，只靠路径排除更安全
        } catch {}
        Write-Host "  已排除路径: $root" -ForegroundColor Green
    } else {
        Write-Host "  无 Defender cmdlet，跳过。" -ForegroundColor DarkYellow
    }
} catch {
    Write-Host "  Defender 排除失败：$($_.Exception.Message)" -ForegroundColor DarkYellow
}

# 3) 智能应用控制：关闭（0）。评估=1 强制=2
Write-Host "[3/4] 调整智能应用控制策略..." -ForegroundColor Yellow
$sacKey = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
try {
    if (-not (Test-Path $sacKey)) { New-Item -Path $sacKey -Force | Out-Null }
    $before = $null
    try {
        $before = (Get-ItemProperty -Path $sacKey -Name VerifiedAndReputablePolicyState -EA SilentlyContinue).VerifiedAndReputablePolicyState
    } catch {}
    Write-Host "  调整前: $before （0关/1评估/2强制）"
    # 先写 0（关闭）。部分 SKU 只允许评估→关闭单向。
    Set-ItemProperty -Path $sacKey -Name "VerifiedAndReputablePolicyState" -Value 0 -Type DWord -Force
    $after = (Get-ItemProperty -Path $sacKey -Name VerifiedAndReputablePolicyState).VerifiedAndReputablePolicyState
    Write-Host "  调整后: $after" -ForegroundColor Green
    if ("$before" -eq "2") {
        Write-Host "  曾为「强制」模式：建议重启一次后再双击 exe。" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  无法修改 SAC 注册表：$($_.Exception.Message)" -ForegroundColor DarkYellow
    Write-Host "  请手动：Windows 安全中心 → 应用和浏览器控制 → 智能应用控制 → 关闭" -ForegroundColor White
}

# 4) 验证并尝试启动
Write-Host "[4/4] 验证签名并启动..." -ForegroundColor Yellow
if (Test-Path $exe) {
    $sig = Get-AuthenticodeSignature -LiteralPath $exe
    Write-Host "  签名: $($sig.Status) | 发布者: $($sig.SignerCertificate.Subject)"
    try { Unblock-File -LiteralPath $exe -EA SilentlyContinue } catch {}
    # 尝试启动
    try {
        Start-Process -FilePath $exe -WorkingDirectory $root
        Write-Host "  已尝试启动 exe。" -ForegroundColor Green
    } catch {
        Write-Host "  启动失败：$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  请重启后再试，或用 启动报告生成器.bat" -ForegroundColor Yellow
    }
} else {
    Write-Host "  未找到 exe" -ForegroundColor Red
}

Write-Host ""
Write-Host "完成。" -ForegroundColor Cyan
Write-Host "若仍提示拦截：1) 重启电脑  2) 再双击 保险报告生成器.exe" -ForegroundColor White
Write-Host "说明：自签没有「互联网信誉」，换电脑后仍可能被拦；多机分发需购买代码签名证书或 Azure Trusted Signing。" -ForegroundColor DarkGray
Write-Host ""
if ([Environment]::UserInteractive) {
    Write-Host "按回车退出..."
    try { Read-Host | Out-Null } catch {}
}
