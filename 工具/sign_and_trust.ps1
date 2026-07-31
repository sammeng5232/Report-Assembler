# 为本机「保险报告生成器」做完整代码签名 + 信任证书 + 解除封锁
# 可被 build_win.ps1 或「安装信任并放行.ps1」调用
# 用法：powershell -ExecutionPolicy Bypass -File .\sign_and_trust.ps1
# 可选环境变量：SIGN_ALL_PE=1  签名 _internal 内全部 exe/dll/pyd（更利于通过本机策略）
$ErrorActionPreference = "Stop"
$outRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$exe = Join-Path $outRoot "保险报告生成器.exe"
$internal = Join-Path $outRoot "_internal"
$signAll = ($env:SIGN_ALL_PE -eq "1") -or ($args -contains "-All")

if (-not (Test-Path $exe)) {
    throw "未找到：$exe"
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Remove-ZoneId([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch {}
            try {
                $zi = $_.FullName + ":Zone.Identifier"
                if (Test-Path -LiteralPath $zi) {
                    Remove-Item -LiteralPath $zi -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }
}

Write-Host "==> 解除文件封锁标记 (Zone.Identifier) ..."
Remove-ZoneId $outRoot

Write-Host "==> 准备本机代码签名证书 ..."
$subject = "CN=Insurance Report Assembler Local Build"
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date) -and $_.HasPrivateKey } |
    Select-Object -First 1

if (-not $cert) {
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $subject `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -NotAfter (Get-Date).AddYears(10) `
        -FriendlyName "Insurance Report Assembler Local" `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
    Write-Host "    已新建证书: $($cert.Thumbprint)"
} else {
    Write-Host "    复用证书: $($cert.Thumbprint)"
}

# 导出 CER 装入信任存储（当前用户 + 若管理员则本机）
$tmpCer = Join-Path $env:TEMP "report_assembler_local.cer"
try {
    Export-Certificate -Cert $cert -FilePath $tmpCer | Out-Null
    foreach ($storeName in @("Root", "TrustedPublisher")) {
        $out = & certutil -user -addstore $storeName $tmpCer 2>&1 | Out-String
        Write-Host "    CurrentUser\$storeName : OK"
    }
    if (Test-IsAdmin) {
        foreach ($storeName in @("Root", "TrustedPublisher")) {
            $null = & certutil -addstore $storeName $tmpCer 2>&1
            Write-Host "    LocalMachine\$storeName : OK"
        }
    }
} catch {
    Write-Host "    警告：导入信任存储：$($_.Exception.Message)"
} finally {
    Remove-Item $tmpCer -Force -ErrorAction SilentlyContinue
}

function Sign-OneFile([string]$path, $certificate, [bool]$useTimestamp) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    # 已是本证书有效签名则跳过
    try {
        $cur = Get-AuthenticodeSignature -LiteralPath $path
        if ($cur.Status -eq "Valid" -and $cur.SignerCertificate -and
            $cur.SignerCertificate.Thumbprint -eq $certificate.Thumbprint) {
            return $cur
        }
    } catch {}
    $params = @{
        FilePath      = $path
        Certificate   = $certificate
        HashAlgorithm = "SHA256"
    }
    if ($useTimestamp) {
        try {
            return Set-AuthenticodeSignature @params -TimestampServer "http://timestamp.digicert.com" -ErrorAction Stop
        } catch {
            return Set-AuthenticodeSignature @params
        }
    }
    return Set-AuthenticodeSignature @params
}

Write-Host "==> 签名主程序 ..."
$r = Sign-OneFile $exe $cert $true
Write-Host "    主程序: $($r.Status) $($r.StatusMessage)"

# 同步签名 dist 副本（若存在）
$distExe = Join-Path $PSScriptRoot "dist\保险报告生成器\保险报告生成器.exe"
if (Test-Path $distExe) {
    $null = Sign-OneFile $distExe $cert $true
}

if ($signAll -and (Test-Path $internal)) {
    Write-Host "==> 签名 _internal 内 PE（exe/dll/pyd）..."
    $files = Get-ChildItem -LiteralPath $internal -Recurse -Include *.exe, *.dll, *.pyd -File -ErrorAction SilentlyContinue
    $ok = 0; $fail = 0; $n = 0
    foreach ($f in $files) {
        $n++
        try {
            # 大批量不打时间戳，避免超时/限流
            $sr = Sign-OneFile $f.FullName $cert $false
            if ($sr -and $sr.Status -eq "Valid") { $ok++ } else { $fail++ }
        } catch { $fail++ }
        if (($n % 40) -eq 0) { Write-Host "    ... $n / $($files.Count)" }
    }
    Write-Host "    完成: 成功 $ok / 共 $($files.Count)（失败 $fail）"
} else {
    Write-Host "==> 跳过全量 PE 签名（需要时设 SIGN_ALL_PE=1 或以 -All 运行）"
}

# 再次解除封锁
Remove-ZoneId $outRoot
if (Test-Path $internal) {
    Get-ChildItem $internal -Recurse -Include *.exe, *.dll, *.pyd -ErrorAction SilentlyContinue |
        ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -EA SilentlyContinue } catch {} }
}

Write-Host ""
Write-Host "签名完成。主程序状态: $((Get-AuthenticodeSignature -LiteralPath $exe).Status)"
Write-Host "若仍被智能应用控制/Defender 拦截，请右键以管理员运行上级目录："
Write-Host "  安装信任并放行.ps1"
