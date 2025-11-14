param(
    [string]$DebugRoot = "",  # 如果为空，自动检测 x64 或 Any CPU 目录
    [switch]$RegisterCom = $true  # 默认启用 COM 注册，因为 WPS 需要
)

Write-Host ">>> WPS Debug 注册脚本启动" -ForegroundColor Cyan

# 如果未指定路径，自动检测 x64 或 Any CPU 输出目录
if ([string]::IsNullOrEmpty($DebugRoot)) {
    $x64Path = "D:\CODES\PPA\PPA\bin\x64\Debug"
    $anyCpuPath = "D:\CODES\PPA\PPA\bin\Debug"
    
    if (Test-Path $x64Path) {
        $DebugRoot = $x64Path
        Write-Host ">>> 自动检测到 x64 输出目录：$DebugRoot" -ForegroundColor Green
    } elseif (Test-Path $anyCpuPath) {
        $DebugRoot = $anyCpuPath
        Write-Host ">>> 自动检测到 Any CPU 输出目录：$DebugRoot" -ForegroundColor Yellow
        Write-Host ">>> 提示：建议使用 x64 配置编译以确保与 64 位 WPS 完全兼容" -ForegroundColor Yellow
    } else {
        throw "未找到输出目录。请先编译项目（推荐使用 Debug|x64 配置），或手动指定 -DebugRoot 参数。"
    }
}

if (!(Test-Path $DebugRoot)) {
    throw "指定的 Debug 目录不存在：$DebugRoot"
}

$manifestPath = Join-Path $DebugRoot "PPA.vsto"
if (!(Test-Path $manifestPath)) {
    throw "未找到 PPA.vsto，确保先在 Debug 模式下编译。路径：$manifestPath"
}

$dllPath = Join-Path $DebugRoot "PPA.dll"
if (!(Test-Path $dllPath)) {
    throw "未找到 PPA.dll，确保先在 Debug 模式下编译。路径：$dllPath"
}

# 生成 manifest 字符串（VSTO 要求 file:/// 前缀，路径中使用正斜杠）
$manifestValue = "file:///" + ($manifestPath -replace '\\', '/')

# WPS 使用的注册表路径（根据实际注册表信息）
$wpsKey = "HKCU:\Software\Kingsoft\Office\WPS\Addins\wpp\PPA.Debug"

# 也保留 Office\Addins 路径（兼容性）
$mainKey = "HKCU:\Software\Kingsoft\Office\Addins\PPA.Debug"
$wowKey = "HKCU:\Software\WOW6432Node\Kingsoft\Office\Addins\PPA.Debug"

function Set-AddInKey {
    param(
        [string]$KeyPath
    )

    New-Item -Path $KeyPath -Force | Out-Null
    Set-ItemProperty -Path $KeyPath -Name FriendlyName -Value "PPA Debug" -Type String
    Set-ItemProperty -Path $KeyPath -Name Description -Value "PPA Debug Build Add-in" -Type String
    Set-ItemProperty -Path $KeyPath -Name Manifest -Value $manifestValue -Type String
    Set-ItemProperty -Path $KeyPath -Name LoadBehavior -Value 3 -Type DWord
    Write-Host "注册表配置完成：$KeyPath" -ForegroundColor Green
}

function Set-WpsAddInKey {
    param(
        [string]$KeyPath
    )

    New-Item -Path $KeyPath -Force | Out-Null
    # WPS 使用 path、name、description、load 字段
    Set-ItemProperty -Path $KeyPath -Name "path" -Value $dllPath -Type String
    Set-ItemProperty -Path $KeyPath -Name "name" -Value "PPA Add-in" -Type String
    Set-ItemProperty -Path $KeyPath -Name "description" -Value "PPA Add-in for WPS Debug" -Type String
    Set-ItemProperty -Path $KeyPath -Name "load" -Value 1 -Type DWord
    Write-Host "WPS 注册表配置完成：$KeyPath" -ForegroundColor Green
}

Write-Host ">>> 写入注册表键值..." -ForegroundColor Cyan

# 关键：WPS 也会读取 Microsoft Office 的注册表路径
$officeKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPA"
Write-Host ">>> 写入 Microsoft Office 注册表路径（WPS 也会读取）..." -ForegroundColor Cyan
New-Item -Path $officeKey -Force | Out-Null
Set-ItemProperty -Path $officeKey -Name FriendlyName -Value "PPA" -Type String
Set-ItemProperty -Path $officeKey -Name Description -Value "PPA" -Type String
Set-ItemProperty -Path $officeKey -Name Manifest -Value ($manifestValue + "|vstolocal") -Type String
Set-ItemProperty -Path $officeKey -Name LoadBehavior -Value 3 -Type DWord
Write-Host "Microsoft Office 注册表配置完成：$officeKey" -ForegroundColor Green

# WPS 专用路径
Set-WpsAddInKey -KeyPath $wpsKey
# 兼容性路径
Set-AddInKey -KeyPath $mainKey
Set-AddInKey -KeyPath $wowKey

# WPS 会在 WPP\AddinsWL 下查找加载项（字符串值，不是子键）
Write-Host ">>> 写入 WPP\AddinsWL 注册表值..." -ForegroundColor Cyan
$addinsWLKeys = @(
    "HKCU:\Software\Kingsoft\Office\WPP\AddinsWL",
    "HKCU:\Software\WOW6432Node\Kingsoft\Office\WPP\AddinsWL"
)

# 尝试写入 HKLM（需要管理员权限，失败则跳过）
$addinsWLKeys += "HKLM:\Software\Kingsoft\Office\WPP\AddinsWL"
$addinsWLKeys += "HKLM:\Software\WOW6432Node\Kingsoft\Office\WPP\AddinsWL"

foreach ($key in $addinsWLKeys) {
    try {
        # 确保父键存在
        $parentKey = Split-Path $key -Parent
        if (!(Test-Path $parentKey)) {
            New-Item -Path $parentKey -Force | Out-Null
        }
        if (!(Test-Path $key)) {
            New-Item -Path $key -Force | Out-Null
        }
        # 写入字符串值（名称为 PPA.Debug 和 PPA，值为空字符串）
        Set-ItemProperty -Path $key -Name "PPA.Debug" -Value "" -Type String -ErrorAction Stop
        Set-ItemProperty -Path $key -Name "PPA" -Value "" -Type String -ErrorAction Stop
        Write-Host "WPP\AddinsWL 配置完成：$key" -ForegroundColor Green
    } catch {
        Write-Host "警告：无法写入 $key（可能需要管理员权限）：$_" -ForegroundColor Yellow
    }
}

if ($RegisterCom) {
    # 64 位 WPS 必须使用 64 位 regasm
    $framework64 = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
    $framework32 = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\regasm.exe"

    # 优先使用 64 位 regasm（因为 WPS 是 64 位）
    if (Test-Path $framework64) {
        $regasm = $framework64
        Write-Host ">>> 检测到 64 位 WPS，使用 64 位 regasm" -ForegroundColor Cyan
    } elseif (Test-Path $framework32) {
        $regasm = $framework32
        Write-Host ">>> 警告：未找到 64 位 regasm，使用 32 位（可能导致兼容性问题）" -ForegroundColor Yellow
    } else {
        throw "未找到 regasm.exe，请检查 .NET Framework 是否安装完整。"
    }

    Write-Host (">>> 注册 COM 互操作（需要管理员权限）：{0} ""{1}""" -f $regasm, $dllPath) -ForegroundColor Cyan
    & $regasm $dllPath /codebase /tlb
    if ($LASTEXITCODE -ne 0) {
        throw "regasm 注册失败，错误码 $LASTEXITCODE。请确保以管理员身份运行。"
    }
    Write-Host "regasm 注册成功。" -ForegroundColor Green
}

Write-Host ">>> WPS Debug 加载项注册完成。" -ForegroundColor Cyan
Write-Host ("Manifest：{0}" -f $manifestValue) -ForegroundColor Yellow
Write-Host "" -ForegroundColor Yellow
Write-Host "重要提示：" -ForegroundColor Yellow
Write-Host "1. 如果 COM 注册失败，请以管理员身份运行此脚本。" -ForegroundColor Yellow
Write-Host "2. 执行 VSTOInstaller：& 'C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe' /Install `"$manifestPath`"" -ForegroundColor Yellow
Write-Host "3. 重启 WPS，打开“选项 → 插件管理”确认加载项已启用。" -ForegroundColor Yellow
Write-Host "4. 若仍无显示，可在插件管理中手动“添加”指向 PPA.vsto。" -ForegroundColor Yellow


