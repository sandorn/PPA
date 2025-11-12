# PPA 插件快速发布脚本
# 使用方法：在 PowerShell 中运行 .\发布脚本.ps1
# 注意：此文件应使用 UTF-8 with BOM 编码保存

param(
    [string]$Configuration = "Release",
    [string]$PublishPath = "publish"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PPA 插件发布脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查是否在项目根目录
if (-not (Test-Path "PPA\PPA.csproj")) {
    Write-Host "错误：未找到项目文件。请在项目根目录运行此脚本。" -ForegroundColor Red
    exit 1
}

# 检查 MSBuild 是否可用
$msbuildPath = ""
$vsPath = "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
if (Test-Path $vsPath) {
    $msbuildPath = $vsPath
} else {
    $vsPath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    if (Test-Path $vsPath) {
        $msbuildPath = $vsPath
    } else {
        # 尝试使用系统 PATH 中的 MSBuild
        $msbuildPath = "MSBuild.exe"
    }
}

Write-Host "使用 MSBuild: $msbuildPath" -ForegroundColor Yellow
Write-Host ""

# 步骤 1: 清理之前的构建
Write-Host "[1/4] 清理之前的构建..." -ForegroundColor Green
& $msbuildPath "PPA\PPA.sln" /t:Clean /p:Configuration=$Configuration /p:Platform="Any CPU" /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    Write-Host "警告：清理失败，继续执行..." -ForegroundColor Yellow
}

# 步骤 2: 构建项目
Write-Host "[2/4] 构建项目 ($Configuration)..." -ForegroundColor Green
& $msbuildPath "PPA\PPA.sln" /t:Build /p:Configuration=$Configuration /p:Platform="Any CPU" /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误：构建失败！" -ForegroundColor Red
    exit 1
}
Write-Host "✓ 构建成功" -ForegroundColor Green
Write-Host ""

# 步骤 3: 发布 ClickOnce
Write-Host "[3/4] 发布 ClickOnce 安装程序..." -ForegroundColor Green
$publishAbs = Join-Path (Get-Location) "publish"
if (-not (Test-Path $publishAbs)) { New-Item -ItemType Directory -Path $publishAbs | Out-Null }
$outputPath = Join-Path "PPA" "bin\$Configuration\"
& $msbuildPath "PPA\PPA.csproj" /t:Publish /p:Configuration=$Configuration /p:Platform=AnyCPU /p:OutputPath=$outputPath /p:PublishDir=$publishAbs\ /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误：发布失败！" -ForegroundColor Red
    exit 1
}
Write-Host "✓ 发布成功" -ForegroundColor Green
Write-Host ""

# 步骤 4: 检查输出文件
Write-Host "[4/4] 检查输出文件..." -ForegroundColor Green
$setupExe = Join-Path $publishAbs "setup.exe"
$appManifest = Join-Path $publishAbs "PPA.application"

if (Test-Path $setupExe) {
    Write-Host "✓ 找到安装程序: $setupExe" -ForegroundColor Green
} else {
    Write-Host "⚠ 警告：未找到 setup.exe" -ForegroundColor Yellow
}

if (Test-Path $appManifest) {
    Write-Host "✓ 找到应用程序清单: $appManifest" -ForegroundColor Green
} else {
    # 递归查找 .application 文件
    $foundManifest = Get-ChildItem -Path $publishAbs -Recurse -Filter "*.application" -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq "PPA.application" } | Select-Object -First 1
    if (-not $foundManifest) {
        $foundManifest = Get-ChildItem -Path $publishAbs -Recurse -Filter "*.application" -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if ($foundManifest) {
        Write-Host "✓ 找到应用程序清单: $($foundManifest.FullName)" -ForegroundColor Green
    } else {
        Write-Host "⚠ 警告：未找到 PPA.application" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "发布完成！" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "发布文件位置: $publishAbs" -ForegroundColor Yellow
Write-Host ""
Write-Host "下一步：" -ForegroundColor Cyan
Write-Host "1. 检查发布文件夹中的文件" -ForegroundColor White
Write-Host "2. 测试安装程序" -ForegroundColor White
Write-Host "3. 将整个 publish 文件夹打包分发" -ForegroundColor White
Write-Host ""

# 询问是否打开发布文件夹
$openFolder = Read-Host "是否打开发布文件夹？(Y/N)"
if ($openFolder -eq "Y" -or $openFolder -eq "y") {
    if (Test-Path $publishAbs) {
        Start-Process explorer.exe -ArgumentList "`"$publishAbs`""
    } else {
        Write-Host "Warning: Publish folder not found: $publishAbs" -ForegroundColor Yellow
    }
}

