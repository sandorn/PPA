# PPA 代码审查自动化扫描脚本
# 用于检查代码规范违反情况

param(
    [switch]$All,
    [switch]$Application,
    [switch]$Logging,
    [switch]$ComLifecycle,
    [switch]$DependencyInjection,
    [switch]$Formatting
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sourcePath = Join-Path $projectRoot "PPA"

Write-Host "=== PPA 代码审查扫描 ===" -ForegroundColor Cyan
Write-Host "项目根目录: $projectRoot" -ForegroundColor Gray
Write-Host ""

$issues = @()
$warnings = @()

# 1. Application 调用规范检查
if ($All -or $Application) {
    Write-Host "1. 检查 Application 调用规范..." -ForegroundColor Yellow
    
    # 检查直接访问 Globals.ThisAddIn
    $globalsAccess = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "Globals\.ThisAddIn" | 
        Where-Object { $_.Line -notmatch "//.*Globals" -and $_.Line -notmatch "禁止.*Globals" }
    
    if ($globalsAccess) {
        foreach ($match in $globalsAccess) {
            $issues += [PSCustomObject]@{
                Category = "Application调用"
                Severity = "错误"
                File = $match.Path
                Line = $match.LineNumber
                Message = "直接访问 Globals.ThisAddIn，应使用 ApplicationHelper"
                Code = $match.Line.Trim()
            }
        }
    }
    
    # 检查缓存 Application 对象
    $cachedApp = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "private.*Application.*_|private.*NETOP\.Application" |
        Where-Object { $_.Line -notmatch "//.*允许" }
    
    if ($cachedApp) {
        foreach ($match in $cachedApp) {
            $warnings += [PSCustomObject]@{
                Category = "Application调用"
                Severity = "警告"
                File = $match.Path
                Line = $match.LineNumber
                Message = "可能缓存了 Application 对象，建议通过 ApplicationHelper 获取"
                Code = $match.Line.Trim()
            }
        }
    }
    
    # 检查未受保护的 GetNativeComApplication 调用
    $nativeComCalls = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "GetNativeComApplication" |
        ForEach-Object {
            $file = Get-Content $_.Path
            $lineNum = $_.LineNumber
            $context = $file[($lineNum - 5)..($lineNum + 5)] -join "`n"
            
            # 检查是否有注释说明或 Guard 保护
            if ($context -notmatch "NativeComGuard|//.*允许|//.*必要") {
                [PSCustomObject]@{
                    Category = "Application调用"
                    Severity = "警告"
                    File = $_.Path
                    Line = $lineNum
                    Message = "GetNativeComApplication 调用可能未受 Guard 保护"
                    Code = $_.Line.Trim()
                }
            }
        }
    
    if ($nativeComCalls) {
        $warnings += $nativeComCalls
    }
}

# 2. 日志记录规范检查
if ($All -or $Logging) {
    Write-Host "2. 检查日志记录规范..." -ForegroundColor Yellow
    
    # 检查 Profiler.LogMessage 调用
    $profilerCalls = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "Profiler\.LogMessage" |
        Where-Object { 
            $_.Path -notmatch "Profiler\.cs|ProfilerLoggerAdapter\.cs" -and
            $_.Line -notmatch "//.*允许|//.*遗留"
        }
    
    if ($profilerCalls) {
        foreach ($match in $profilerCalls) {
            $issues += [PSCustomObject]@{
                Category = "日志记录"
                Severity = "错误"
                File = $match.Path
                Line = $match.LineNumber
                Message = "使用 Profiler.LogMessage，应改为 ILogger"
                Code = $match.Line.Trim()
            }
        }
    }
    
    # 检查日志是否包含上下文信息（简单启发式检查）
    $loggerCalls = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "_logger\.Log(Information|Error|Warning)" |
        Where-Object { 
            $_.Line -match 'Log(Information|Error|Warning)\(".*"\)' -and
            $_.Line -notmatch '\$|参数|类型|数量|启动|完成'
        }
    
    if ($loggerCalls) {
        foreach ($match in $loggerCalls) {
            $warnings += [PSCustomObject]@{
                Category = "日志记录"
                Severity = "警告"
                File = $match.Path
                Line = $match.LineNumber
                Message = "日志消息可能缺少上下文信息"
                Code = $match.Line.Trim()
            }
        }
    }
}

# 3. COM 对象生命周期检查
if ($All -or $ComLifecycle) {
    Write-Host "3. 检查 COM 对象生命周期..." -ForegroundColor Yellow
    
    # 检查循环中的 COM 对象收集（简单模式匹配）
    $csFiles = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse
    
    foreach ($file in $csFiles) {
        $content = Get-Content $file.FullName -Raw
        $lines = Get-Content $file.FullName
        
        # 检查 foreach 循环中收集 NETOP 对象到列表
        if ($content -match 'foreach.*NETOP\.\w+.*in.*\{[\s\S]{0,500}?\.Add\(') {
            $lineNum = 1
            $inForeach = $false
            $hasAdd = $false
            
            foreach ($line in $lines) {
                if ($line -match 'foreach.*NETOP\.') {
                    $inForeach = $true
                    $hasAdd = $false
                }
                if ($inForeach -and $line -match '\.Add\(') {
                    $hasAdd = $true
                }
                if ($inForeach -and $line -match '\}' -and $hasAdd) {
                    # 检查是否有 DisposeAll 调用
                    $remainingLines = $lines[$lineNum..($lines.Count - 1)] -join "`n"
                    if ($remainingLines -notmatch 'DisposeAll|using.*Dispose') {
                        $warnings += [PSCustomObject]@{
                            Category = "COM生命周期"
                            Severity = "警告"
                            File = $file.FullName
                            Line = $lineNum
                            Message = "循环中收集 COM 对象到列表，可能需要在循环后释放"
                            Code = $line.Trim()
                        }
                    }
                    $inForeach = $false
                }
                $lineNum++
            }
        }
        
        # 检查双重循环中的 Row/Cell 对象
        if ($content -match 'for.*rows.*\{[\s\S]{0,300}?for.*cols.*\{[\s\S]{0,200}?\.Cells\[') {
            $warnings += [PSCustomObject]@{
                Category = "COM生命周期"
                Severity = "警告"
                File = $file.FullName
                Line = 0
                Message = "双重循环中创建 Row/Cell 对象，检查是否需要释放"
                Code = ""
            }
        }
    }
    
    # 检查未使用 using 的 Shapes 集合访问
    $shapesAccess = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "\.Shapes\s*\{|foreach.*\.Shapes" |
        Where-Object { 
            $fileContent = Get-Content $_.Path -Raw
            $lineNum = $_.LineNumber
            $context = $fileContent.Substring([Math]::Max(0, $fileContent.IndexOf($_.Line) - 100), 200)
            $context -notmatch "using.*Shapes"
        }
    
    if ($shapesAccess) {
        foreach ($match in $shapesAccess) {
            $warnings += [PSCustomObject]@{
                Category = "COM生命周期"
                Severity = "警告"
                File = $match.Path
                Line = $match.LineNumber
                Message = "访问 Shapes 集合，建议使用 using 语句释放"
                Code = $match.Line.Trim()
            }
        }
    }
}

# 4. 依赖注入检查
if ($All -or $DependencyInjection) {
    Write-Host "4. 检查依赖注入规范..." -ForegroundColor Yellow
    
    # 检查直接 new 服务实例
    $directNew = Get-ChildItem -Path $sourcePath -Filter "*.cs" -Recurse | 
        Select-String -Pattern "new (Table|Text|Chart|Shape)(Format|Batch)Helper" |
        Where-Object { 
            $_.Path -notmatch "ServiceCollectionExtensions|Test" -and
            $_.Line -notmatch "//.*允许|//.*测试"
        }
    
    if ($directNew) {
        foreach ($match in $directNew) {
            $issues += [PSCustomObject]@{
                Category = "依赖注入"
                Severity = "错误"
                File = $match.Path
                Line = $match.LineNumber
                Message = "直接 new 服务实例，应通过依赖注入获取"
                Code = $match.Line.Trim()
            }
        }
    }
}

# 5. 格式化操作检查
if ($All -or $Formatting) {
    Write-Host "5. 检查格式化操作规范..." -ForegroundColor Yellow
    
    # 检查是否缺少 UndoHelper.BeginUndoEntry
    $formatMethods = Get-ChildItem -Path $sourcePath -Filter "*BatchHelper.cs" -Recurse | 
        Select-String -Pattern "Format(Text|Tables|Charts)" |
        ForEach-Object {
            $file = Get-Content $_.Path
            $methodStart = $_.LineNumber
            $methodContent = $file[($methodStart - 1)..($methodStart + 50)] -join "`n"
            
            if ($methodContent -notmatch "BeginUndoEntry") {
                [PSCustomObject]@{
                    Category = "格式化操作"
                    Severity = "警告"
                    File = $_.Path
                    Line = $methodStart
                    Message = "格式化方法可能缺少 UndoHelper.BeginUndoEntry 调用"
                    Code = $_.Line.Trim()
                }
            }
        }
    
    if ($formatMethods) {
        $warnings += $formatMethods
    }
    
    # 检查是否缺少 Toast 反馈
    $batchMethods = Get-ChildItem -Path $sourcePath -Filter "*BatchHelper.cs" -Recurse | 
        Select-String -Pattern "Format(Text|Tables|Charts)Internal" |
        ForEach-Object {
            $file = Get-Content $_.Path
            $methodStart = $_.LineNumber
            $methodContent = $file[($methodStart - 1)..($methodStart + 100)] -join "`n"
            
            if ($methodContent -notmatch "Toast\.Show") {
                [PSCustomObject]@{
                    Category = "格式化操作"
                    Severity = "警告"
                    File = $_.Path
                    Line = $methodStart
                    Message = "批量操作方法可能缺少 Toast 用户反馈"
                    Code = $_.Line.Trim()
                }
            }
        }
    
    if ($batchMethods) {
        $warnings += $batchMethods
        $warnings = $warnings | Where-Object { 
            -not ($batchMethods | Where-Object { $_.File -eq $_.File -and $_.Line -eq $_.Line })
        }
        $warnings += $batchMethods
    }
}

# 输出结果
Write-Host ""
Write-Host "=== 扫描结果 ===" -ForegroundColor Cyan
Write-Host ""

if ($issues.Count -eq 0 -and $warnings.Count -eq 0) {
    Write-Host "✓ 未发现规范违反" -ForegroundColor Green
    exit 0
}

if ($issues.Count -gt 0) {
    Write-Host "发现 $($issues.Count) 个错误：" -ForegroundColor Red
    Write-Host ""
    
    $issues | Group-Object Category | ForEach-Object {
        Write-Host "  [$($_.Name)]" -ForegroundColor Yellow
        $_.Group | ForEach-Object {
            Write-Host "    $($_.File):$($_.Line) - $($_.Message)" -ForegroundColor Red
            Write-Host "      $($_.Code)" -ForegroundColor Gray
        }
        Write-Host ""
    }
}

if ($warnings.Count -gt 0) {
    Write-Host "发现 $($warnings.Count) 个警告：" -ForegroundColor Yellow
    Write-Host ""
    
    $warnings | Group-Object Category | ForEach-Object {
        Write-Host "  [$($_.Name)]" -ForegroundColor Yellow
        $_.Group | ForEach-Object {
            Write-Host "    $($_.File):$($_.Line) - $($_.Message)" -ForegroundColor Yellow
            if ($_.Code) {
                Write-Host "      $($_.Code)" -ForegroundColor Gray
            }
        }
        Write-Host ""
    }
}

# 生成报告文件
$reportPath = Join-Path $projectRoot "docs\code-review-scan-report.md"
$report = @"
# 代码审查扫描报告

**扫描时间**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**扫描范围**: $sourcePath

## 错误 ($($issues.Count))

$($issues | ForEach-Object { "- **$($_.File):$($_.Line)** - $($_.Message)`n  ```$($_.Code)```" } | Out-String)

## 警告 ($($warnings.Count))

$($warnings | ForEach-Object { "- **$($_.File):$($_.Line)** - $($_.Message)" + $(if ($_.Code) { "`n  ```$($_.Code)```" } else { "" }) } | Out-String)
"@

$report | Out-File -FilePath $reportPath -Encoding UTF8
Write-Host "报告已保存到: $reportPath" -ForegroundColor Cyan

# 返回退出代码
if ($issues.Count -gt 0) {
    exit 1
} else {
    exit 0
}

