@echo off
chcp 65001 >nul
echo ========================================
echo PPA 插件发布脚本
echo ========================================
echo.

REM 检查是否在项目根目录
if not exist "PPA\PPA.csproj" (
    echo 错误：未找到项目文件。请在项目根目录运行此脚本。
    pause
    exit /b 1
)

REM 查找 MSBuild
set MSBUILD_PATH=
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
) else (
    echo 错误：未找到 MSBuild。请确保已安装 Visual Studio。
    pause
    exit /b 1
)

echo 使用 MSBuild: "%MSBUILD_PATH%"
echo.

REM 绝对发布目录（避免引号参与到 MSBuild 的属性值中）
set "PUBLISH_DIR=%CD%\publish"
if not exist "%PUBLISH_DIR%" mkdir "%PUBLISH_DIR%" >nul 2>&1

echo 发布目录: "%PUBLISH_DIR%"

echo [1/4] 清理之前的构建...
call "%MSBUILD_PATH%" "PPA\PPA.sln" /t:Clean /p:Configuration=Release /p:Platform="Any CPU" /nologo /verbosity:minimal
if errorlevel 1 (
    echo 警告：清理失败，继续执行...
)

echo [2/4] 构建项目 (Release)...
call "%MSBUILD_PATH%" "PPA\PPA.sln" /t:Build /p:Configuration=Release /p:Platform="Any CPU" /nologo /verbosity:minimal
if errorlevel 1 (
    echo 错误：构建失败！
    pause
    exit /b 1
)
echo ✓ 构建成功
echo.

echo [3/4] 发布 ClickOnce 安装程序...
REM 这里不要给 PublishDir 加整体引号，避免 MSBuild 将后续参数并入属性值
call "%MSBUILD_PATH%" "PPA\PPA.csproj" /t:Publish /p:Configuration=Release /p:Platform=AnyCPU /p:OutputPath=bin\Release\ /p:PublishDir=%PUBLISH_DIR%\ /nologo /verbosity:minimal
if errorlevel 1 (
    echo 错误：发布失败！
    pause
    exit /b 1
)
echo ✓ 发布成功
echo.

echo [4/4] 检查输出文件...
if exist "%PUBLISH_DIR%\setup.exe" (
    echo ✓ 找到安装程序: %PUBLISH_DIR%\setup.exe
) else (
    echo ⚠ 警告：未找到 setup.exe
)

set "APP_MANIFEST="
if exist "%PUBLISH_DIR%\PPA.application" set "APP_MANIFEST=%PUBLISH_DIR%\PPA.application"
if not defined APP_MANIFEST (
    for /r "%PUBLISH_DIR%" %%f in (*.application) do (
        if not defined APP_MANIFEST set "APP_MANIFEST=%%f"
    )
)
if defined APP_MANIFEST (
    echo ✓ 找到应用程序清单: %APP_MANIFEST%
) else (
    echo ⚠ 警告：未找到 PPA.application
)

echo.
echo ========================================
echo 发布完成！
echo ========================================
echo.
echo 发布文件位置: %PUBLISH_DIR%
echo.
echo 下一步：
echo 1. 检查发布文件夹中的文件
echo 2. 测试安装程序
echo 3. 将整个 publish 文件夹打包分发
echo.

set /p OPEN_FOLDER="是否打开发布文件夹？(Y/N): "
if /i "%OPEN_FOLDER%"=="Y" (
    start explorer "%PUBLISH_DIR%"
)

pause

