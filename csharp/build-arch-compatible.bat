@echo off
chcp 65001 > nul
echo ================================================
echo 太阳能电池测试打印监控系统 - 架构兼容构建
echo 版本: 1.1.20 (解决Access数据库引擎兼容性问题)
echo ================================================

REM 设置变量
set PROJECT_NAME=ZebraPrinterMonitor
set VERSION=1.1.20

echo.
echo 🔍 检测系统架构信息...
echo 当前系统: %PROCESSOR_ARCHITECTURE%
if defined PROCESSOR_ARCHITEW6432 (
    echo WOW64架构: %PROCESSOR_ARCHITEW6432%
)

REM 检测Office安装情况
echo.
echo 🔍 检测Office/Access安装情况...
reg query "HKLM\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorlevel% == 0 (
    echo 检测到Office安装 - 检查架构...
    reg query "HKLM\SOFTWARE\Microsoft\Office" /s | findstr /i "Bitness" >nul 2>&1
    if %errorlevel% == 0 (
        echo Office架构信息:
        reg query "HKLM\SOFTWARE\Microsoft\Office" /s | findstr /i "Bitness"
    )
) else (
    echo 未检测到Office安装
)

REM 检测ACE驱动
echo.
echo 🔍 检测Access数据库引擎...
reg query "HKLM\SOFTWARE\Microsoft\Office\*\Access Connectivity Engine" >nul 2>&1
if %errorlevel% == 0 (
    echo 检测到Access数据库引擎
) else (
    reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\*\Access Connectivity Engine" >nul 2>&1
    if %errorlevel% == 0 (
        echo 检测到32位Access数据库引擎（WOW64）
    ) else (
        echo 未检测到Access数据库引擎
    )
)

echo.
echo ⚙️  选择构建架构:
echo    [1] 32位版本 (x86) - 兼容32位Office/Access
echo    [2] 64位版本 (x64) - 需要64位Office/Access  
echo    [3] 两个版本都构建
echo    [4] 退出
echo.
set /p choice=请输入选择 (1-4): 

if "%choice%"=="1" goto build_x86
if "%choice%"=="2" goto build_x64
if "%choice%"=="3" goto build_both
if "%choice%"=="4" goto exit
echo 无效选择！
pause
goto exit

:build_x86
echo.
echo 📦 构建32位版本...
call :clean_build
call :build_version win-x86 x86
goto end

:build_x64
echo.
echo 📦 构建64位版本...
call :clean_build
call :build_version win-x64 x64
goto end

:build_both
echo.
echo 📦 构建两个架构版本...
call :clean_build
call :build_version win-x86 x86
call :build_version win-x64 x64
goto end

:build_version
set RUNTIME_ID=%1
set ARCH_NAME=%2
set OUTPUT_DIR=publish-%ARCH_NAME%
set FINAL_DIR=release-%ARCH_NAME%

echo.
echo 构建 %ARCH_NAME% 版本 (运行时: %RUNTIME_ID%)...

if exist "%OUTPUT_DIR%" rmdir /s /q "%OUTPUT_DIR%"
if exist "%FINAL_DIR%" rmdir /s /q "%FINAL_DIR%"

echo 发布单文件版本...
dotnet publish ^
    --configuration Release ^
    --runtime %RUNTIME_ID% ^
    --self-contained true ^
    --output "%OUTPUT_DIR%" ^
    /p:PublishSingleFile=true ^
    /p:IncludeNativeLibrariesForSelfExtract=true ^
    /p:EnableCompressionInSingleFile=true ^
    /p:DebugType=None ^
    /p:DebugSymbols=false ^
    /p:PublishTrimmed=false

if errorlevel 1 (
    echo %ARCH_NAME% 版本构建失败!
    pause
    exit /b 1
)

echo 创建发布目录...
mkdir "%FINAL_DIR%"

REM 复制主要文件
copy "%OUTPUT_DIR%\%PROJECT_NAME%.exe" "%FINAL_DIR%\"
copy "appsettings.json" "%FINAL_DIR%\"

REM 创建版本说明
(
echo 太阳能电池测试打印监控系统 v%VERSION% - %ARCH_NAME% 版本
echo.
echo 🔧 架构兼容性说明:
echo    • 此版本为 %ARCH_NAME% 架构
if "%ARCH_NAME%"=="x86" (
echo    • 兼容32位和64位系统
echo    • 需要32位Access数据库引擎
echo    • 推荐用于已安装32位Office的系统
) else (
echo    • 只能在64位系统上运行
echo    • 需要64位Access数据库引擎
echo    • 推荐用于已安装64位Office的系统
)
echo.
echo 📋 使用说明:
echo    1. 双击 %PROJECT_NAME%.exe 启动程序
echo    2. 如果出现数据库连接错误，请根据错误提示安装对应的Access数据库引擎
echo    3. 配置数据库路径和打印机设置
echo    4. 开始监控数据库变化
echo.
echo 🔗 获取Access数据库引擎:
echo    https://www.microsoft.com/zh-cn/download/details.aspx?id=54920
echo.
echo 构建时间: %date% %time%
echo 构建版本: %VERSION%
) > "%FINAL_DIR%\README-%ARCH_NAME%.txt"

echo ✅ %ARCH_NAME% 版本构建完成: %FINAL_DIR%
goto :eof

:clean_build
echo 清理构建目录...
if exist "bin" rmdir /s /q "bin"
if exist "obj" rmdir /s /q "obj"

echo 恢复NuGet包...
dotnet restore
if errorlevel 1 (
    echo 恢复NuGet包失败!
    pause
    exit /b 1
)

echo 构建Release版本...
dotnet build --configuration Release --no-restore
if errorlevel 1 (
    echo 构建失败!
    pause
    exit /b 1
)
goto :eof

:end
echo.
echo 🎉 构建完成！请查看相应的release目录。
echo.
echo 💡 使用建议:
echo    • 如果系统安装了32位Office，使用32位版本
echo    • 如果系统安装了64位Office，使用64位版本
echo    • 如果不确定，可以先尝试32位版本（兼容性更好）
echo.
pause

:exit 