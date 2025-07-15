@echo off
chcp 65001 > nul
echo ================================================
echo 太阳能电池测试打印监控系统 - 单文件构建
echo 版本: 1.1.49 (C# WinForms版本)
echo ================================================

REM 设置变量
set PROJECT_NAME=ZebraPrinterMonitor
set VERSION=1.1.49
set OUTPUT_DIR=publish-single
set FINAL_DIR=release

REM 清理输出目录
echo 清理输出目录...
if exist "%OUTPUT_DIR%" rmdir /s /q "%OUTPUT_DIR%"
if exist "%FINAL_DIR%" rmdir /s /q "%FINAL_DIR%"
if exist "bin" rmdir /s /q "bin"
if exist "obj" rmdir /s /q "obj"

echo.
echo 恢复NuGet包...
dotnet restore
if errorlevel 1 (
    echo 恢复NuGet包失败!
    pause
    exit /b 1
)

echo.
echo 构建Release版本...
dotnet build --configuration Release --no-restore
if errorlevel 1 (
    echo 构建失败!
    pause
    exit /b 1
)

echo.
echo 发布单文件版本（绿色可执行文件）...
dotnet publish ^
    --configuration Release ^
    --runtime win-x64 ^
    --self-contained true ^
    --output "%OUTPUT_DIR%" ^
    /p:PublishSingleFile=true ^
    /p:IncludeNativeLibrariesForSelfExtract=true ^
    /p:EnableCompressionInSingleFile=true ^
    /p:DebugType=None ^
    /p:DebugSymbols=false ^
    /p:PublishTrimmed=false

if errorlevel 1 (
    echo 发布失败!
    pause
    exit /b 1
)

echo.
echo 创建最终发布目录...
mkdir "%FINAL_DIR%"

REM 复制主要文件
copy "%OUTPUT_DIR%\%PROJECT_NAME%.exe" "%FINAL_DIR%\"
copy "appsettings.json" "%FINAL_DIR%\"

REM 创建示例数据库文件（如果存在）
if exist "..\data\database_access.mdb" (
    mkdir "%FINAL_DIR%\sample"
    copy "..\data\database_access.mdb" "%FINAL_DIR%\sample\sample_database.mdb"
    echo 已复制示例数据库文件
)

REM 创建使用说明
echo 创建使用说明...
(
echo 太阳能电池测试打印监控系统 v%VERSION%
echo.
echo 这是一个绿色单文件版本，无需安装即可运行。
echo.
echo 快速开始:
echo 1. 直接双击 %PROJECT_NAME%.exe 运行程序
echo 2. 在"数据监控"页面点击"浏览"选择Access数据库文件
echo 3. 在"系统配置"页面选择打印机和打印格式  
echo 4. 点击"开始监控"即可自动监控并打印
echo.
echo 系统要求:
echo - Windows 10 或更高版本
echo - Microsoft Access 数据库引擎（用于.mdb文件访问）
echo.
echo 功能特性:
echo - 实时监控Access数据库中的TestRecord表
echo - 自动检测新的测试记录并打印标签
echo - 支持多种打印格式：文本、ZPL、Code128条码、QR码
echo - 系统托盘最小化运行
echo - 完整的配置管理和日志记录
echo.
echo 技术支持:
echo 如遇问题请查看logs目录下的日志文件
echo.
echo 构建时间: %date% %time%
) > "%FINAL_DIR%\使用说明.txt"

REM 创建快速启动脚本
(
echo @echo off
echo echo 启动太阳能电池测试监控系统...
echo start "" "%PROJECT_NAME%.exe"
echo echo 程序已启动，请查看桌面图标
echo timeout /t 3 > nul
) > "%FINAL_DIR%\启动程序.bat"

REM 计算文件大小
for %%A in ("%FINAL_DIR%\%PROJECT_NAME%.exe") do set SIZE=%%~zA
set /a SIZE_MB=%SIZE%/1024/1024

echo.
echo ================================================
echo 单文件构建完成！
echo ================================================
echo 输出目录: %FINAL_DIR%\
echo 主程序: %PROJECT_NAME%.exe (%SIZE_MB% MB)
echo 配置文件: appsettings.json
echo 使用说明: 使用说明.txt
echo.
echo 这是一个绿色单文件版本，特点：
echo ✓ 免安装运行，直接双击exe文件即可
echo ✓ 自包含所有依赖，无需安装.NET运行时
echo ✓ 单文件部署，便于分发和使用
echo ✓ 程序数据和日志保存在exe同目录
echo.
echo 您现在可以：
echo 1. 直接运行 %FINAL_DIR%\%PROJECT_NAME%.exe
echo 2. 将整个 %FINAL_DIR% 文件夹复制到任何Windows电脑上使用
echo 3. 或者只复制 %PROJECT_NAME%.exe 文件（最简方式）
echo ================================================

pause 