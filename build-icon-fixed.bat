@echo off
chcp 65001 > nul
echo ================================================
echo 太阳能电池测试打印监控系统 - 图标修复版构建
echo 版本: 1.1.21 (修复图标+Access兼容性)
echo 公司: ooitech
echo ================================================

REM 设置变量
set PROJECT_NAME=ZebraPrinterMonitor
set VERSION=1.1.21
set OUTPUT_DIR_X86=publish-x86
set OUTPUT_DIR_X64=publish-x64
set FINAL_DIR_X86=release-x86-v%VERSION%
set FINAL_DIR_X64=release-x64-v%VERSION%

echo.
echo 🔍 检测系统架构信息...
echo 当前系统: %PROCESSOR_ARCHITECTURE%
if defined PROCESSOR_ARCHITEW6432 (
    echo WOW64架构: %PROCESSOR_ARCHITEW6432%
)

echo.
echo ⚙️  构建选项:
echo    [1] 32位版本 (推荐) - 兼容性最佳，解决Access数据库问题
echo    [2] 64位版本 - 仅限64位Office环境
echo    [3] 两个版本都构建
echo    [4] 退出
echo.
set /p choice=请输入选择 (1-4): 

if "%choice%"=="1" goto build_x86
if "%choice%"=="2" goto build_x64
if "%choice%"=="3" goto build_both
if "%choice%"=="4" goto exit
echo 无效选择，默认构建32位版本...
goto build_x86

:build_x86
echo.
echo 📦 构建32位版本（推荐）...
call :clean_build
call :build_version win-x86 x86 %OUTPUT_DIR_X86% %FINAL_DIR_X86%
goto end

:build_x64
echo.
echo 📦 构建64位版本...
call :clean_build
call :build_version win-x64 x64 %OUTPUT_DIR_X64% %FINAL_DIR_X64%
goto end

:build_both
echo.
echo 📦 构建两个架构版本...
call :clean_build
call :build_version win-x86 x86 %OUTPUT_DIR_X86% %FINAL_DIR_X86%
call :build_version win-x64 x64 %OUTPUT_DIR_X64% %FINAL_DIR_X64%
goto end

:build_version
set RUNTIME_ID=%1
set ARCH_NAME=%2
set OUTPUT_DIR=%3
set FINAL_DIR=%4

echo.
echo 🏗️ 构建 %ARCH_NAME% 版本 (运行时: %RUNTIME_ID%)...

REM 清理输出目录
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
    echo ❌ %ARCH_NAME% 版本构建失败!
    pause
    exit /b 1
)

echo 创建发布目录...
mkdir "%FINAL_DIR%"

REM 复制主要文件
copy "%OUTPUT_DIR%\%PROJECT_NAME%.exe" "%FINAL_DIR%\"
copy "appsettings.json" "%FINAL_DIR%\"

REM 复制图标文件到发布目录（备用）
if exist "Zebra.ico" copy "Zebra.ico" "%FINAL_DIR%\"
if exist "zebra_icon.ico" copy "zebra_icon.ico" "%FINAL_DIR%\"

REM 创建详细说明文档
(
echo 太阳能电池测试打印监控系统 v%VERSION% - %ARCH_NAME% 版本
echo 开发公司: ooitech
echo 构建时间: %date% %time%
echo.
echo 🔧 本版本特色:
echo    ✅ 修复了系统托盘图标显示问题
echo    ✅ 图标现已嵌入到程序集中，确保正常显示
echo    ✅ 解决了Access数据库引擎兼容性问题
echo    ✅ 智能架构检测和错误诊断
echo.
echo 🔧 架构信息:
echo    • 此版本为 %ARCH_NAME% 架构
if "%ARCH_NAME%"=="x86" (
echo    • ✅ 推荐版本 - 兼容性最佳
echo    • ✅ 可在32位和64位系统上运行
echo    • ✅ 兼容32位Office/Access（最常见）
echo    • ✅ 解决了大部分数据库连接问题
) else (
echo    • ⚠️  仅适用于64位系统
echo    • ⚠️  需要64位Office/Access
echo    • ⚠️  如遇数据库连接问题请使用32位版本
)
echo.
echo 🚀 快速开始:
echo    1. 双击 %PROJECT_NAME%.exe 启动程序
echo    2. 程序会自动显示系统托盘图标
echo    3. 点击"浏览"选择Access数据库文件
echo    4. 配置打印机并开始监控
echo.
echo 🔍 如果遇到数据库连接问题:
echo    • 程序会显示详细的错误诊断信息
echo    • 根据提示下载对应架构的Access数据库引擎
echo    • 32位版本通常能解决大部分兼容性问题
echo.
echo 🎯 图标说明:
echo    • 系统托盘图标：程序最小化时显示
echo    • 窗体标题栏图标：主界面左上角显示
echo    • 图标已嵌入程序，无需外部文件
echo.
echo 📞 技术支持:
echo    • 公司: ooitech
echo    • 查看程序目录下的日志文件获取详细信息
echo    • 所有配置保存在 appsettings.json 文件中
) > "%FINAL_DIR%\使用说明-%ARCH_NAME%.txt"

REM 创建快速启动脚本
(
echo @echo off
echo chcp 65001 ^> nul
echo echo 启动太阳能电池测试监控系统 v%VERSION% ^(%ARCH_NAME%^)...
echo echo 开发公司: ooitech
echo echo.
echo start "" "%PROJECT_NAME%.exe"
echo echo ✅ 程序已启动，请查看系统托盘图标
echo timeout /t 3 ^> nul
) > "%FINAL_DIR%\启动程序.bat"

REM 计算文件大小
for %%A in ("%FINAL_DIR%\%PROJECT_NAME%.exe") do set SIZE=%%~zA
set /a SIZE_MB=%SIZE%/1024/1024

echo ✅ %ARCH_NAME% 版本构建完成: %FINAL_DIR%
echo 📁 文件大小: %SIZE_MB% MB
goto :eof

:clean_build
echo 🧹 清理构建目录...
if exist "bin" rmdir /s /q "bin"
if exist "obj" rmdir /s /q "obj"

echo 📦 恢复NuGet包...
dotnet restore
if errorlevel 1 (
    echo ❌ 恢复NuGet包失败!
    pause
    exit /b 1
)

echo 🔨 构建Release版本...
dotnet build --configuration Release --no-restore
if errorlevel 1 (
    echo ❌ 构建失败!
    pause
    exit /b 1
)
goto :eof

:end
echo.
echo 🎉 构建完成！
echo.
echo 💡 版本说明:
echo    • v1.1.21 - 图标修复版 + Access兼容性修复
echo    • 公司更新为: ooitech
echo    • 推荐使用32位版本（兼容性最佳）
echo.
echo 🔧 主要改进:
echo    ✅ 系统托盘图标正常显示
echo    ✅ 图标嵌入程序集，无需外部文件
echo    ✅ 解决Access数据库引擎兼容性问题
echo    ✅ 智能错误诊断和解决方案提示
echo.
pause

:exit 