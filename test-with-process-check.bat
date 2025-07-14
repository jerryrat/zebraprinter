@echo off
chcp 65001 > nul
echo ================================================
echo 太阳能电池测试打印监控系统 - 测试启动
echo 版本: 1.1.21 (图标修复+架构兼容版)
echo 公司: ooitech
echo ================================================

REM 设置程序名称
set PROCESS_NAME=ZebraPrinterMonitor

echo.
echo 🔍 检测是否有运行中的程序实例...

REM 查找正在运行的进程
tasklist /FI "IMAGENAME eq %PROCESS_NAME%.exe" 2>NUL | find /I "%PROCESS_NAME%.exe" >NUL
if %errorlevel% == 0 (
    echo ⚠️  发现运行中的程序实例，正在关闭...
    
    REM 尝试优雅关闭
    taskkill /IM "%PROCESS_NAME%.exe" /F >NUL 2>&1
    if %errorlevel% == 0 (
        echo ✅ 程序实例已关闭
        timeout /t 2 /nobreak >NUL
    ) else (
        echo ❌ 无法关闭程序实例，请手动关闭后重试
        pause
        exit /b 1
    )
) else (
    echo ✅ 没有发现运行中的程序实例
)

echo.
echo 🏗️  正在构建测试版本...

REM 清理之前的构建
if exist "bin" rmdir /s /q "bin"
if exist "obj" rmdir /s /q "obj"

REM 恢复包
dotnet restore
if errorlevel 1 (
    echo ❌ NuGet包恢复失败!
    pause
    exit /b 1
)

REM 构建Debug版本进行测试
dotnet build --configuration Debug
if errorlevel 1 (
    echo ❌ 构建失败!
    pause
    exit /b 1
)

echo ✅ 构建完成

echo.
echo 🚀 启动测试...
echo 💡 本版本特色:
echo    ✅ 修复了系统托盘图标显示问题
echo    ✅ 图标已嵌入到程序集中，确保正常显示
echo    ✅ 解决了Access数据库引擎兼容性问题
echo    ✅ 自动检测应用程序架构(32位/64位)
echo    ✅ 智能选择合适的数据库驱动
echo    ✅ 提供详细的错误诊断信息
echo    ✅ 公司信息更新为: ooitech

echo.
echo 🔧 测试要点:
echo    1. 检查系统托盘图标是否正常显示
echo    2. 检查窗体标题栏图标是否正常显示
echo    3. 测试数据库连接（会显示架构检测信息）
echo    4. 验证错误诊断的准确性

echo.
echo 按任意键启动程序进行测试...
pause >NUL

REM 启动程序
dotnet run --configuration Debug

echo.
echo 🏁 程序已退出
echo.
echo 📋 如果测试通过，可以运行以下脚本进行正式构建:
echo    • build-icon-fixed.bat - 构建图标修复版
echo    • build-arch-compatible.bat - 构建架构兼容版
echo.
pause 