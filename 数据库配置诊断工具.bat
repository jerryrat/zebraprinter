@echo off
chcp 65001 >nul
echo ========================================
echo 📊 数据库配置诊断工具 v1.1.56
echo ========================================
echo.

echo 🔍 正在检查配置文件...
if not exist "appsettings.json" (
    echo ❌ 错误：配置文件 appsettings.json 不存在！
    pause
    exit /b 1
)

echo ✅ 配置文件存在
echo.

echo 📋 当前配置内容：
type appsettings.json | findstr /i "DatabasePath"
echo.

echo 🔧 问题诊断：
type appsettings.json | findstr /C:"\"DatabasePath\": \"\"" >nul
if %errorlevel% == 0 (
    echo ❌ 发现问题：数据库路径未配置（DatabasePath 为空）
    echo.
    echo 💡 解决方案：
    echo    1. 启动 ZebraPrinterMonitor.exe
    echo    2. 点击程序界面的"浏览"按钮
    echo    3. 选择您的 Access 数据库文件（.mdb 或 .accdb）
    echo    4. 点击"测试连接"确认数据库可以访问
    echo    5. 程序会自动保存配置并启动监控
    echo.
) else (
    echo ✅ 数据库路径已配置
    type appsettings.json | findstr /C:"DatabasePath" | findstr /v "Comments"
    echo.
    
    echo 🔍 检查数据库文件是否存在...
    for /f "tokens=2 delims=:" %%i in ('type appsettings.json ^| findstr /C:"DatabasePath" ^| findstr /v "Comments"') do (
        set "dbpath=%%i"
        set "dbpath=!dbpath: =!"
        set "dbpath=!dbpath:"=!"
        set "dbpath=!dbpath:,=!"
        if exist "!dbpath!" (
            echo ✅ 数据库文件存在: !dbpath!
        ) else (
            echo ❌ 数据库文件不存在: !dbpath!
            echo.
            echo 💡 请检查：
            echo    1. 文件路径是否正确
            echo    2. 文件是否被移动或删除
            echo    3. 网络驱动器是否连接正常
        )
    )
)

echo.
echo 📊 TestRecord 表结构检查：
echo    当增强监控启动时，程序会自动：
echo    1. 连接到指定的数据库
echo    2. 读取 TestRecord 表的所有列
echo    3. 监控任意字段的数据变化
echo    4. 实时刷新最近记录窗口

echo.
echo 🚀 监控功能特点：
echo    ⚡ 50ms 文件变化检测
echo    🔄 500ms 定时轮询检测
echo    📊 全字段监控支持
echo    🎯 智能重复检测避免
echo    📱 实时界面刷新

echo.
echo ========================================
echo 按任意键关闭...
pause >nul 