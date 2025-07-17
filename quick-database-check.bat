@echo off
echo ====================================
echo   数据库监控快速诊断工具 v1.1.56
echo ====================================
echo.

:: 检查配置文件
echo [步骤1] 检查配置文件...
if not exist "appsettings.json" (
    echo ❌ 配置文件 appsettings.json 不存在！
    echo.
    pause
    exit /b 1
)
echo ✅ 配置文件存在

:: 读取数据库路径
echo.
echo [步骤2] 检查数据库路径配置...
findstr "DatabasePath" appsettings.json > temp_path.txt
set /p db_line=<temp_path.txt
del temp_path.txt

echo 当前配置: %db_line%
echo %db_line% | findstr /C:"\"\"" >nul
if %errorlevel%==0 (
    echo ❌ 数据库路径为空！这是监控无法工作的主要原因。
    echo.
    echo 解决方案：
    echo 1. 启动 ZebraPrinterMonitor.exe
    echo 2. 在"数据库配置"区域点击"浏览"按钮
    echo 3. 选择您的 .mdb 或 .accdb 文件
    echo 4. 点击"测试连接"确认
    echo.
    pause
    exit /b 1
) else (
    echo ✅ 数据库路径已配置
)

:: 检查程序运行状态
echo.
echo [步骤3] 检查程序运行状态...
tasklist /FI "IMAGENAME eq ZebraPrinterMonitor.exe" 2>NUL | find /I /N "ZebraPrinterMonitor.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo ✅ 程序正在运行
) else (
    echo ⚠️  程序未运行，启动程序后监控才会开始
)

:: 检查文件
echo.
echo [步骤4] 检查相关文件...
if exist "ZebraPrinterMonitor.exe" (
    echo ✅ 主程序文件存在
) else (
    echo ❌ 主程序文件 ZebraPrinterMonitor.exe 不存在！
)

if exist "print_templates.json" (
    echo ✅ 打印模板文件存在
) else (
    echo ⚠️  打印模板文件不存在，将使用默认模板
)

:: 显示监控优化建议
echo.
echo [监控优化建议]
echo 📈 新版本监控改进：
echo    • 双重监控：文件监控 + 定时轮询
echo    • 即时响应：文件变化时立即检测新数据
echo    • 线程安全：避免并发冲突
echo    • 错误重试：提高稳定性
echo.
echo 📝 使用提示：
echo    1. 确保数据库文件不被其他程序独占锁定
echo    2. 建议轮询间隔保持1000ms（1秒）
echo    3. 文件监控会在数据变化时立即触发检测
echo    4. 查看程序日志了解详细监控状态
echo.

:: 显示版本信息
echo [版本信息]
echo 当前版本: v1.1.56 (数据库监控增强版)
echo 更新内容: 新增文件监控机制，提高数据检测响应速度
echo.

echo 诊断完成！
echo 如果监控仍有问题，请查看程序日志获取详细信息。
echo.
pause 