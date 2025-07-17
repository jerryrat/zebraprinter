@echo off
echo ========================================
echo 斑马打印监控 - 快速诊断工具
echo ========================================
echo.

echo 检查配置文件...
if exist "appsettings.json" (
    echo ✅ 配置文件存在
    type appsettings.json
) else (
    echo ❌ 配置文件不存在
)

echo.
echo ========================================
echo 检查程序文件...
if exist "ZebraPrinterMonitor.exe" (
    echo ✅ 程序文件存在
) else (
    echo ❌ 程序文件不存在
)

echo.
echo ========================================
echo 当前问题分析：
echo 1. 检查上面显示的配置文件中 DatabasePath 是否为空
echo 2. 如果为空，请按以下步骤解决：
echo    - 启动程序
echo    - 点击"浏览"按钮选择数据库文件
echo    - 点击"测试连接"确认
echo    - 监控会自动启动
echo.
echo 或者直接编辑 appsettings.json 文件，
echo 在 DatabasePath 中设置你的数据库文件路径
echo.
pause 