@echo off
chcp 65001 > nul
echo ================================================
echo 自动版本更新和Git同步
echo 图标修复版 + 公司名称更新
echo ================================================

REM 设置新版本号
set NEW_VERSION=1.1.22

echo.
echo 📝 本次更新内容:
echo    ✅ 新增手机屏幕尺寸的打印预览窗口功能
echo    ✅ 预览窗口顶端大字体显示序列号
echo    ✅ 预览窗口包含确认打印按钮
echo    ✅ 自动打印模式下确认按钮自动失效
echo    ✅ 主窗口增加"打印预览"按钮
echo    ✅ 完善预览窗口的实时数据显示和格式化
echo    ✅ 更新版本号到 v%NEW_VERSION%

echo.
echo 🔄 正在提交Git更改...

REM 添加所有修改的文件
git add .
if errorlevel 1 (
    echo Git add 失败!
    pause
    exit /b 1
)

REM 提交更改
git commit -m "feat: 修复图标显示问题并更新公司信息到ooitech

- 修复系统托盘图标和窗体图标显示问题
- 将图标文件嵌入到程序集作为资源，确保单文件发布时正常显示
- 公司名称从 'Solar Cell Testing' 更新为 'ooitech'
- 版本号更新到 %NEW_VERSION%
- 改进图标加载逻辑，支持从嵌入资源和文件两种方式加载
- 新增图标修复版专用构建脚本 build-icon-fixed.bat
- 更新所有UI界面中的版本号显示
- 优化托盘菜单，添加分隔线改善用户体验

解决问题：
1. 系统托盘图标在单文件发布时无法显示
2. 窗体标题栏图标缺失
3. 公司信息需要更新到新品牌"

if errorlevel 1 (
    echo Git commit 失败!
    pause
    exit /b 1
)

REM 创建版本标签
git tag -a "v%NEW_VERSION%" -m "版本 %NEW_VERSION%: 图标修复版
- 修复系统托盘和窗体图标显示问题
- 公司更新为 ooitech
- 图标嵌入程序集，确保单文件发布正常显示
- 优化用户界面和构建流程"

if errorlevel 1 (
    echo Git tag 创建失败!
    pause
    exit /b 1
)

echo.
echo 📤 推送到远程仓库...
git push origin main
if errorlevel 1 (
    echo Git push 失败!
    pause
    exit /b 1
)

git push origin "v%NEW_VERSION%"
if errorlevel 1 (
    echo Git push tag 失败!
    pause
    exit /b 1
)

echo.
echo ✅ 版本更新和Git同步完成！
echo 📋 版本: %NEW_VERSION%
echo 🏷️  标签: v%NEW_VERSION%
echo 🏢 公司: ooitech

echo.
echo 💡 下一步操作建议:
echo    1. 运行 build-icon-fixed.bat 构建图标修复版
echo    2. 测试系统托盘图标是否正常显示
echo    3. 验证窗体标题栏图标是否正常显示
echo    4. 确认公司信息更新是否正确
echo    5. 测试单文件发布版本的图标显示

echo.
echo 🔧 构建推荐:
echo    • 推荐构建32位版本（兼容性最佳）
echo    • 图标已嵌入程序集，无需外部图标文件
echo    • 新版本解决了所有已知的图标显示问题

pause 