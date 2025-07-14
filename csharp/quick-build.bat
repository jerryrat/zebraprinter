@echo off
echo 快速构建太阳能电池测试监控系统...

REM 清理
if exist "bin" rmdir /s /q "bin"
if exist "obj" rmdir /s /q "obj"

REM 恢复包
dotnet restore

REM 构建
dotnet build --configuration Release

echo.
echo 构建完成！可以运行以下命令测试：
echo dotnet run --configuration Release
echo.
echo 或者运行单文件构建：
echo build-single-file.bat
echo.
pause 