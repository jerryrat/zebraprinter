# Access Database Monitor

Access数据库监控程序 - 实时监控TestRecord表的新增记录

## 功能特性

- 实时监控Access数据库中TestRecord表的新增记录
- 重点监控TR_SerialNum和TR_ID字段
- 用户友好的图形界面
- 可自定义监控间隔（1-60秒）
- 实时显示监控数据和日志信息
- 支持.mdb和.accdb格式的Access数据库

## 系统要求

- Windows 10 或更高版本
- Microsoft Access Database Engine (通常随Office安装)
- .NET 6.0 Runtime (如果使用独立发布版本则无需安装)

## 构建说明

### 使用PowerShell构建脚本

```powershell
cd AccessDatabaseMonitor
.\build.ps1
```

### 手动构建

```bash
cd AccessDatabaseMonitor
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ../Release
```

## 使用说明

1. 运行 `AccessDatabaseMonitor.exe`
2. 点击"选择数据库文件"按钮，选择要监控的Access数据库文件
3. 设置监控间隔（默认5秒）
4. 点击"开始监控"开始实时监控
5. 程序会在左侧面板显示检测到的新记录，右侧面板显示详细日志

## 程序结构

- `Program.cs` - 程序入口点
- `MainForm.cs` - 主窗口界面
- `DatabaseMonitor.cs` - 数据库监控逻辑
- `TestRecord.cs` - 数据模型
- `build.ps1` - 构建脚本

## 版本历史

- v1.0.0 - 初始版本，基本监控功能 