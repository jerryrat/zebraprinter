# 斑马条形码打印机监控系统

一个用于监控Access数据库并自动打印条形码标签的Node.js应用程序。

## 功能特点

- 🔍 **实时监控**: 实时监控Access数据库中的新记录
- 🖨️ **自动打印**: 当检测到新记录时自动打印条形码标签
- 🦓 **斑马打印机支持**: 支持ZPL语言的斑马打印机
- 🌐 **Web界面**: 提供实时监控和管理的Web界面
- 📊 **统计信息**: 显示打印统计和系统状态
- 🔄 **手动打印**: 支持手动选择记录进行打印
- 📝 **实时日志**: 显示系统运行状态和错误信息

## 支持的条形码类型

- 标准条形码 (Code 39)
- Code 128
- QR码

## 系统要求

- Node.js 14.0 或更高版本
- Windows操作系统 (支持Access数据库)
- 斑马打印机 (支持ZPL语言)
- Microsoft Access Runtime 或 Office

## 安装步骤

1. **克隆项目**
   ```bash
   git clone <repository-url>
   cd zebra-barcode-printer
   ```

2. **安装依赖**
   ```bash
   npm install
   ```

3. **配置环境**
   - 复制 `.env.example` 为 `.env`
   - 修改配置文件中的参数

4. **设置数据库路径**
   - 将您的Access数据库文件放在 `./data/` 目录下
   - 或修改配置文件中的 `ACCESS_DB_PATH` 参数

## 配置说明

### 环境变量配置

在项目根目录创建 `.env` 文件：

```env
# 数据库配置
ACCESS_DB_PATH=./data/database.accdb
POLL_INTERVAL=1000
MONITOR_TABLE=Records
MONITOR_FIELD=id

# 打印机配置
PRINTER_NAME=ZDesigner GC420t
PRINTER_TYPE=RAW

# 应用配置
PORT=3000
LOG_LEVEL=info
LOG_FILE=./logs/app.log
```

### 参数说明

- `ACCESS_DB_PATH`: Access数据库文件路径
- `POLL_INTERVAL`: 数据库轮询间隔（毫秒）
- `MONITOR_TABLE`: 要监控的数据表名
- `MONITOR_FIELD`: 用于检测新记录的字段名（通常是ID字段）
- `PRINTER_NAME`: 打印机名称
- `PRINTER_TYPE`: 打印机类型 (RAW)
- `PORT`: Web界面端口号
- `LOG_LEVEL`: 日志级别 (error, warn, info, debug)

## 使用方法

### 启动应用

```bash
# 开发模式
npm run dev

# 生产模式
npm start
```

### 访问Web界面

在浏览器中打开 `http://localhost:3000`

### 数据库表结构

确保您的Access数据库表包含以下字段：

```sql
CREATE TABLE Records (
    id INTEGER PRIMARY KEY,
    name TEXT,
    barcode TEXT,
    description TEXT
);
```

### API接口

#### 获取系统状态
```
GET /api/status
```

#### 获取所有记录
```
GET /api/records
```

#### 获取指定记录
```
GET /api/records/:id
```

#### 手动打印记录
```
POST /api/print/:id
Content-Type: application/json

{
    "barcodeType": "standard" // 可选: standard, code128, qr
}
```

#### 测试打印
```
POST /api/test-print
```

#### 控制监控
```
POST /api/monitoring/start
POST /api/monitoring/stop
```

#### 获取可用打印机
```
GET /api/printers
```

## 打印机设置

### 支持的打印机

- Zebra GC420t
- Zebra GK420t
- Zebra ZT410
- 其他支持ZPL的斑马打印机

### 打印机配置

1. 确保打印机已正确安装驱动
2. 设置打印机为默认打印机或在配置中指定名称
3. 确保打印机支持RAW数据类型

## 故障排除

### 常见问题

1. **数据库连接失败**
   - 检查Access数据库文件路径
   - 确保已安装Access Runtime
   - 检查数据库文件权限

2. **打印机不可用**
   - 检查打印机是否已连接
   - 确认打印机名称配置正确
   - 检查打印机驱动是否正确安装

3. **监控不工作**
   - 检查数据库表名和字段名配置
   - 确保数据库文件未被其他程序锁定
   - 检查轮询间隔设置

### 日志查看

日志文件位于 `./logs/app.log`

```bash
# 实时查看日志
tail -f logs/app.log
```

## 开发说明

### 项目结构

```
zebra-barcode-printer/
├── src/
│   ├── config/
│   │   └── config.js          # 配置文件
│   ├── services/
│   │   ├── AccessMonitor.js   # 数据库监控服务
│   │   └── PrinterService.js  # 打印机服务
│   ├── utils/
│   │   └── logger.js          # 日志工具
│   └── index.js               # 主入口文件
├── public/
│   └── index.html             # Web界面
├── logs/                      # 日志目录
├── data/                      # 数据库文件目录
├── package.json
└── README.md
```

### 扩展功能

1. **添加新的条形码类型**
   - 修改 `PrinterService.js` 中的ZPL模板

2. **支持其他数据库**
   - 创建新的数据库监控服务
   - 实现相同的接口

3. **添加邮件通知**
   - 在打印成功/失败时发送邮件

## 许可证

MIT License

## 支持

如有问题或建议，请提交Issue或联系开发者。 