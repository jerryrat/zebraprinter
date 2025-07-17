using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;
using Microsoft.Win32;
using System.IO; // 添加用于FileSystemWatcher
using System.Threading.Tasks; // 添加用于Task.Delay
using System.Threading; // 添加用于Threading.Timer

namespace ZebraPrinterMonitor.Services
{
    public class DatabaseMonitor : IDisposable
    {
        // 按照AccessDatabaseMonitor的字段定义
        private string _connectionString = "";
        private readonly System.Threading.Timer _monitorTimer;
        private readonly HashSet<TestRecord> _knownRecords;
        private bool _isRunning;
        
        // 保留兼容性字段
        private string _currentTableName = "";
        private string _lastSerialNum = "";
        private int _retryCount = 0;
        private const int MaxRetries = 5;
        
        // 🔧 新增：监控周期计数器，用于减少日志频率
        private int _monitoringCycleCount = 0;

        // 按照AccessDatabaseMonitor的事件定义
        public event Action<List<TestRecord>>? NewRecordsDetected;
        public event Action<string>? ErrorOccurred;
        
        // 保留兼容性事件
        public event EventHandler<TestRecord>? NewRecordFound;
        public event EventHandler<string>? MonitoringError;
        public event EventHandler<string>? StatusChanged;
        
        // 🔧 新增：统一数据更新事件 - 基于GetLastRecord的完整数据刷新
        public event EventHandler<DataUpdateEventArgs>? DataUpdated;

        public bool IsMonitoring => _isRunning;
        public string LastSerialNum => _lastSerialNum;

        // 新增：存储表结构信息
        private List<string> _tableColumns = new List<string>();
        public List<string> TableColumns => _tableColumns.ToList();

        public DatabaseMonitor()
        {
            // 按照AccessDatabaseMonitor的方式初始化
            _knownRecords = new HashSet<TestRecord>();
            // 🔧 修复：使用基于GetLastRecord的简化监控逻辑
            _monitorTimer = new System.Threading.Timer(CheckForLastRecordChanges, null, Timeout.Infinite, Timeout.Infinite);
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor项目的方式实现数据库连接（确保与诊断一致）
        /// </summary>
        public async Task<bool> ConnectAsync(string databasePath, string tableName = "TestRecord")
        {
            try
            {
                Logger.Info($"🔗 开始数据库连接: {databasePath}");
                
                if (string.IsNullOrEmpty(databasePath))
                {
                    Logger.Error("❌ 数据库路径为空！");
                    return false;
                }
                
                if (!System.IO.File.Exists(databasePath))
                {
                    Logger.Error($"❌ 数据库文件不存在: {databasePath}");
                    return false;
                }
                
                // 使用与AccessDatabaseMonitor完全相同的连接字符串
                _connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Mode=Share Deny None;Persist Security Info=false;Jet OLEDB:Database Locking Mode=1;";
                _currentTableName = tableName;
                
                Logger.Info($"🔗 使用连接字符串: {_connectionString}");
                
                // 测试连接（与AccessDatabaseMonitor方式一致）
                using var connection = new OleDbConnection(_connectionString);
                await connection.OpenAsync();
                
                // 验证表存在
                using var command = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", connection);
                var count = await command.ExecuteScalarAsync();
                
                Logger.Info($"✅ 数据库连接成功！表 [{tableName}] 记录数: {count}");
                
                // 初始化已知记录（按照AccessDatabaseMonitor方式）
                await InitializeKnownRecordsAsync();
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 数据库连接失败: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// 旧的同步连接方法 - 已禁用，请使用ConnectAsync
        /// </summary>
        [Obsolete("此方法已禁用，请使用ConnectAsync方法以确保连接一致性")]
        public bool Connect(string databasePath, string tableName)
        {
            Logger.Error("❌ Connect方法已禁用！请使用ConnectAsync方法以确保与诊断的连接一致性");
            return false;
        }

        private static (string Provider, string ConnectionString, string Architecture)[] GetConnectionAttempts(
            string databasePath, bool isApp64Bit, bool isOS64Bit)
        {
            var attempts = new List<(string Provider, string ConnectionString, string Architecture)>();
            
            // Access数据库并发访问优化的连接字符串参数
            string concurrentParams = "Mode=Share Deny None;Persist Security Info=false;Jet OLEDB:Database Locking Mode=1;";

            if (isApp64Bit)
            {
                // 64位应用程序 - 优先使用64位驱动，优先尝试并发模式
                attempts.AddRange(new[]
                {
                    ("Microsoft.ACE.OLEDB.16.0", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};{concurrentParams}", "64位-并发"),
                    ("Microsoft.ACE.OLEDB.12.0", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};{concurrentParams}", "64位-并发"),
                    ("Microsoft.ACE.OLEDB.16.0", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};", "64位-标准"),
                    ("Microsoft.ACE.OLEDB.12.0", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};", "64位-标准"),
                });

                // 如果是64位系统，可能还安装了32位Office，但64位应用无法直接使用
                if (isOS64Bit)
                {
                    Logger.Info("64位应用程序无法使用32位Office驱动，跳过32位驱动测试");
                }
            }
            else
            {
                // 32位应用程序 - 可以使用32位驱动，优先尝试并发模式
                attempts.AddRange(new[]
                {
                    ("Microsoft.ACE.OLEDB.16.0", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};{concurrentParams}", "32位-并发"),
                    ("Microsoft.ACE.OLEDB.12.0", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};{concurrentParams}", "32位-并发"),
                    ("Microsoft.Jet.OLEDB.4.0", $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath};{concurrentParams}Jet OLEDB:Engine Type=5;", "32位-并发"),
                    ("Microsoft.ACE.OLEDB.16.0", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};", "32位-标准"),
                    ("Microsoft.ACE.OLEDB.12.0", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};", "32位-标准"),
                    ("Microsoft.Jet.OLEDB.4.0", $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath};", "32位-标准")
                });
            }

            return attempts.ToArray();
        }

        private static string GenerateDetailedErrorMessage(Exception lastException, bool isApp64Bit, bool isOS64Bit)
        {
            string baseError = lastException?.Message ?? "未知错误";
            var errorBuilder = new System.Text.StringBuilder();

            errorBuilder.AppendLine("数据库连接失败：Access数据库引擎架构不匹配！");
            errorBuilder.AppendLine();
            errorBuilder.AppendLine("🔍 问题诊断：");
            errorBuilder.AppendLine($"   • 当前应用程序：{(!isApp64Bit ? "32位" : "64位")}");
            errorBuilder.AppendLine($"   • 操作系统：{(!isOS64Bit ? "32位" : "64位")}");
            errorBuilder.AppendLine();

            if (isApp64Bit)
            {
                errorBuilder.AppendLine("❌ 64位应用程序需要64位Access数据库引擎");
                errorBuilder.AppendLine();
                errorBuilder.AppendLine("✅ 解决方案（任选其一）：");
                errorBuilder.AppendLine("   1. 安装64位Microsoft Access数据库引擎：");
                errorBuilder.AppendLine("      下载地址：https://www.microsoft.com/zh-cn/download/details.aspx?id=54920");
                errorBuilder.AppendLine("      注意：选择 'AccessDatabaseEngine_X64.exe'");
                errorBuilder.AppendLine();
                errorBuilder.AppendLine("   2. 或者重新编译应用程序为32位版本");
                errorBuilder.AppendLine("      (如果系统已安装32位Office)");
            }
            else
            {
                errorBuilder.AppendLine("❌ 32位应用程序需要32位Access数据库引擎");
                errorBuilder.AppendLine();
                errorBuilder.AppendLine("✅ 解决方案：");
                errorBuilder.AppendLine("   安装32位Microsoft Access数据库引擎：");
                errorBuilder.AppendLine("   下载地址：https://www.microsoft.com/zh-cn/download/details.aspx?id=54920");
                errorBuilder.AppendLine("   注意：选择 'AccessDatabaseEngine.exe' (32位版本)");
            }

            errorBuilder.AppendLine();
            errorBuilder.AppendLine("📝 安装说明：");
            errorBuilder.AppendLine("   1. 如果已安装Office，可能需要使用 /passive 参数强制安装");
            errorBuilder.AppendLine("   2. 安装完成后重启计算机");
            errorBuilder.AppendLine("   3. 重新启动此程序");
            errorBuilder.AppendLine();
            errorBuilder.AppendLine($"📋 详细错误：{baseError}");

            return errorBuilder.ToString();
        }

        public void StartMonitoring(int pollInterval = 1000)
        {
            if (_isRunning)
            {
                Logger.Warning("监控已在运行中");
                StatusChanged?.Invoke(this, "⚠️ 监控已在运行中");
                return;
            }

            if (string.IsNullOrEmpty(_connectionString))
            {
                Logger.Error("数据库未连接，无法开始监控");
                MonitoringError?.Invoke(this, "数据库未连接");
                StatusChanged?.Invoke(this, "❌ 数据库未连接，无法开始监控");
                return;
            }

            Logger.Info($"🚀 开始基于GetLastRecord的简化监控，每{pollInterval}ms检查一次");
            Logger.Info($"📊 监控表: {_currentTableName}");
            
            StatusChanged?.Invoke(this, $"🚀 启动GetLastRecord监控 - 表:{_currentTableName}, 间隔:{pollInterval}ms");

            try
            {
                // 🔧 简化：重置监控状态，无需复杂的记录集合初始化
                _lastKnownRecord = null; // 重置最后已知记录
                _monitoringCycleCount = 0; // 重置监控周期计数
                
                Logger.Info("🔍 准备初始化最后记录基线...");
                StatusChanged?.Invoke(this, "🔍 正在初始化最后记录基线...");
                
                // 启动监控定时器
                _monitorTimer.Change(0, pollInterval);
                _isRunning = true;
                _retryCount = 0;
                
                Logger.Info($"🚀 基于GetLastRecord的监控已成功启动！");
                StatusChanged?.Invoke(this, $"🚀 监控已启动！每{pollInterval}ms检查最后记录变化");
                
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 启动监控失败: {ex.Message}", ex);
                MonitoringError?.Invoke(this, ex.Message);
                StatusChanged?.Invoke(this, $"❌ 启动监控失败: {ex.Message}");
                _isRunning = false;
            }
        }

        // 强制刷新检查（用户手动触发）
        public void ForceRefresh()
        {
            if (!_isRunning)
            {
                Logger.Warning("监控未启动，无法执行强制刷新");
                StatusChanged?.Invoke(this, "⚠️ 监控未启动，无法强制刷新");
                return;
            }

            Logger.Info("🔄 用户触发强制刷新检查");
            StatusChanged?.Invoke(this, "🔄 用户触发强制刷新...");
            
            Task.Run(() =>
            {
                try
                {
                    // 🛠️ 修复：重置重试计数和连接状态
                    _retryCount = 0;
                    
                    // 🛠️ 修复：检查连接健康状况
                    if (!IsConnectionHealthy())
                    {
                        Logger.Warning("🔧 强制刷新时检测到连接问题，尝试重新连接...");
                        StatusChanged?.Invoke(this, "🔧 强制刷新：重新建立连接...");
                        
                        if (!AttemptReconnection())
                        {
                            Logger.Error("❌ 强制刷新失败：无法重建连接");
                            StatusChanged?.Invoke(this, "❌ 强制刷新失败：连接异常");
                            return;
                        }
                    }
                    
                    // 🛠️ 修复：调用新的监控检查方法
                    CheckForLastRecordChanges(null);
                    Logger.Info("✅ 强制刷新完成");
                    StatusChanged?.Invoke(this, "✅ 强制刷新完成");
                }
                catch (Exception ex)
                {
                    Logger.Error($"❌ 强制刷新检查失败: {ex.Message}", ex);
                    StatusChanged?.Invoke(this, $"❌ 强制刷新失败: {ex.Message}");
                }
            });
        }
        
        /// <summary>
        /// 强制刷新数据库连接 - 确保获取最新数据
        /// </summary>
        public void ForceRefreshConnection()
        {
            try
            {
                Logger.Info("🔄 强制刷新数据库连接以获取最新数据");
                
                // 🛠️ 修复：重置重试计数
                _retryCount = 0;
                
                // 🔧 增强实现：创建新的数据库连接来确保获取最新数据
                // Access数据库的特性需要新连接才能看到其他连接的最新更改
                if (!string.IsNullOrEmpty(_connectionString))
                {
                    using var testConnection = new System.Data.OleDb.OleDbConnection(_connectionString);
                    testConnection.Open();
                    
                    // 执行一个简单查询来确保连接活跃并同步数据
                    var testQuery = $"SELECT COUNT(*) FROM [{_currentTableName}]";
                    using var testCommand = new System.Data.OleDb.OleDbCommand(testQuery, testConnection);
                    var count = testCommand.ExecuteScalar();
                    
                    Logger.Info($"✅ 数据库连接刷新完成，当前表记录数: {count}");
                    
                    // 🛠️ 新增：刷新后立即触发一次监控检查
                    if (_isRunning)
                    {
                        Logger.Info("🔄 连接刷新后立即检查数据更新...");
                        Task.Run(() => CheckForLastRecordChanges(null));
                    }
                }
                else
                {
                    Logger.Warning("⚠️ 连接字符串为空，跳过连接刷新");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 强制刷新数据库连接失败: {ex.Message}", ex);
                
                // 🛠️ 新增：连接刷新失败时尝试重连
                if (_isRunning && !string.IsNullOrEmpty(_connectionString))
                {
                    Logger.Info("🔧 尝试重新建立连接...");
                    AttemptReconnection();
                }
            }
        }

        public void StopMonitoring()
        {
            if (!_isRunning) return;

            _monitorTimer.Change(Timeout.Infinite, 0);
            _isRunning = false;

            Logger.Info("TR_SerialNum监控已停止");
            StatusChanged?.Invoke(this, "监控已停止");
        }

        // 按照AccessDatabaseMonitor方式初始化已知记录基线
        private void InitializeKnownRecords()
        {
            var tableName = !string.IsNullOrEmpty(_currentTableName) ? _currentTableName : "TestRecord";
            
            Logger.Info($"🔍 初始化已知记录基线: 表={tableName}");

            try
            {
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();
                Logger.Info("✅ 数据库连接成功");

                // 按照AccessDatabaseMonitor的方式获取所有记录
                var query = $"SELECT TR_SerialNum, TR_ID FROM [{tableName}]";
                
                using var command = new OleDbCommand(query, connection);
                using var reader = command.ExecuteReader();
                
                var currentRecords = new List<TestRecord>();
                while (reader.Read())
                {
                    var serialNum = reader.IsDBNull("TR_SerialNum") ? string.Empty : reader.GetString("TR_SerialNum");
                    var id = reader.IsDBNull("TR_ID") ? string.Empty : reader.GetString("TR_ID");

                    currentRecords.Add(new TestRecord
                    {
                        TR_SerialNum = serialNum,
                        TR_ID = id
                    });
                }
                reader.Close();

                lock (_knownRecords)
                {
                    _knownRecords.Clear();
                    foreach (var record in currentRecords)
                    {
                        _knownRecords.Add(record);
                    }
                }

                Logger.Info($"🏁 已知记录基线初始化完成，共 {currentRecords.Count} 条记录");
                
                // 🔧 修复：使用TR_SerialNum排序而不是TR_ID来获取最新记录
                if (currentRecords.Count > 0)
                {
                    var latestRecord = currentRecords
                        .Where(r => !string.IsNullOrEmpty(r.TR_SerialNum))
                        .OrderByDescending(r => r.TR_SerialNum)
                        .FirstOrDefault();
                    if (latestRecord != null)
                    {
                        _lastSerialNum = latestRecord.TR_SerialNum;
                        Logger.Info($"🔄 设置最新SerialNum基线: '{_lastSerialNum}'");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 初始化已知记录基线失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式初始化已知记录（异步版本）
        /// </summary>
        private async Task InitializeKnownRecordsAsync()
        {
            try
            {
                Logger.Info("🏁 开始异步初始化已知记录基线...");
                
                using var connection = new OleDbConnection(_connectionString);
                await connection.OpenAsync();
                
                using var command = new OleDbCommand($"SELECT TR_SerialNum, TR_ID FROM [{_currentTableName}]", connection);
                using var reader = await command.ExecuteReaderAsync();
                
                var currentRecords = new List<TestRecord>();
                while (await reader.ReadAsync())
                {
                    var serialNum = reader.IsDBNull("TR_SerialNum") ? string.Empty : reader.GetString("TR_SerialNum");
                    var id = reader.IsDBNull("TR_ID") ? string.Empty : reader.GetString("TR_ID");
                    
                    currentRecords.Add(new TestRecord
                    {
                        TR_SerialNum = serialNum,
                        TR_ID = id
                    });
                }
                
                lock (_knownRecords)
                {
                    _knownRecords.Clear();
                    foreach (var record in currentRecords)
                    {
                        _knownRecords.Add(record);
                    }
                }
                
                Logger.Info($"✅ 异步已知记录基线初始化完成，共 {currentRecords.Count} 条记录");
                
                // 🔧 修复：使用TR_SerialNum排序而不是TR_ID来获取最新记录
                if (currentRecords.Count > 0)
                {
                    var latestRecord = currentRecords
                        .Where(r => !string.IsNullOrEmpty(r.TR_SerialNum))
                        .OrderByDescending(r => r.TR_SerialNum)
                        .FirstOrDefault();
                    if (latestRecord != null)
                    {
                        _lastSerialNum = latestRecord.TR_SerialNum;
                        Logger.Info($"🔄 设置最新SerialNum基线（异步）: '{_lastSerialNum}'");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 异步初始化已知记录基线失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式检查新记录（简化逻辑）
        /// </summary>
        private async void CheckForNewRecords(object? state)
        {
            if (!_isRunning) return;

            try
            {
                var tableName = !string.IsNullOrEmpty(_currentTableName) ? _currentTableName : "TestRecord";
                Logger.Info($"⏰ 检查新记录: 表={tableName}");
                
                // 🔧 新增：向UI报告监控活动状态
                StatusChanged?.Invoke(this, $"⏰ 正在检查新记录: 表={tableName}");

                // 获取所有当前记录（按照AccessDatabaseMonitor方式）
                var currentRecords = await GetAllRecordsAsync();
                
                // 🔧 修复：改进新记录检测逻辑，更加宽松
                List<TestRecord> newRecords;
                lock (_knownRecords)
                {
                    // 🔧 关键修复：不要求TR_ID必须非空，使用更宽松的条件
                    // 只要TR_SerialNum不为空，或者整个记录在已知记录中不存在，就认为是新记录
                    newRecords = currentRecords.Where(record => 
                        !string.IsNullOrEmpty(record.TR_SerialNum) && // 至少有SerialNum
                        !_knownRecords.Contains(record) // 不在已知记录中
                    ).ToList();
                    
                    // 🔧 增强调试：显示详细的检测信息
                    Logger.Info($"🔍 检测详情 - 当前记录: {currentRecords.Count}, 已知记录: {_knownRecords.Count}");
                    foreach (var record in currentRecords.Take(3)) // 显示前3条记录用于调试
                    {
                        var isKnown = _knownRecords.Contains(record);
                        var hasSerialNum = !string.IsNullOrEmpty(record.TR_SerialNum);
                        Logger.Info($"🔍 记录检查: SerialNum={record.TR_SerialNum}, ID={record.TR_ID}, 有SerialNum={hasSerialNum}, 已知={isKnown}");
                    }
                    
                    // 添加新记录到已知记录集合
                    foreach (var record in newRecords)
                    {
                        _knownRecords.Add(record);
                    }
                }

                Logger.Info($"📊 检查结果 - 当前记录总数: {currentRecords.Count}, 新增记录: {newRecords.Count}");
                
                // 🔧 新增：向UI报告监控检查结果
                StatusChanged?.Invoke(this, $"📊 监控检查完成 - 总记录: {currentRecords.Count}, 已知记录: {_knownRecords.Count}, 新增: {newRecords.Count}");

                // 处理新记录
                if (newRecords.Count > 0)
                {
                    Logger.Info($"🎯 发现 {newRecords.Count} 条新记录");
                    
                    // 🔧 新增：向UI报告发现新记录
                    StatusChanged?.Invoke(this, $"🎯 发现 {newRecords.Count} 条新记录！");
                    
                    foreach (var record in newRecords)
                    {
                        Logger.Info($"✅ 新记录: TR_ID={record.TR_ID}, SerialNum={record.TR_SerialNum}");
                        
                        // 🔧 新增：向UI报告每条新记录
                        StatusChanged?.Invoke(this, $"📋 新记录详情: ID={record.TR_ID}, SerialNum={record.TR_SerialNum}");
                        
                        // 更新最新SerialNum以保持兼容性
                        if (!string.IsNullOrEmpty(record.TR_SerialNum))
                        {
                            _lastSerialNum = record.TR_SerialNum;
                        }
                        
                        // 触发新记录事件
                        NewRecordFound?.Invoke(this, record);
                    }
                    NewRecordsDetected?.Invoke(newRecords); // 触发新记录事件
                }
                else
                {
                    Logger.Info("📝 未发现新记录");
                    // 🔧 新增：向UI报告未发现新记录（但不要过于频繁）
                    // 只在每5次检查时报告一次状态，避免日志过多
                    if (_monitoringCycleCount % 5 == 0)
                    {
                        StatusChanged?.Invoke(this, $"📝 监控正常运行 - 未发现新记录 (检查周期: {_monitoringCycleCount})");
                    }
                    _monitoringCycleCount++;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 监控检查失败: {ex.Message}", ex);
                MonitoringError?.Invoke(this, $"监控检查失败: {ex.Message}");
                // 🔧 新增：向UI报告监控异常
                StatusChanged?.Invoke(this, $"❌ 监控检查异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式获取所有记录（智能字段检测）
        /// </summary>
        private async Task<List<TestRecord>> GetAllRecordsAsync()
        {
            var records = new List<TestRecord>();

            using var connection = new OleDbConnection(_connectionString);
            await connection.OpenAsync();

            // 🔧 添加智能字段检测，确保监控功能正常
            var availableFields = GetAvailableFields(connection, _currentTableName);
            Logger.Info($"🔍 监控查询 - 数据表 [{_currentTableName}] 中可用字段: {string.Join(", ", availableFields)}");

            // 构建查询语句，确保字段存在
            var fieldList = new List<string>();
            
            if (availableFields.Contains("TR_SerialNum", StringComparer.OrdinalIgnoreCase))
            {
                fieldList.Add("TR_SerialNum");
            }
            
            if (availableFields.Contains("TR_ID", StringComparer.OrdinalIgnoreCase))
            {
                fieldList.Add("TR_ID");
            }

            if (fieldList.Count == 0)
            {
                Logger.Error("❌ 无法找到TR_SerialNum或TR_ID字段，监控无法工作");
                return records;
            }

            var fieldsToSelect = string.Join(", ", fieldList);
            var query = $"SELECT {fieldsToSelect} FROM [{_currentTableName}]";
            
            Logger.Info($"🔍 监控查询SQL: {query}");

            using var command = new OleDbCommand(query, connection);
            using var reader = await command.ExecuteReaderAsync() as OleDbDataReader;

            if (reader == null)
            {
                Logger.Error("❌ 无法获取数据读取器");
                return records;
            }

            while (await reader.ReadAsync())
            {
                var record = new TestRecord
                {
                    TR_SerialNum = GetSafeString(reader, "TR_SerialNum"),
                    TR_ID = GetSafeString(reader, "TR_ID")
                };

                records.Add(record);
            }

            Logger.Info($"✅ 监控查询成功获取到 {records.Count} 条记录");
            return records;
        }

        // 现代化的SerialNum检查方法（保留兼容性，现在调用新的检查方法）
        private void CheckForNewSerialNum()
        {
            Logger.Info($"⏰ 调用兼容性SerialNum检查方法（现已使用新的记录检查机制）");
            CheckForNewRecords(null); // Pass null as state
        }

        // 获取当前最新SerialNum（用于设置基线，现代化查询）
        private void GetCurrentLatestSerialNum()
        {
            var tableName = !string.IsNullOrEmpty(_currentTableName) ? _currentTableName : "TestRecord";

            try
            {
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 使用现代化查询方式
                var query = @"
                    SELECT TR_SerialNum 
                    FROM [" + tableName + @"] 
                    WHERE TR_SerialNum IS NOT NULL 
                    AND LEN(TRIM(TR_SerialNum)) > 0 
                    ORDER BY TR_ID DESC";
                    
                using var command = new OleDbCommand(query, connection);
                using var reader = command.ExecuteReader();
                
                if (reader.Read())
                {
                    _lastSerialNum = reader["TR_SerialNum"]?.ToString()?.Trim() ?? "";
                }
                else
                {
                    _lastSerialNum = "";
                }
                reader.Close();
                
                Logger.Info($"📋 初始化最新SerialNum基线: '{_lastSerialNum}'");
            }
            catch (Exception ex)
            {
                Logger.Error($"获取最新SerialNum失败: {ex.Message}", ex);
                _lastSerialNum = "";
            }
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式根据记录ID获取记录（简化查询）
        /// </summary>
        private TestRecord? GetRecordByRecordId(OleDbConnection? connection, string tableName, string recordId)
        {
            try
            {
                // 按照AccessDatabaseMonitor的简单查询方式
                var query = $"SELECT TR_SerialNum, TR_ID FROM [{tableName}] WHERE TR_ID = ?";
                    
                using var command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("?", recordId);
                using var reader = command.ExecuteReader();

                if (reader.Read())
                {
                    var serialNum = reader.IsDBNull("TR_SerialNum") ? string.Empty : reader.GetString("TR_SerialNum");
                    var id = reader.IsDBNull("TR_ID") ? string.Empty : reader.GetString("TR_ID");

                    var record = new TestRecord
                    {
                        TR_SerialNum = serialNum,
                        TR_ID = id
                    };
                    
                    Logger.Info($"✅ 成功获取记录: ID={record.TR_ID}, SerialNum={record.TR_SerialNum}");
                    return record;
                }
                else
                {
                    Logger.Warning($"⚠️ 未找到记录ID为 {recordId} 的记录");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 根据记录ID获取记录失败: {ex.Message}", ex);
            }
            return null;
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式根据SerialNum获取记录（简化查询）
        /// </summary>
        private TestRecord? GetRecordBySerialNum(OleDbConnection connection, string tableName, string serialNum)
        {
            try
            {
                // 按照AccessDatabaseMonitor的简单查询方式
                var query = $"SELECT TR_SerialNum, TR_ID FROM [{tableName}] WHERE TR_SerialNum = ? ORDER BY TR_ID DESC";
                    
                using var command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("?", serialNum);
                using var reader = command.ExecuteReader();

                if (reader.Read())
                {
                    var serialNumResult = reader.IsDBNull("TR_SerialNum") ? string.Empty : reader.GetString("TR_SerialNum");
                    var id = reader.IsDBNull("TR_ID") ? string.Empty : reader.GetString("TR_ID");

                    return new TestRecord
                    {
                        TR_SerialNum = serialNumResult,
                        TR_ID = id
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 根据SerialNum获取记录失败: {ex.Message}", ex);
            }
            return null;
        }

        private TestRecord MapReaderToTestRecord(IDataReader reader)
        {
            var record = new TestRecord();

            try
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var fieldName = reader.GetName(i);
                    var value = reader.IsDBNull(i) ? null : reader.GetValue(i);

                    switch (fieldName)
                    {
                        case "TR_ID":
                            record.TR_ID = value?.ToString();
                            break;
                        case "TR_SerialNum":
                            record.TR_SerialNum = value?.ToString();
                            break;
                        case "TR_Isc":
                            record.TR_Isc = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Voc":
                            record.TR_Voc = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Pm":
                            record.TR_Pm = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Ipm":
                            record.TR_Ipm = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Vpm":
                            record.TR_Vpm = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_DateTime":
                            record.TR_DateTime = value as DateTime? ?? (value != null ? Convert.ToDateTime(value) : null);
                            break;
                        case "TR_Print":
                            record.TR_Print = value as int? ?? (value != null ? Convert.ToInt32(value) : null);
                            break;
                        case "TR_CellEfficiency":
                            record.TR_CellEfficiency = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_FF":
                            record.TR_FF = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Grade":
                            record.TR_Grade = value?.ToString();
                            break;
                        case "TR_Temp":
                            record.TR_Temp = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Irradiance":
                            record.TR_Irradiance = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Rs":
                            record.TR_Rs = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_Rsh":
                            record.TR_Rsh = value as decimal? ?? (value != null ? Convert.ToDecimal(value) : null);
                            break;
                        case "TR_CellArea":
                            record.TR_CellArea = value?.ToString();
                            break;
                        case "TR_Operater":
                            record.TR_Operater = value?.ToString();
                            break;
                        case "TR_FontColor":
                            record.TR_FontColor = value?.ToString();
                            break;
                        case "TR_BackColor":
                            record.TR_BackColor = value?.ToString();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"记录映射失败: {ex.Message}", ex);
            }

            return record;
        }









        public bool UpdatePrintCount(string recordId)
        {
            try
            {
                var config = ConfigurationManager.Config;
                
                // 检查是否启用打印次数功能
                if (!config.Database.EnablePrintCount)
                {
                    Logger.Info($"打印次数统计已禁用，跳过更新: 记录ID={recordId}");
                    return true; // 返回成功但不执行数据库更新
                }
                
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 检查TR_Print字段是否存在
                var availableFields = GetAvailableFields(connection, config.Database.TableName);
                if (!availableFields.Contains("TR_Print", StringComparer.OrdinalIgnoreCase))
                {
                    Logger.Warning("数据库中不存在TR_Print字段，无法更新打印次数统计，但不影响打印操作");
                    return true; // 返回成功但不执行数据库更新
                }

                var updateQuery = $"UPDATE [{config.Database.TableName}] SET TR_Print = IIF(IsNull(TR_Print), 0, TR_Print) + 1 WHERE [{config.Database.MonitorField}] = ?";
                using var command = new OleDbCommand(updateQuery, connection);
                command.Parameters.AddWithValue("?", recordId);

                var affectedRows = command.ExecuteNonQuery();
                Logger.Info($"更新打印计数: 记录ID={recordId}, 影响行数={affectedRows}");

                return affectedRows > 0;
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印计数失败: {ex.Message}");
                // 如果是因为字段不存在导致的错误，返回true以免阻止打印
                if (ex.Message.Contains("TR_Print") || ex.Message.Contains("找不到列") || ex.Message.Contains("column"))
                {
                    Logger.Warning("TR_Print字段不存在，跳过打印次数更新");
                    return true;
                }
                return false;
            }
        }

        // 重载方法：接受TestRecord对象，使用TR_ID（主键）进行更新
        public bool UpdatePrintCount(TestRecord record)
        {
            try
            {
                var config = ConfigurationManager.Config;
                
                // 检查是否启用打印次数功能
                if (!config.Database.EnablePrintCount)
                {
                    Logger.Info($"打印次数统计已禁用，跳过更新: TR_ID={record.TR_ID}");
                    return true; // 返回成功但不执行数据库更新
                }
                
                // TR_ID是主键，必定存在且唯一
                if (string.IsNullOrEmpty(record.TR_ID))
                {
                    Logger.Error("TR_ID为空，无法更新打印计数（TR_ID是必需的主键）");
                    return false;
                }
                
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 检查TR_Print字段是否存在
                var availableFields = GetAvailableFields(connection, config.Database.TableName);
                if (!availableFields.Contains("TR_Print", StringComparer.OrdinalIgnoreCase))
                {
                    Logger.Warning("数据库中不存在TR_Print字段，无法更新打印次数统计，但不影响打印操作");
                    return true; // 返回成功但不执行数据库更新
                }

                // 直接使用TR_ID主键进行精确更新
                var updateQuery = $"UPDATE [{config.Database.TableName}] SET TR_Print = IIF(IsNull(TR_Print), 0, TR_Print) + 1 WHERE TR_ID = ?";
                using var command = new OleDbCommand(updateQuery, connection);
                command.Parameters.AddWithValue("?", record.TR_ID);

                var affectedRows = command.ExecuteNonQuery();
                Logger.Info($"通过主键TR_ID更新打印计数: {record.TR_ID}, 影响行数={affectedRows}");

                return affectedRows > 0;
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印计数失败: {ex.Message}", ex);
                // 如果是因为字段不存在导致的错误，返回true以免阻止打印
                if (ex.Message.Contains("TR_Print") || ex.Message.Contains("找不到列") || ex.Message.Contains("column"))
                {
                    Logger.Warning("TR_Print字段不存在，跳过打印次数更新");
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式获取最近记录（智能字段检测）
        /// </summary>
        public List<TestRecord> GetRecentRecords(int limit = 10)
        {
            var records = new List<TestRecord>();

            try
            {
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 🔧 动态检测表中存在的字段
                var availableFields = GetAvailableFields(connection, _currentTableName);
                Logger.Info($"🔍 数据表 [{_currentTableName}] 中可用字段: {string.Join(", ", availableFields)}");

                // 构建查询语句，只查询存在的字段
                var fieldList = new List<string> { "TR_SerialNum", "TR_ID" }; // 基础必需字段

                // 添加可选字段（如果存在）
                var optionalFields = new[]
                {
                    "TR_DateTime", "TR_Isc", "TR_Voc", "TR_Pm", "TR_Ipm", "TR_Vpm", "TR_Print"
                };

                foreach (var field in optionalFields)
                {
                    if (availableFields.Contains(field, StringComparer.OrdinalIgnoreCase))
                    {
                        fieldList.Add(field);
                        Logger.Info($"✅ 字段 {field} 存在，将被查询");
                    }
                    else
                    {
                        Logger.Warning($"❌ 字段 {field} 不存在，将跳过");
                    }
                }

                var fieldsToSelect = string.Join(", ", fieldList);
                
                // 构建查询语句 - 优先按数据库默认顺序，后备按时间倒序
                string orderClause = "";
                string orderDescription = "按数据库默认顺序";
                
                if (availableFields.Contains("TR_DateTime", StringComparer.OrdinalIgnoreCase))
                {
                    orderClause = " ORDER BY TR_DateTime DESC";
                    orderDescription = "按测试时间倒序";
                }
                else if (availableFields.Contains("TR_ID", StringComparer.OrdinalIgnoreCase))
                {
                    orderClause = " ORDER BY TR_ID DESC";
                    orderDescription = "按记录ID倒序";
                }
                
                var query = $@"SELECT TOP {limit} {fieldsToSelect}
                    FROM [{_currentTableName}]{orderClause}";
                
                Logger.Info($"🔍 执行智能字段查询（{orderDescription}），限制 {limit} 条");
                Logger.Info($"🔍 查询SQL: {query}");
                
                using var command = new OleDbCommand(query, connection);
                using var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    var record = new TestRecord
                    {
                        // 基础必需字段
                        TR_SerialNum = GetSafeString(reader, "TR_SerialNum"),
                        TR_ID = GetSafeString(reader, "TR_ID"),
                        
                        // 可选字段（只有在字段存在时才读取）
                        TR_DateTime = fieldList.Contains("TR_DateTime") ? GetSafeDateTime(reader, "TR_DateTime") : null,
                        TR_Isc = fieldList.Contains("TR_Isc") ? GetSafeDecimal(reader, "TR_Isc") : null,
                        TR_Voc = fieldList.Contains("TR_Voc") ? GetSafeDecimal(reader, "TR_Voc") : null,
                        TR_Pm = fieldList.Contains("TR_Pm") ? GetSafeDecimal(reader, "TR_Pm") : null,
                        TR_Ipm = fieldList.Contains("TR_Ipm") ? GetSafeDecimal(reader, "TR_Ipm") : null,
                        TR_Vpm = fieldList.Contains("TR_Vpm") ? GetSafeDecimal(reader, "TR_Vpm") : null,
                        TR_Print = fieldList.Contains("TR_Print") ? GetSafeInt(reader, "TR_Print") : null
                    };

                    records.Add(record);
                }

                Logger.Info($"✅ 成功获取到 {records.Count} 条最近记录（智能字段查询）");
                
                // 🔧 添加第一条记录的字段值检查（用于诊断）
                if (records.Count > 0)
                {
                    var firstRecord = records[0];
                    Logger.Info($"🔍 第一条记录字段值检查:");
                    Logger.Info($"  TR_SerialNum: '{firstRecord.TR_SerialNum}'");
                    Logger.Info($"  TR_ID: '{firstRecord.TR_ID}'");
                    Logger.Info($"  TR_DateTime: {firstRecord.TR_DateTime}");
                    Logger.Info($"  TR_Isc: {firstRecord.TR_Isc}");
                    Logger.Info($"  TR_Voc: {firstRecord.TR_Voc}");
                    Logger.Info($"  TR_Pm: {firstRecord.TR_Pm}");
                    Logger.Info($"  TR_Ipm: {firstRecord.TR_Ipm}");
                    Logger.Info($"  TR_Vpm: {firstRecord.TR_Vpm}");
                    Logger.Info($"  TR_Print: {firstRecord.TR_Print}");
                    
                    // 🔧 增加字段包含状态检查
                    Logger.Info($"🔍 字段包含状态检查:");
                    Logger.Info($"  包含TR_DateTime: {fieldList.Contains("TR_DateTime")}");
                    Logger.Info($"  包含TR_Isc: {fieldList.Contains("TR_Isc")}");
                    Logger.Info($"  包含TR_Voc: {fieldList.Contains("TR_Voc")}");
                    Logger.Info($"  包含TR_Pm: {fieldList.Contains("TR_Pm")}");
                    Logger.Info($"  包含TR_Ipm: {fieldList.Contains("TR_Ipm")}");
                    Logger.Info($"  包含TR_Vpm: {fieldList.Contains("TR_Vpm")}");
                    Logger.Info($"  包含TR_Print: {fieldList.Contains("TR_Print")}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 获取最近记录失败: {ex.Message}", ex);
                Logger.Error($"❌ 连接字符串: {_connectionString}");
                Logger.Error($"❌ 表名: {_currentTableName}");
            }

            return records;
        }

        /// <summary>
        /// 获取数据表中的所有可用字段
        /// </summary>
        private List<string> GetAvailableFields(OleDbConnection connection, string tableName)
        {
            var fields = new List<string>();
            try
            {
                var schema = connection.GetSchema("Columns", new[] { null, null, tableName, null });
                foreach (DataRow row in schema.Rows)
                {
                    var columnName = row["COLUMN_NAME"].ToString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        fields.Add(columnName);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"获取表字段失败: {ex.Message}");
                // 如果获取字段失败，返回基础字段列表
                fields.AddRange(new[] { "TR_SerialNum", "TR_ID", "TR_DateTime", "TR_Isc", "TR_Voc", "TR_Pm", "TR_Ipm", "TR_Vpm" });
            }
            return fields;
        }

        /// <summary>
        /// 安全获取字符串字段
        /// </summary>
        private string GetSafeString(OleDbDataReader reader, string fieldName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(fieldName);
                return reader.IsDBNull(ordinal) ? string.Empty : reader.GetString(ordinal);
            }
            catch (Exception ex)
            {
                Logger.Warning($"⚠️ 字段 {fieldName} 读取失败: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 安全获取日期时间字段（支持多种日期类型）
        /// </summary>
        private DateTime? GetSafeDateTime(OleDbDataReader reader, string fieldName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(fieldName);
                if (reader.IsDBNull(ordinal))
                {
                    return null;
                }

                // 尝试多种日期时间类型转换
                var value = reader.GetValue(ordinal);
                // 🔧 减少日志输出，只在转换失败时记录
                
                if (value is DateTime dateTime)
                {
                    return dateTime;
                }
                else if (value is string strValue && DateTime.TryParse(strValue, out var parsedDateTime))
                {
                    return parsedDateTime;
                }
                else if (DateTime.TryParse(value?.ToString(), out var parsedFromToString))
                {
                    return parsedFromToString;
                }
                else
                {
                    Logger.Warning($"⚠️ 字段 {fieldName} 无法转换为DateTime，原始值: {value}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"⚠️ 字段 {fieldName} 读取失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 安全获取decimal字段（支持多种数值类型）
        /// </summary>
        private decimal? GetSafeDecimal(OleDbDataReader reader, string fieldName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(fieldName);
                if (reader.IsDBNull(ordinal))
                {
                    return null;
                }

                // 尝试多种数据类型转换
                var value = reader.GetValue(ordinal);
                // 🔧 减少日志输出，只在转换失败时记录

                if (value is decimal dec)
                {
                    return dec;
                }
                else if (value is double dbl)
                {
                    return (decimal)dbl;
                }
                else if (value is float flt)
                {
                    return (decimal)flt;
                }
                else if (value is int intVal)
                {
                    return intVal;
                }
                else if (value is long longVal)
                {
                    return longVal;
                }
                else if (decimal.TryParse(value?.ToString(), out var parsedDecimal))
                {
                    return parsedDecimal;
                }
                else
                {
                    Logger.Warning($"⚠️ 字段 {fieldName} 无法转换为decimal，原始值: {value}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"⚠️ 字段 {fieldName} 读取失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 安全获取int字段（支持多种数值类型）
        /// </summary>
        private int? GetSafeInt(OleDbDataReader reader, string fieldName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(fieldName);
                if (reader.IsDBNull(ordinal))
                {
                    return 0; // 打印次数默认为0
                }

                // 尝试多种数据类型转换
                var value = reader.GetValue(ordinal);
                // 🔧 减少日志输出，只在转换失败时记录

                if (value is int intVal)
                {
                    return intVal;
                }
                else if (value is long longVal)
                {
                    return (int)longVal;
                }
                else if (value is decimal decVal)
                {
                    return (int)decVal;
                }
                else if (value is double dblVal)
                {
                    return (int)dblVal;
                }
                else if (value is float fltVal)
                {
                    return (int)fltVal;
                }
                else if (int.TryParse(value?.ToString(), out var parsedInt))
                {
                    return parsedInt;
                }
                else
                {
                    Logger.Warning($"⚠️ 字段 {fieldName} 无法转换为int，原始值: {value}，返回默认值0");
                    return 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"⚠️ 字段 {fieldName} 读取失败: {ex.Message}，返回默认值0");
                return 0; // 默认打印次数为0
            }
        }

        /// <summary>
        /// 按照AccessDatabaseMonitor方式获取最后一条记录（修复排序）
        /// 🔧 修复：获取完整记录数据，而不只是TR_SerialNum和TR_ID
        /// </summary>
        public TestRecord? GetLastRecord()
        {
            try
            {
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 🔧 动态检测表中存在的字段
                var availableFields = GetAvailableFields(connection, _currentTableName);
                Logger.Info($"🔍 获取最后记录时检测到的字段: {string.Join(", ", availableFields)}");

                // 构建完整的字段列表
                var fieldList = new List<string> { "TR_SerialNum", "TR_ID" }; // 基础必需字段

                // 添加可选字段（如果存在）
                var optionalFields = new[]
                {
                    "TR_DateTime", "TR_Isc", "TR_Voc", "TR_Pm", "TR_Ipm", "TR_Vpm", "TR_Print",
                    "TR_CellEfficiency", "TR_FF", "TR_Grade", "TR_Temp", "TR_Irradiance", 
                    "TR_Rs", "TR_Rsh", "TR_CellArea", "TR_Operater", "TR_FontColor", "TR_BackColor"
                };

                foreach (var field in optionalFields)
                {
                    if (availableFields.Contains(field, StringComparer.OrdinalIgnoreCase))
                    {
                        fieldList.Add(field);
                    }
                }

                var fieldsToSelect = string.Join(", ", fieldList);
                
                // 🔧 修复：使用TR_SerialNum排序而不是TR_ID，与AccessDatabaseMonitor一致
                var query = $"SELECT TOP 1 {fieldsToSelect} FROM [{_currentTableName}] ORDER BY TR_SerialNum DESC";
                
                Logger.Info($"🔍 执行完整最后记录查询（包含所有字段）");
                Logger.Info($"🔍 查询SQL: {query}");
                    
                using var command = new OleDbCommand(query, connection);
                using var reader = command.ExecuteReader();

                if (reader.Read())
                {
                    var record = new TestRecord
                    {
                        // 基础必需字段
                        TR_SerialNum = GetSafeString(reader, "TR_SerialNum"),
                        TR_ID = GetSafeString(reader, "TR_ID"),
                        
                        // 🔧 修复：添加所有数值字段，确保打印时有完整数据
                        TR_DateTime = fieldList.Contains("TR_DateTime") ? GetSafeDateTime(reader, "TR_DateTime") : null,
                        TR_Isc = fieldList.Contains("TR_Isc") ? GetSafeDecimal(reader, "TR_Isc") : null,
                        TR_Voc = fieldList.Contains("TR_Voc") ? GetSafeDecimal(reader, "TR_Voc") : null,
                        TR_Pm = fieldList.Contains("TR_Pm") ? GetSafeDecimal(reader, "TR_Pm") : null,
                        TR_Ipm = fieldList.Contains("TR_Ipm") ? GetSafeDecimal(reader, "TR_Ipm") : null,
                        TR_Vpm = fieldList.Contains("TR_Vpm") ? GetSafeDecimal(reader, "TR_Vpm") : null,
                        TR_Print = fieldList.Contains("TR_Print") ? GetSafeInt(reader, "TR_Print") : null,
                        TR_CellEfficiency = fieldList.Contains("TR_CellEfficiency") ? GetSafeDecimal(reader, "TR_CellEfficiency") : null,
                        TR_FF = fieldList.Contains("TR_FF") ? GetSafeDecimal(reader, "TR_FF") : null,
                        TR_Grade = fieldList.Contains("TR_Grade") ? GetSafeString(reader, "TR_Grade") : null,
                        TR_Temp = fieldList.Contains("TR_Temp") ? GetSafeDecimal(reader, "TR_Temp") : null,
                        TR_Irradiance = fieldList.Contains("TR_Irradiance") ? GetSafeDecimal(reader, "TR_Irradiance") : null,
                        TR_Rs = fieldList.Contains("TR_Rs") ? GetSafeDecimal(reader, "TR_Rs") : null,
                        TR_Rsh = fieldList.Contains("TR_Rsh") ? GetSafeDecimal(reader, "TR_Rsh") : null,
                        TR_CellArea = fieldList.Contains("TR_CellArea") ? GetSafeString(reader, "TR_CellArea") : null,
                        TR_Operater = fieldList.Contains("TR_Operater") ? GetSafeString(reader, "TR_Operater") : null,
                        TR_FontColor = fieldList.Contains("TR_FontColor") ? GetSafeString(reader, "TR_FontColor") : null,
                        TR_BackColor = fieldList.Contains("TR_BackColor") ? GetSafeString(reader, "TR_BackColor") : null
                    };
                    
                    Logger.Info($"✅ 获取完整最后记录: TR_ID={record.TR_ID}, SerialNum={record.TR_SerialNum}");
                    Logger.Info($"🔍 数值字段检查: Isc={record.TR_Isc}, Voc={record.TR_Voc}, Pm={record.TR_Pm}, Ipm={record.TR_Ipm}, Vpm={record.TR_Vpm}");
                    return record;
                }
                else
                {
                    Logger.Info("📊 表中没有记录");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 获取最后记录失败: {ex.Message}", ex);
                Logger.Error($"❌ 连接字符串: {_connectionString}");
                Logger.Error($"❌ 表名: {_currentTableName}");
            }

            return null;
        }

        // 🔧 新增：基于GetLastRecord的简化监控逻辑
        private TestRecord? _lastKnownRecord = null;
        
        /// <summary>
        /// 🔧 统一监控系统：基于GetLastRecord的完整数据管理
        /// 检测最后记录变化，同时获取最新50条记录，实现统一数据刷新
        /// 🛠️ 修复：增强连接健康检查和异常恢复机制
        /// </summary>
        private void CheckForLastRecordChanges(object? state)
        {
            if (!_isRunning) return;

            try
            {
                var tableName = !string.IsNullOrEmpty(_currentTableName) ? _currentTableName : "TestRecord";
                
                // 🔧 每10次检查才输出一次日志，避免日志过多
                if (_monitoringCycleCount % 10 == 0)
                {
                    Logger.Info($"⏰ 统一监控检查: 表={tableName} (周期#{_monitoringCycleCount})");
                    StatusChanged?.Invoke(this, $"⏰ 统一监控运行: 表={tableName} (周期#{_monitoringCycleCount})");
                }
                
                _monitoringCycleCount++;

                // 🛠️ 新增：连接健康检查和重新连接逻辑
                if (!IsConnectionHealthy())
                {
                    Logger.Warning("🔧 检测到连接问题，尝试重新建立连接...");
                    StatusChanged?.Invoke(this, "🔧 连接异常，正在重新连接...");
                    
                    if (!AttemptReconnection())
                    {
                        Logger.Error("❌ 连接重建失败，监控暂停本次检查");
                        StatusChanged?.Invoke(this, "❌ 连接重建失败");
                        return;
                    }
                    
                    Logger.Info("✅ 连接重建成功，继续监控");
                    StatusChanged?.Invoke(this, "✅ 连接重建成功");
                }

                // 🔧 核心：统一监控只使用GetLastRecord
                var currentLastRecord = GetLastRecord();
                
                if (currentLastRecord == null)
                {
                    if (_monitoringCycleCount % 10 == 0)
                    {
                        Logger.Info("📊 数据库中没有记录");
                        StatusChanged?.Invoke(this, "📊 数据库中没有记录");
                    }
                    return;
                }

                // 检查是否是第一次获取记录
                if (_lastKnownRecord == null)
                {
                    _lastKnownRecord = currentLastRecord;
                    Logger.Info($"🏁 统一监控基线: ID={currentLastRecord.TR_ID}, SerialNum={currentLastRecord.TR_SerialNum}");
                    StatusChanged?.Invoke(this, $"🏁 监控基线: {currentLastRecord.TR_SerialNum}");
                    
                    // 🔧 新增：初始化时也获取50条记录并发送统一事件
                    var initialRecords = GetRecentRecords(50);
                    var initialEventArgs = new DataUpdateEventArgs(
                        currentLastRecord, 
                        initialRecords, 
                        "初始化", 
                        $"监控基线设置: {currentLastRecord.TR_SerialNum}"
                    );
                    DataUpdated?.Invoke(this, initialEventArgs);
                    Logger.Info($"📋 初始化获取到 {initialRecords.Count} 条记录");
                    
                    return;
                }

                // 🔧 关键逻辑：检测最后记录是否发生变化
                bool hasChanged = false;
                string changeDetails = "";
                
                if (!string.Equals(_lastKnownRecord.TR_SerialNum, currentLastRecord.TR_SerialNum, StringComparison.OrdinalIgnoreCase))
                {
                    hasChanged = true;
                    changeDetails = $"SerialNum: '{_lastKnownRecord.TR_SerialNum}' -> '{currentLastRecord.TR_SerialNum}'";
                }
                else if (!string.Equals(_lastKnownRecord.TR_ID, currentLastRecord.TR_ID, StringComparison.OrdinalIgnoreCase))
                {
                    hasChanged = true;
                    changeDetails = $"TR_ID: '{_lastKnownRecord.TR_ID}' -> '{currentLastRecord.TR_ID}'";
                }

                if (hasChanged)
                {
                    Logger.Info($"🎯 统一监控检测到数据库更新！{changeDetails}");
                    StatusChanged?.Invoke(this, $"🎯 检测到数据更新：{currentLastRecord.TR_SerialNum}");
                    
                    // 更新已知的最后记录
                    _lastKnownRecord = currentLastRecord;
                    
                    // 🔧 核心：统一数据获取 - 基于GetLastRecord检测，一次性获取完整数据
                    Logger.Info("📋 基于LastRecord变化，获取最新50条记录...");
                    var recentRecords = GetRecentRecords(50);
                    Logger.Info($"📋 统一获取到 {recentRecords.Count} 条最新记录");
                    
                    // 🔧 核心：发送统一数据更新事件 - 包含最后记录和50条记录列表
                    var dataUpdateArgs = new DataUpdateEventArgs(
                        currentLastRecord, 
                        recentRecords, 
                        "记录更新", 
                        changeDetails
                    );
                    DataUpdated?.Invoke(this, dataUpdateArgs);
                    
                    // 保持兼容性：继续触发原有事件
                    Logger.Info($"🔔 触发兼容性事件: TR_ID={currentLastRecord.TR_ID}, SerialNum={currentLastRecord.TR_SerialNum}");
                    NewRecordFound?.Invoke(this, currentLastRecord);
                    NewRecordsDetected?.Invoke(new List<TestRecord> { currentLastRecord });
                }
                else
                {
                    // 只在特定周期输出"无变化"的日志
                    if (_monitoringCycleCount % 30 == 0) // 每30次检查输出一次
                    {
                        Logger.Info($"📝 LastRecord无变化: {currentLastRecord.TR_SerialNum}");
                        StatusChanged?.Invoke(this, $"📝 监控正常 - 最后记录: {currentLastRecord.TR_SerialNum}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 统一监控检查失败: {ex.Message}", ex);
                MonitoringError?.Invoke(this, $"监控检查失败: {ex.Message}");
                StatusChanged?.Invoke(this, $"❌ 监控异常: {ex.Message}");
                
                // 🛠️ 新增：异常后的恢复逻辑
                _retryCount++;
                if (_retryCount <= MaxRetries)
                {
                    Logger.Warning($"⚠️ 监控异常，第 {_retryCount}/{MaxRetries} 次重试");
                    StatusChanged?.Invoke(this, $"⚠️ 监控异常重试 {_retryCount}/{MaxRetries}");
                }
                else
                {
                    Logger.Error($"❌ 监控连续失败超过 {MaxRetries} 次，停止监控");
                    StatusChanged?.Invoke(this, "❌ 监控失败次数过多，已停止");
                    StopMonitoring();
                }
            }
        }

        /// <summary>
        /// 🛠️ 新增：检查数据库连接健康状况
        /// </summary>
        private bool IsConnectionHealthy()
        {
            try
            {
                if (string.IsNullOrEmpty(_connectionString))
                {
                    return false;
                }

                using var connection = new OleDbConnection(_connectionString);
                connection.Open();
                
                // 执行简单查询测试连接
                using var command = new OleDbCommand($"SELECT COUNT(*) FROM [{_currentTableName}]", connection);
                var result = command.ExecuteScalar();
                
                return result != null;
            }
            catch (Exception ex)
            {
                Logger.Warning($"⚠️ 连接健康检查失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 🛠️ 新增：尝试重新建立数据库连接
        /// </summary>
        private bool AttemptReconnection()
        {
            try
            {
                // 重置重试计数
                _retryCount = 0;
                
                // 如果有有效的连接字符串，测试连接
                if (!string.IsNullOrEmpty(_connectionString))
                {
                    Logger.Info("🔄 尝试重新建立数据库连接...");
                    
                    using var connection = new OleDbConnection(_connectionString);
                    connection.Open();
                    
                    // 验证表仍然存在
                    using var command = new OleDbCommand($"SELECT COUNT(*) FROM [{_currentTableName}]", connection);
                    var count = command.ExecuteScalar();
                    
                    Logger.Info($"✅ 重新连接成功，表 [{_currentTableName}] 记录数: {count}");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 重新连接失败: {ex.Message}", ex);
                return false;
            }
        }

        public void Dispose()
        {
            StopMonitoring();
            _monitorTimer?.Dispose();
        }

        private bool IsACEDriverInstalled()
        {
            try
            {
                // 检查多个可能的ACE驱动注册表位置
                var registryPaths = new[]
                {
                    @"SOFTWARE\Microsoft\Jet\4.0\Engines\ACE",
                    @"SOFTWARE\Microsoft\Office\16.0\Access Connectivity Engine\Engines\ACE",
                    @"SOFTWARE\Microsoft\Office\15.0\Access Connectivity Engine\Engines\ACE",
                    @"SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\ACE",
                    @"SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\ACE",
                    @"SOFTWARE\WOW6432Node\Microsoft\Jet\4.0\Engines\ACE",
                    @"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Access Connectivity Engine\Engines\ACE"
                };

                foreach (var path in registryPaths)
                {
                    try
                    {
                        using (RegistryKey key = Registry.LocalMachine.OpenSubKey(path))
                        {
                            if (key != null)
                            {
                                Logger.Info($"找到ACE驱动注册表项: {path}");
                                return true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"检查注册表路径失败 {path}: {ex.Message}");
                    }
                }

                // 尝试通过创建连接测试来验证ACE驱动
                try
                {
                    var testConnStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=:memory:";
                    using (var conn = new OleDbConnection(testConnStr))
                    {
                        // 不需要实际连接，只需要看是否抛出"provider not registered"错误
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("not registered"))
                    {
                        Logger.Warning("ACE驱动未注册，无法创建连接");
                        return false;
                    }
                }

                Logger.Warning("未找到ACE驱动的注册表项");
                return false;
            }
            catch (Exception ex)
            {
                Logger.Warning($"检查ACE驱动失败: {ex.Message}");
                return false;
            }
        }



        // 新增：自动读取TestRecord表的所有列
        public List<string> GetTableColumns(string tableName = "")
        {
            var columns = new List<string>();
            
            try
            {
                if (string.IsNullOrEmpty(_connectionString))
                {
                    Logger.Warning("数据库未连接，无法获取表结构");
                    return columns;
                }

                var targetTable = !string.IsNullOrEmpty(tableName) ? tableName : _currentTableName;
                if (string.IsNullOrEmpty(targetTable))
                {
                    targetTable = "TestRecord"; // 默认表名
                }

                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 获取表结构信息
                var schemaTable = connection.GetSchema("Columns", new[] { null, null, targetTable, null });
                
                foreach (DataRow row in schemaTable.Rows)
                {
                    var columnName = row["COLUMN_NAME"].ToString();
                    var dataType = row["DATA_TYPE"].ToString();
                    var isNullable = row["IS_NULLABLE"].ToString();
                    
                    columns.Add(columnName);
                    Logger.Debug($"发现列: {columnName} (类型: {dataType}, 可空: {isNullable})");
                }

                // 如果GetSchema失败，使用备用方法
                if (columns.Count == 0)
                {
                    Logger.Info("使用备用方法获取表结构...");
                    using var command = new OleDbCommand($"SELECT TOP 1 * FROM [{targetTable}]", connection);
                    using var reader = command.ExecuteReader();
                    
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        var columnName = reader.GetName(i);
                        var fieldType = reader.GetFieldType(i);
                        columns.Add(columnName);
                        Logger.Debug($"发现列: {columnName} (类型: {fieldType.Name})");
                    }
                }

                // 缓存列信息
                _tableColumns = columns.ToList();
                
                Logger.Info($"成功获取表 [{targetTable}] 结构，共 {columns.Count} 列:");
                Logger.Info($"列名: {string.Join(", ", columns)}");
                
                return columns;
            }
            catch (Exception ex)
            {
                Logger.Error($"获取表结构失败: {ex.Message}", ex);
                return columns;
            }
        }

        // 新增：增强的监控功能 - 支持监控任意字段的数据变化
        // 暂时禁用，避免与新的异步连接方式冲突
        /*
        public bool StartEnhancedMonitoring(string databasePath, string tableName = "TestRecord", int pollInterval = 500)
        {
            try
            {
                // 首先连接数据库
                if (!Connect(databasePath, tableName))
                {
                    Logger.Error("增强监控启动失败：数据库连接失败");
                    return false;
                }

                // 获取表结构
                var columns = GetTableColumns(tableName);
                if (columns.Count == 0)
                {
                    Logger.Error("增强监控启动失败：无法获取表结构");
                    return false;
                }

                Logger.Info($"🔍 增强监控已配置，监控表 [{tableName}] 的所有字段变化");
                Logger.Info($"📊 监控字段清单 ({columns.Count} 个): {string.Join(", ", columns)}");

                // 启动监控
                StartMonitoring(pollInterval);
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"启动增强监控失败: {ex.Message}", ex);
                return false;
            }
        }
        */
    }
} 