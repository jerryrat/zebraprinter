using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Timers;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;
using Microsoft.Win32;

namespace ZebraPrinterMonitor.Services
{
    public class DatabaseMonitor : IDisposable
    {
        private readonly System.Timers.Timer _monitorTimer;
        private string _connectionString = "";
        private string _lastRecordId = "";
        private bool _isMonitoring = false;
        private int _retryCount = 0;
        private const int MaxRetries = 5;

        public event EventHandler<TestRecord>? NewRecordFound;
        public event EventHandler<string>? MonitoringError;
        public event EventHandler<string>? StatusChanged;

        public bool IsMonitoring => _isMonitoring;
        public string LastRecordId => _lastRecordId;

        public DatabaseMonitor()
        {
            _monitorTimer = new System.Timers.Timer();
            _monitorTimer.Elapsed += OnTimerElapsed;
            _monitorTimer.AutoReset = true;
        }

        public bool Connect(string databasePath, string tableName, string monitorField)
        {
            try
            {
                if (string.IsNullOrEmpty(databasePath) || !System.IO.File.Exists(databasePath))
                {
                    throw new Exception($"数据库文件不存在: {databasePath}");
                }

                // 检测应用程序架构
                bool isApp64Bit = Environment.Is64BitProcess;
                bool isOS64Bit = Environment.Is64BitOperatingSystem;
                Logger.Info($"应用程序架构检测: 进程={(!isApp64Bit ? "32位" : "64位")}, 操作系统={(!isOS64Bit ? "32位" : "64位")}");

                // 根据应用程序架构和系统架构选择合适的驱动程序
                var connectionAttempts = GetConnectionAttempts(databasePath, isApp64Bit, isOS64Bit);

                Exception lastException = null;
                bool connected = false;
                string successfulProvider = "";

                foreach (var attempt in connectionAttempts)
                {
                    try
                    {
                        Logger.Info($"尝试使用 {attempt.Provider} ({attempt.Architecture}) 连接数据库...");
                        
                        using (var connection = new OleDbConnection(attempt.ConnectionString))
                        {
                            connection.Open();
                            
                            // 测试查询以验证连接
                            using (var command = new OleDbCommand($"SELECT TOP 1 * FROM [{tableName}]", connection))
                            {
                                using (var reader = command.ExecuteReader())
                                {
                                    // 连接成功
                                    _connectionString = attempt.ConnectionString;
                                    connected = true;
                                    successfulProvider = $"{attempt.Provider} ({attempt.Architecture})";
                                    Logger.Info($"数据库连接成功，使用提供程序: {successfulProvider}");
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        Logger.Warning($"使用 {attempt.Provider} ({attempt.Architecture}) 连接失败: {ex.Message}");
                        continue;
                    }
                }

                if (!connected)
                {
                    string detailedError = GenerateDetailedErrorMessage(lastException, isApp64Bit, isOS64Bit);
                    throw new Exception(detailedError);
                }

                // 获取最后记录ID
                GetLastRecordId(tableName, monitorField);

                // 确保TR_Print列存在
                EnsureTRPrintColumn(tableName);

                StatusChanged?.Invoke(this, $"数据库连接成功 - {successfulProvider}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"数据库连接失败: {ex.Message}");
                MonitoringError?.Invoke(this, ex.Message);
                return false;
            }
        }

        private static (string Provider, string ConnectionString, string Architecture)[] GetConnectionAttempts(
            string databasePath, bool isApp64Bit, bool isOS64Bit)
        {
            var attempts = new List<(string Provider, string ConnectionString, string Architecture)>();

            if (isApp64Bit)
            {
                // 64位应用程序 - 优先使用64位驱动
                attempts.AddRange(new[]
                {
                    ("Microsoft.ACE.OLEDB.16.0", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};", "64位"),
                    ("Microsoft.ACE.OLEDB.12.0", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};", "64位"),
                });

                // 如果是64位系统，可能还安装了32位Office，但64位应用无法直接使用
                if (isOS64Bit)
                {
                    Logger.Info("64位应用程序无法使用32位Office驱动，跳过32位驱动测试");
                }
            }
            else
            {
                // 32位应用程序 - 可以使用32位驱动
                attempts.AddRange(new[]
                {
                    ("Microsoft.ACE.OLEDB.16.0", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};", "32位"),
                    ("Microsoft.ACE.OLEDB.12.0", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};", "32位"),
                    ("Microsoft.Jet.OLEDB.4.0", $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath};", "32位")
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
            if (_isMonitoring)
            {
                Logger.Warning("监控已在运行中");
                return;
            }

            if (string.IsNullOrEmpty(_connectionString))
            {
                Logger.Error("数据库未连接，无法开始监控");
                MonitoringError?.Invoke(this, "数据库未连接");
                return;
            }

            _monitorTimer.Interval = pollInterval;
            _monitorTimer.Start();
            _isMonitoring = true;
            _retryCount = 0;

            Logger.Info($"开始监控数据库，轮询间隔: {pollInterval}ms");
            StatusChanged?.Invoke(this, "监控已启动");
        }

        public void StopMonitoring()
        {
            if (!_isMonitoring) return;

            _monitorTimer.Stop();
            _isMonitoring = false;

            Logger.Info("数据库监控已停止");
            StatusChanged?.Invoke(this, "监控已停止");
        }

        private void OnTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            try
            {
                CheckForNewRecords();
                _retryCount = 0; // 重置重试计数
            }
            catch (Exception ex)
            {
                _retryCount++;
                Logger.Error($"监控检查失败 (尝试 {_retryCount}/{MaxRetries}): {ex.Message}", ex);

                if (_retryCount >= MaxRetries)
                {
                    Logger.Error("达到最大重试次数，停止监控");
                    StopMonitoring();
                    MonitoringError?.Invoke(this, $"监控失败，已停止: {ex.Message}");
                }
            }
        }

        private void CheckForNewRecords()
        {
            var config = ConfigurationManager.Config.Database;
            
            using var connection = new OleDbConnection(_connectionString);
            connection.Open();

            var query = $"SELECT * FROM [{config.TableName}] WHERE [{config.MonitorField}] > ? ORDER BY [{config.MonitorField}] ASC";
            using var command = new OleDbCommand(query, connection);
            command.Parameters.AddWithValue("?", _lastRecordId);

            using var reader = command.ExecuteReader();
            var newRecords = new List<TestRecord>();

            while (reader.Read())
            {
                var record = MapReaderToTestRecord(reader);
                newRecords.Add(record);
                _lastRecordId = record.TR_SerialNum ?? record.TR_ID ?? "";
            }

            foreach (var record in newRecords)
            {
                Logger.Info($"发现新记录: {record.TR_SerialNum}");
                NewRecordFound?.Invoke(this, record);
            }
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

        private void GetLastRecordId(string tableName, string monitorField)
        {
            try
            {
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                var query = $"SELECT MAX([{monitorField}]) FROM [{tableName}]";
                using var command = new OleDbCommand(query, connection);
                var result = command.ExecuteScalar();

                _lastRecordId = result?.ToString() ?? "";
                Logger.Info($"获取最后记录ID: {_lastRecordId}");
            }
            catch (Exception ex)
            {
                Logger.Error($"获取最后记录ID失败: {ex.Message}", ex);
                _lastRecordId = "";
            }
        }

        private void EnsureTRPrintColumn(string tableName)
        {
            try
            {
                var config = ConfigurationManager.Config;
                
                // 检查是否启用打印次数功能
                if (!config.Database.EnablePrintCount)
                {
                    Logger.Info("打印次数统计已禁用，跳过TR_Print列检查");
                    return;
                }
                
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                // 检查TR_Print列是否存在
                var schemaTable = connection.GetSchema("Columns", new[] { null, null, tableName, null });
                bool hasTRPrint = false;

                foreach (DataRow row in schemaTable.Rows)
                {
                    if (row["COLUMN_NAME"].ToString() == "TR_Print")
                    {
                        hasTRPrint = true;
                        break;
                    }
                }

                if (!hasTRPrint)
                {
                    Logger.Info("添加TR_Print列");
                    var alterQuery = $"ALTER TABLE [{tableName}] ADD TR_Print INTEGER";
                    using var command = new OleDbCommand(alterQuery, connection);
                    command.ExecuteNonQuery();

                    // 设置默认值
                    var updateQuery = $"UPDATE [{tableName}] SET TR_Print = 0 WHERE TR_Print IS NULL";
                    using var updateCommand = new OleDbCommand(updateQuery, connection);
                    updateCommand.ExecuteNonQuery();

                    Logger.Info("TR_Print列添加完成");
                }
                else
                {
                    Logger.Info("TR_Print列已存在");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"TR_Print列检查/添加失败: {ex.Message}", ex);
            }
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

                var updateQuery = $"UPDATE [{config.Database.TableName}] SET TR_Print = IIF(IsNull(TR_Print), 0, TR_Print) + 1 WHERE [{config.Database.MonitorField}] = ?";
                using var command = new OleDbCommand(updateQuery, connection);
                command.Parameters.AddWithValue("?", recordId);

                var affectedRows = command.ExecuteNonQuery();
                Logger.Info($"更新打印计数: 记录ID={recordId}, 影响行数={affectedRows}");

                return affectedRows > 0;
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印计数失败: {ex.Message}", ex);
                return false;
            }
        }

        public List<TestRecord> GetRecentRecords(int limit = 10)
        {
            var records = new List<TestRecord>();

            try
            {
                var config = ConfigurationManager.Config.Database;
                
                using var connection = new OleDbConnection(_connectionString);
                connection.Open();

                var query = $"SELECT TOP {limit} * FROM [{config.TableName}] ORDER BY TR_DateTime DESC";
                using var command = new OleDbCommand(query, connection);
                using var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    records.Add(MapReaderToTestRecord(reader));
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"获取最近记录失败: {ex.Message}", ex);
            }

            return records;
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
    }
} 