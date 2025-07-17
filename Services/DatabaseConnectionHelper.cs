using System;
using System.IO;
using System.Data.OleDb;
using ZebraPrinterMonitor.Utils;
using System.Threading.Tasks; // Added for Task

namespace ZebraPrinterMonitor.Services
{
    public static class DatabaseConnectionHelper
    {
        public static (bool Success, string Message) DiagnoseConnection(string databasePath, string tableName = "TestRecord")
        {
            try
            {
                Logger.Info("🔍 开始数据库连接诊断...");
                
                // 1. 检查文件存在性
                if (string.IsNullOrEmpty(databasePath))
                {
                    return (false, "❌ 数据库路径为空！\n\n解决方案：\n1. 点击'浏览'按钮选择数据库文件\n2. 确保选择的是 .mdb 或 .accdb 文件");
                }
                
                if (!File.Exists(databasePath))
                {
                    return (false, $"❌ 数据库文件不存在！\n路径: {databasePath}\n\n解决方案：\n1. 检查文件路径是否正确\n2. 确认文件没有被移动或删除\n3. 重新选择正确的数据库文件");
                }
                
                // 2. 检查文件权限和占用状态
                try
                {
                    using (var fs = new FileStream(databasePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        Logger.Info($"✅ 文件可访问，大小: {fs.Length} 字节");
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    return (false, "❌ 没有访问数据库文件的权限！\n\n解决方案：\n1. 以管理员身份运行程序\n2. 检查文件是否被其他程序占用\n3. 修改文件权限设置");
                }
                catch (IOException ioEx)
                {
                    return (false, $"❌ 文件访问错误：{ioEx.Message}\n\n解决方案：\n1. 关闭可能占用文件的程序（如Access）\n2. 检查文件是否损坏\n3. 重启计算机后重试");
                }
                
                // 3. 检查应用程序架构
                bool isApp64Bit = Environment.Is64BitProcess;
                bool isOS64Bit = Environment.Is64BitOperatingSystem;
                Logger.Info($"应用架构: {(isApp64Bit ? "64位" : "32位")}, 系统架构: {(isOS64Bit ? "64位" : "32位")}");
                
                // 4. 尝试连接数据库
                var providers = GetProviderAttempts(databasePath, isApp64Bit);
                Exception lastException = null;
                
                foreach (var provider in providers)
                {
                    try
                    {
                        Logger.Info($"尝试连接: {provider.Name}");
                        using (var connection = new OleDbConnection(provider.ConnectionString))
                        {
                            connection.Open();
                            
                            // 测试表访问
                            using (var command = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", connection))
                            {
                                var count = command.ExecuteScalar();
                                return (true, $"✅ 数据库连接成功！\n\n连接信息：\n• 提供程序: {provider.Name}\n• 表 [{tableName}] 记录数: {count}\n• 应用程序: {(isApp64Bit ? "64位" : "32位")}\n• 系统: {(isOS64Bit ? "64位" : "32位")}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        Logger.Warning($"提供程序 {provider.Name} 失败: {ex.Message}");
                        continue;
                    }
                }
                
                // 所有提供程序都失败了
                return (false, GenerateConnectionErrorSolution(lastException, isApp64Bit, isOS64Bit));
            }
            catch (Exception ex)
            {
                Logger.Error($"诊断过程出错: {ex.Message}", ex);
                return (false, $"❌ 诊断过程出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 简化的异步数据库连接测试（借鉴AccessDatabaseMonitor项目）
        /// </summary>
        public static async Task<(bool Success, string Message, string ConnectionString)> TestConnectionSimpleAsync(string databasePath)
        {
            try
            {
                Logger.Info($"🔍 开始简化连接测试: {databasePath}");

                if (!File.Exists(databasePath))
                {
                    var error = $"数据库文件不存在: {databasePath}";
                    Logger.Error(error);
                    return (false, error, "");
                }

                // 使用AccessDatabaseMonitor项目的简化连接字符串，支持并发访问
                var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Mode=Share Deny None;Persist Security Info=false;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Lock Retry=1;Jet OLEDB:Lock Delay=100;";
                
                Logger.Info($"🔗 测试连接字符串: {connectionString}");

                using var connection = new OleDbConnection(connectionString);
                await connection.OpenAsync();
                
                // 测试查询表结构
                using var command = new OleDbCommand("SELECT TOP 1 * FROM TestRecord", connection);
                using var reader = await command.ExecuteReaderAsync();
                
                Logger.Info("✅ 简化连接测试成功！");
                return (true, "数据库连接成功", connectionString);
            }
            catch (Exception ex)
            {
                var errorMsg = $"简化连接测试失败: {ex.Message}";
                Logger.Error(errorMsg, ex);
                
                // 如果简化方式失败，尝试回退到原有的复杂检测方式
                Logger.Info("🔄 回退到完整驱动检测方式...");
                var diagnosis = DiagnoseConnection(databasePath);
                return (diagnosis.Success, diagnosis.Message, "");
            }
        }

        private static (string Name, string ConnectionString)[] GetProviderAttempts(string databasePath, bool isApp64Bit)
        {
            // Access数据库并发访问优化的连接字符串参数
            string commonParams = "Mode=Share Deny None;Persist Security Info=false;Jet OLEDB:Database Locking Mode=1;";
            
            if (isApp64Bit)
            {
                return new[]
                {
                    ("Microsoft.ACE.OLEDB.16.0 (64位-并发)", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};{commonParams}"),
                    ("Microsoft.ACE.OLEDB.12.0 (64位-并发)", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};{commonParams}"),
                    ("Microsoft.ACE.OLEDB.16.0 (64位-标准)", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};"),
                    ("Microsoft.ACE.OLEDB.12.0 (64位-标准)", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};")
                };
            }
            else
            {
                return new[]
                {
                    ("Microsoft.ACE.OLEDB.16.0 (32位-并发)", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};{commonParams}"),
                    ("Microsoft.ACE.OLEDB.12.0 (32位-并发)", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};{commonParams}"),
                    ("Microsoft.Jet.OLEDB.4.0 (32位-并发)", $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath};{commonParams}Jet OLEDB:Engine Type=5;"),
                    ("Microsoft.ACE.OLEDB.16.0 (32位-标准)", $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};"),
                    ("Microsoft.ACE.OLEDB.12.0 (32位-标准)", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};"),
                    ("Microsoft.Jet.OLEDB.4.0 (32位-标准)", $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath};")
                };
            }
        }
        
        private static string GenerateConnectionErrorSolution(Exception lastException, bool isApp64Bit, bool isOS64Bit)
        {
            var message = "❌ 数据库连接失败：Access数据库引擎架构不匹配！\n\n";
            
            message += $"🔍 诊断信息：\n";
            message += $"• 当前应用程序：{(isApp64Bit ? "64位" : "32位")}\n";
            message += $"• 操作系统：{(isOS64Bit ? "64位" : "32位")}\n";
            message += $"• 错误详情：{lastException?.Message}\n\n";
            
            if (isApp64Bit)
            {
                message += "💡 解决方案（64位应用程序）：\n";
                message += "1. 【推荐】安装64位Microsoft Access数据库引擎：\n";
                message += "   • 下载地址：https://www.microsoft.com/zh-cn/download/details.aspx?id=54920\n";
                message += "   • 选择：AccessDatabaseEngine_X64.exe\n";
                message += "   • 如果已安装Office，使用命令：/passive 参数强制安装\n\n";
                message += "2. 或者使用32位版本的程序（如果已安装32位Office）\n\n";
            }
            else
            {
                message += "💡 解决方案（32位应用程序）：\n";
                message += "1. 安装32位Microsoft Access数据库引擎：\n";
                message += "   • 下载地址：https://www.microsoft.com/zh-cn/download/details.aspx?id=54920\n";
                message += "   • 选择：AccessDatabaseEngine.exe（32位版本）\n";
                message += "   • 如果已安装Office，使用命令：/passive 参数强制安装\n\n";
            }
            
            message += "📝 安装步骤：\n";
            message += "1. 关闭所有Office程序和此应用程序\n";
            message += "2. 下载并安装对应架构的Access数据库引擎\n";
            message += "3. 重启计算机\n";
            message += "4. 重新启动此程序";
            
            return message;
        }
    }
} 