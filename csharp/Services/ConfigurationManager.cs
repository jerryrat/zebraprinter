using System;
using System.IO;
using Newtonsoft.Json;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Services
{
    public static class ConfigurationManager
    {
        private static AppConfig? _config;
        private static readonly string ConfigFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

        public static AppConfig Config => _config ?? throw new InvalidOperationException("配置未初始化");

        public static void Initialize()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var jsonContent = File.ReadAllText(ConfigFilePath);
                    _config = JsonConvert.DeserializeObject<AppConfig>(jsonContent) ?? GetDefaultConfig();
                    Logger.Info($"配置文件加载成功: {ConfigFilePath}");
                }
                else
                {
                    _config = GetDefaultConfig();
                    SaveConfig();
                    Logger.Info($"创建默认配置文件: {ConfigFilePath}");
                }

                // 验证配置
                ValidateConfig();
            }
            catch (Exception ex)
            {
                Logger.Error($"配置初始化失败: {ex.Message}", ex);
                _config = GetDefaultConfig();
            }
        }

        public static void SaveConfig()
        {
            try
            {
                if (_config == null) return;

                var jsonContent = JsonConvert.SerializeObject(_config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, jsonContent);
                Logger.Info("配置文件保存成功");
            }
            catch (Exception ex)
            {
                Logger.Error($"配置文件保存失败: {ex.Message}", ex);
                throw;
            }
        }

        public static void UpdateDatabaseConfig(DatabaseConfig databaseConfig)
        {
            if (_config == null) return;

            _config.Database = databaseConfig;
            SaveConfig();
            Logger.Info($"数据库配置已更新: {databaseConfig.DatabasePath}");
        }

        public static void UpdatePrinterConfig(PrinterConfig printerConfig)
        {
            if (_config == null) return;

            _config.Printer = printerConfig;
            SaveConfig();
            Logger.Info($"打印机配置已更新: {printerConfig.PrinterName}");
        }

        private static AppConfig GetDefaultConfig()
        {
            return new AppConfig
            {
                Database = new DatabaseConfig
                {
                    DatabasePath = "",
                    TableName = "TestRecord",
                    MonitorField = "TR_SerialNum",
                    PollInterval = 1000
                },
                Printer = new PrinterConfig
                {
                    PrinterName = "Microsoft Print to PDF",
                    PrinterType = "Text",
                    AutoPrint = true,
                    PrintFormat = "Text"
                },
                Application = new ApplicationConfig
                {
                    LogLevel = "Info",
                    StartMinimized = false,
                    MinimizeToTray = true,
                    AutoStartMonitoring = false
                },
                Version = "1.1.31"
            };
        }

        private static void ValidateConfig()
        {
            if (_config == null) return;

            // 验证数据库配置
            if (string.IsNullOrEmpty(_config.Database.TableName))
            {
                _config.Database.TableName = "TestRecord";
            }

            if (string.IsNullOrEmpty(_config.Database.MonitorField))
            {
                _config.Database.MonitorField = "TR_SerialNum";
            }

            if (_config.Database.PollInterval <= 0)
            {
                _config.Database.PollInterval = 1000;
            }

            // 验证打印机配置
            if (string.IsNullOrEmpty(_config.Printer.PrinterName))
            {
                _config.Printer.PrinterName = "Microsoft Print to PDF";
            }

            if (string.IsNullOrEmpty(_config.Printer.PrintFormat))
            {
                _config.Printer.PrintFormat = "Text";
            }

            Logger.Info("配置验证完成");
        }
    }
} 