using System;

namespace ZebraPrinterMonitor.Models
{
    public enum PrintFormat
    {
        Text = 0,
        ZPL = 1,
        Code128 = 2,
        QRCode = 3
    }

    public class AppConfig
    {
        public string Version { get; set; } = "1.3.9.8";
        public DatabaseConfig Database { get; set; } = new();
        public PrinterConfig Printer { get; set; } = new();
        public ApplicationConfig Application { get; set; } = new();
        public UIConfig UI { get; set; } = new();

        public void Initialize()
        {
            Database = new DatabaseConfig
            {
                DatabasePath = "",
                TableName = "TestRecord",
                MonitorField = "TR_ID",  // 使用主键作为默认监控字段
                PollInterval = 1000,
                EnablePrintCount = false
            };

            Application = new ApplicationConfig
            {
                LogLevel = "Info",
                StartMinimized = false,
                MinimizeToTray = true,
                AutoStartMonitoring = true  // 默认启用自动开始监控
            };

            Printer = new PrinterConfig
            {
                PrinterName = "",
                PrintFormat = "Text",
                AutoPrint = true,
                DefaultTemplate = "Default",
                EnablePrintCount = true
            };

            UI = new UIConfig
            {
                Language = "zh-CN",
                EnableSystemTray = true,
                StartMinimized = false
            };

            // 设置版本号
            Version = "1.3.9.8";
        }
    }

    public class DatabaseConfig
    {
        public string DatabasePath { get; set; } = "";
        public string TableName { get; set; } = "TestRecord";
        public string MonitorField { get; set; } = "TR_ID";
        public int PollInterval { get; set; } = 1000;
        public bool EnablePrintCount { get; set; } = false;  // 默认不启用打印次数功能
    }

    public class PrinterConfig
    {
        public string PrinterName { get; set; } = "Microsoft Print to PDF";
        public string PrinterType { get; set; } = "Text";
        public bool AutoPrint { get; set; } = true;
        public string PrintFormat { get; set; } = "Text";
        public string DefaultTemplate { get; set; } = "默认文本模板";
        public bool EnablePrintCount { get; set; } = false; // Added this property
    }

    public class ApplicationConfig
    {
        public string LogLevel { get; set; } = "Info";
        public bool StartMinimized { get; set; } = false;
        public bool MinimizeToTray { get; set; } = true;
        public bool AutoStartMonitoring { get; set; } = true;
    }

    public class UIConfig
    {
        public string Language { get; set; } = "zh-CN";
        public string Theme { get; set; } = "Default";
        public int WindowWidth { get; set; } = 1200;
        public int WindowHeight { get; set; } = 800;
        public bool EnableSystemTray { get; set; } = true; // Added this property
        public bool StartMinimized { get; set; } = false; // Added this property
    }

    public class PrintResult
    {
        public bool Success { get; set; }
        public string? Method { get; set; }
        public string? JobId { get; set; }
        public string? PrinterUsed { get; set; }
        public string? ErrorMessage { get; set; }
        public DateTime PrintTime { get; set; } = DateTime.Now;
    }
} 