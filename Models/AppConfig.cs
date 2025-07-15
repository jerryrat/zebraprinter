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
        public string Version { get; set; } = "1.1.41";
        public DatabaseConfig Database { get; set; } = new();
        public PrinterConfig Printer { get; set; } = new();
        public ApplicationConfig Application { get; set; } = new();
        public UIConfig UI { get; set; } = new();
    }

    public class DatabaseConfig
    {
        public string DatabasePath { get; set; } = "";
        public string TableName { get; set; } = "TestRecord";
        public string MonitorField { get; set; } = "TR_SerialNum";
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
    }

    public class ApplicationConfig
    {
        public string LogLevel { get; set; } = "Info";
        public bool StartMinimized { get; set; } = false;
        public bool MinimizeToTray { get; set; } = true;
        public bool AutoStartMonitoring { get; set; } = false;
    }

    public class UIConfig
    {
        public string Language { get; set; } = "zh-CN";
        public string Theme { get; set; } = "Default";
        public int WindowWidth { get; set; } = 1200;
        public int WindowHeight { get; set; } = 800;
    }

    public class TestRecord
    {
        public string? TR_ID { get; set; }
        public string? TR_SerialNum { get; set; }
        public decimal? TR_Isc { get; set; }
        public decimal? TR_Voc { get; set; }
        public decimal? TR_Pm { get; set; }
        public decimal? TR_Ipm { get; set; }
        public decimal? TR_Vpm { get; set; }
        public decimal? TR_CellEfficiency { get; set; }
        public decimal? TR_FF { get; set; }
        public string? TR_Grade { get; set; }
        public decimal? TR_Temp { get; set; }
        public decimal? TR_Irradiance { get; set; }
        public decimal? TR_Rs { get; set; }
        public decimal? TR_Rsh { get; set; }
        public string? TR_CellArea { get; set; }
        public string? TR_Operater { get; set; }
        public DateTime? TR_DateTime { get; set; }
        public int? TR_Print { get; set; }
        public string? TR_FontColor { get; set; }
        public string? TR_BackColor { get; set; }

        public string GetDisplayText()
        {
            return $"序列号: {TR_SerialNum ?? "N/A"} - 日期: {TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A"}";
        }

        public string FormatNumber(decimal? value)
        {
            if (value == null) return "N/A";
            return value.Value.ToString("F3");
        }
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