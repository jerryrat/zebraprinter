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
        public string Version { get; set; } = "1.1.30";
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

        public string FormatNumber(decimal? value, int decimals = 2)
        {
            if (value == null) return "N/A";
            return value.Value.ToString($"F{decimals}");
        }

        /// <summary>
        /// 验证记录数据的完整性
        /// </summary>
        public bool IsValid()
        {
            return !string.IsNullOrEmpty(TR_SerialNum) &&
                   TR_DateTime.HasValue &&
                   TR_Isc.HasValue &&
                   TR_Ipm.HasValue &&
                   TR_Voc.HasValue &&
                   TR_Vpm.HasValue &&
                   TR_Pm.HasValue;
        }

        /// <summary>
        /// 创建示例测试记录
        /// </summary>
        public static TestRecord CreateSample()
        {
            return new TestRecord
            {
                TR_SerialNum = "SKT600M12120HB-" + DateTime.Now.ToString("yyyyMMddHHmmss"),
                TR_DateTime = DateTime.Now,
                TR_Isc = 11.25m,
                TR_Ipm = 10.89m,
                TR_Voc = 49.8m,
                TR_Vpm = 41.2m,
                TR_Pm = 448.7m,
                TR_Print = 1
            };
        }

        /// <summary>
        /// 从数据库行创建TestRecord
        /// </summary>
        public static TestRecord FromDataRow(System.Data.DataRow row)
        {
            var record = new TestRecord();
            
            try
            {
                record.TR_SerialNum = row["TR_SerialNum"]?.ToString() ?? "";
                
                if (DateTime.TryParse(row["TR_DateTime"]?.ToString(), out DateTime dateTime))
                    record.TR_DateTime = dateTime;
                
                if (decimal.TryParse(row["TR_Isc"]?.ToString(), out decimal isc))
                    record.TR_Isc = isc;
                
                if (decimal.TryParse(row["TR_Ipm"]?.ToString(), out decimal ipm))
                    record.TR_Ipm = ipm;
                
                if (decimal.TryParse(row["TR_Voc"]?.ToString(), out decimal voc))
                    record.TR_Voc = voc;
                
                if (decimal.TryParse(row["TR_Vpm"]?.ToString(), out decimal vpm))
                    record.TR_Vpm = vpm;
                
                if (decimal.TryParse(row["TR_Pm"]?.ToString(), out decimal pm))
                    record.TR_Pm = pm;
                
                if (int.TryParse(row["TR_Print"]?.ToString(), out int print))
                    record.TR_Print = print;
            }
            catch (Exception ex)
            {
                // 注意：这里无法使用Logger，因为可能会导致循环引用
                System.Diagnostics.Debug.WriteLine($"从数据行创建TestRecord失败: {ex.Message}");
            }
            
            return record;
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