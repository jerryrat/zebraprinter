using System;
using System.IO;

namespace ZebraPrinterMonitor.Utils
{
    public static class Logger
    {
        private static string? _logFilePath;
        private static readonly object _lock = new object();

        public static void Initialize()
        {
            try
            {
                var logsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                if (!Directory.Exists(logsDir))
                {
                    Directory.CreateDirectory(logsDir);
                }

                _logFilePath = Path.Combine(logsDir, $"app_{DateTime.Now:yyyyMMdd}.log");
                
                // 写入启动日志
                WriteLog("INFO", "日志系统初始化完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"日志初始化失败: {ex.Message}");
            }
        }

        public static void Info(string message)
        {
            WriteLog("INFO", message);
        }

        public static void Error(string message, Exception? exception = null)
        {
            var fullMessage = exception != null ? $"{message}\n异常详情: {exception}" : message;
            WriteLog("ERROR", fullMessage);
        }

        public static void Warning(string message)
        {
            WriteLog("WARN", message);
        }

        public static void Debug(string message)
        {
            WriteLog("DEBUG", message);
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                if (string.IsNullOrEmpty(_logFilePath)) return;

                lock (_lock)
                {
                    var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    var logEntry = $"[{timestamp}] [{level}] {message}";
                    
                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                    
                    // 同时输出到控制台（调试时有用）
                    Console.WriteLine(logEntry);
                }
            }
            catch
            {
                // 忽略日志写入失败，避免影响主程序
            }
        }

        public static void Shutdown()
        {
            WriteLog("INFO", "应用程序关闭");
        }
    }
} 