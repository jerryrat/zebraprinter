using System;
using System.Collections.Generic;
using System.Globalization;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Services
{
    public static class LanguageManager
    {
        private static Dictionary<string, Dictionary<string, string>> _resources;
        private static string _currentLanguage = "zh-CN"; // 默认简体中文

        static LanguageManager()
        {
            InitializeResources();
        }

        public static string CurrentLanguage 
        { 
            get => _currentLanguage; 
            set 
            { 
                if (_resources.ContainsKey(value))
                {
                    _currentLanguage = value;
                    Logger.Info($"语言已切换到: {value}");
                }
            } 
        }

        public static List<string> SupportedLanguages => new List<string> { "zh-CN", "en-US" };

        public static string GetLanguageName(string code)
        {
            return code switch
            {
                "zh-CN" => "简体中文",
                "en-US" => "English",
                _ => code
            };
        }

        public static string GetString(string key, string defaultValue = "")
        {
            if (_resources.ContainsKey(_currentLanguage) && 
                _resources[_currentLanguage].ContainsKey(key))
            {
                return _resources[_currentLanguage][key];
            }

            // 如果当前语言没有找到，尝试英文
            if (_currentLanguage != "en-US" && 
                _resources.ContainsKey("en-US") && 
                _resources["en-US"].ContainsKey(key))
            {
                return _resources["en-US"][key];
            }

            return defaultValue ?? key;
        }

        private static void InitializeResources()
        {
            _resources = new Dictionary<string, Dictionary<string, string>>
            {
                ["zh-CN"] = new Dictionary<string, string>
                {
                    // 主窗体
                    ["MainTitle"] = "太阳能电池测试打印监控系统",
                    ["TabDataMonitoring"] = "数据监控",
                    ["TabSystemConfig"] = "系统配置", 
                    ["TabRuntimeLogs"] = "运行日志",
                    ["TabPrintTemplate"] = "打印模板",
                    
                    // 数据监控页面
                    ["SerialNumber"] = "序列号",
                    ["TestDateTime"] = "测试时间",
                    ["Current"] = "电流(A)",
                    ["Voltage"] = "电压(V)",
                    ["VoltageVpm"] = "Vpm电压(V)",
                    ["Power"] = "功率(W)",
                    ["PrintCount"] = "打印次数",
                    ["Operation"] = "操作",
                    ["DoubleClickToPrint"] = "双击打印",
                    ["Print"] = "打印",
                    ["ViewDetails"] = "查看详情",
                    
                    // 系统配置页面
                    ["DatabasePath"] = "数据库路径:",
                    ["SelectDatabase"] = "选择数据库",
                    ["SelectedPrinter"] = "选择打印机:",
                    ["AutoStartMonitoring"] = "启动时自动开始监控",
                    ["MinimizeToTray"] = "最小化到系统托盘",
                    ["PollInterval"] = "轮询间隔(毫秒):",
                    ["StartMonitoring"] = "开始监控",
                    ["StopMonitoring"] = "停止监控",
                    ["SaveConfiguration"] = "保存配置",
                    ["Language"] = "语言:",
                    
                    // 打印模板页面
                    ["PrintFormat"] = "打印格式:",
                    ["TemplateContent"] = "模板内容:",
                    ["AvailableFields"] = "可用字段:",
                    ["FieldDescription"] = "字段说明:",
                    ["PreviewTemplate"] = "预览模板",
                    ["SaveTemplate"] = "保存模板",
                    ["ResetTemplate"] = "重置模板",
                    
                    // 按钮和消息
                    ["OK"] = "确定",
                    ["Cancel"] = "取消",
                    ["Browse"] = "浏览",
                    ["Exit"] = "退出",
                    ["ShowMainWindow"] = "显示主界面",
                    ["ExitProgram"] = "退出程序",
                    
                    // 状态信息
                    ["Status"] = "状态:",
                    ["Ready"] = "就绪",
                    ["Monitoring"] = "监控中",
                    ["Stopped"] = "已停止",
                    ["Connected"] = "已连接",
                    ["Disconnected"] = "未连接",
                    ["TotalRecords"] = "处理记录数:",
                    ["TotalPrintJobs"] = "打印任务数:",
                    
                    // 日志信息
                    ["LogLevel"] = "日志级别",
                    ["LogTime"] = "时间",
                    ["LogMessage"] = "消息",
                    ["ClearLogs"] = "清空日志",
                    
                    // 错误和提示消息
                    ["Error"] = "错误",
                    ["Warning"] = "警告",
                    ["Info"] = "信息",
                    ["Success"] = "成功",
                    ["DatabaseConnectionFailed"] = "数据库连接失败",
                    ["ConfigurationSaved"] = "配置已保存",
                    ["MonitoringStarted"] = "监控已开始",
                    ["MonitoringStopped"] = "监控已停止",
                    ["PrintCompleted"] = "打印完成",
                    ["PrintFailed"] = "打印失败",
                    ["TemplateFields"] = "模板字段说明:\n{SerialNumber} - 序列号\n{TestDateTime} - 测试时间\n{Current} - 电流\n{Voltage} - 电压\n{VoltageVpm} - Vpm电压\n{Power} - 功率\n{PrintCount} - 打印次数",
                    
                    // 打印机相关提示
                    ["NoPrinterFound"] = "系统中没有找到任何打印机，请先安装打印机。",
                    ["NoPrinterInstalled"] = "系统中没有安装打印机。\n\n请按照以下步骤安装打印机：\n1. 打开\"设置\" > \"打印机和扫描仪\"\n2. 点击\"添加打印机或扫描仪\"\n3. 选择您的打印机并按照提示完成安装\n\n安装完成后，请重新启动本软件。",
                    ["NoPrinterTitle"] = "未安装打印机",
                    ["ReprintWarningTitle"] = "重复打印确认",
                    ["ReprintWarningMessage"] = "此记录已打印 {0} 次。\n\n确认要继续打印吗？",
                    ["Confirm"] = "确认",
                    ["PrintAgain"] = "继续打印"
                },
                
                ["en-US"] = new Dictionary<string, string>
                {
                    // Main Form
                    ["MainTitle"] = "Solar Cell Test Printing Monitor System",
                    ["TabDataMonitoring"] = "Data Monitoring",
                    ["TabSystemConfig"] = "System Configuration",
                    ["TabRuntimeLogs"] = "Runtime Logs",
                    ["TabPrintTemplate"] = "Print Template",
                    
                    // Data Monitoring Tab
                    ["SerialNumber"] = "Serial Number",
                    ["TestDateTime"] = "Test Date Time",
                    ["Current"] = "Current(A)",
                    ["Voltage"] = "Voltage(V)",
                    ["VoltageVpm"] = "Vpm Voltage(V)",
                    ["Power"] = "Power(W)",
                    ["PrintCount"] = "Print Count",
                    ["Operation"] = "Operation",
                    ["DoubleClickToPrint"] = "Double Click to Print",
                    ["Print"] = "Print",
                    ["ViewDetails"] = "View Details",
                    
                    // System Configuration Tab
                    ["DatabasePath"] = "Database Path:",
                    ["SelectDatabase"] = "Select Database",
                    ["SelectedPrinter"] = "Selected Printer:",
                    ["AutoStartMonitoring"] = "Auto Start Monitoring",
                    ["MinimizeToTray"] = "Minimize to System Tray",
                    ["PollInterval"] = "Poll Interval(ms):",
                    ["StartMonitoring"] = "Start Monitoring",
                    ["StopMonitoring"] = "Stop Monitoring",
                    ["SaveConfiguration"] = "Save Configuration",
                    ["Language"] = "Language:",
                    
                    // Print Template Tab
                    ["PrintFormat"] = "Print Format:",
                    ["TemplateContent"] = "Template Content:",
                    ["AvailableFields"] = "Available Fields:",
                    ["FieldDescription"] = "Field Description:",
                    ["PreviewTemplate"] = "Preview Template",
                    ["SaveTemplate"] = "Save Template",
                    ["ResetTemplate"] = "Reset Template",
                    
                    // Buttons and Messages
                    ["OK"] = "OK",
                    ["Cancel"] = "Cancel",
                    ["Browse"] = "Browse",
                    ["Exit"] = "Exit",
                    ["ShowMainWindow"] = "Show Main Window",
                    ["ExitProgram"] = "Exit Program",
                    
                    // Status Information
                    ["Status"] = "Status:",
                    ["Ready"] = "Ready",
                    ["Monitoring"] = "Monitoring",
                    ["Stopped"] = "Stopped",
                    ["Connected"] = "Connected",
                    ["Disconnected"] = "Disconnected",
                    ["TotalRecords"] = "Total Records:",
                    ["TotalPrintJobs"] = "Total Print Jobs:",
                    
                    // Log Information
                    ["LogLevel"] = "Log Level",
                    ["LogTime"] = "Time",
                    ["LogMessage"] = "Message",
                    ["ClearLogs"] = "Clear Logs",
                    
                    // Error and Info Messages
                    ["Error"] = "Error",
                    ["Warning"] = "Warning",
                    ["Info"] = "Information",
                    ["Success"] = "Success",
                    ["DatabaseConnectionFailed"] = "Database connection failed",
                    ["ConfigurationSaved"] = "Configuration saved",
                    ["MonitoringStarted"] = "Monitoring started",
                    ["MonitoringStopped"] = "Monitoring stopped",
                    ["PrintCompleted"] = "Print completed",
                    ["PrintFailed"] = "Print failed",
                    ["TemplateFields"] = "Template Fields:\n{SerialNumber} - Serial Number\n{TestDateTime} - Test Date Time\n{Current} - Current\n{Voltage} - Voltage\n{VoltageVpm} - Vpm Voltage\n{Power} - Power\n{PrintCount} - Print Count",
                    
                    // Printer related messages
                    ["NoPrinterFound"] = "No printer found in the system. Please install a printer first.",
                    ["NoPrinterInstalled"] = "No printer is installed in the system.\n\nPlease follow these steps to install a printer:\n1. Open \"Settings\" > \"Printers & scanners\"\n2. Click \"Add a printer or scanner\"\n3. Select your printer and follow the prompts to complete installation\n\nPlease restart this software after installation.",
                    ["NoPrinterTitle"] = "No Printer Installed",
                    ["ReprintWarningTitle"] = "Reprint Confirmation",
                    ["ReprintWarningMessage"] = "This record has been printed {0} times.\n\nDo you want to continue printing?",
                    ["Confirm"] = "Confirm",
                    ["PrintAgain"] = "Print Again"
                }
            };
        }
    }
} 