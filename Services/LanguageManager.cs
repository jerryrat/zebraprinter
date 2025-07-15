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
                    ["DatabaseConfig"] = "数据库配置",
                    ["MonitorControl"] = "监控控制",
                    ["StatusInfo"] = "状态信息",
                    ["RecentRecords"] = "最近记录",
                    ["SerialNumber"] = "序列号",
                    ["TestDateTime"] = "测试时间",
                    ["Current"] = "电流(A)",
                    ["Voltage"] = "电压(V)",
                    ["VoltageVpm"] = "Vpm电压(V)",
                    ["Power"] = "功率(W)",
                    ["PrintCount"] = "打印次数",
                    ["Operation"] = "操作",
                    ["DoubleClickToPrint"] = "双击打印",
                    ["ViewDetails"] = "查看详情",
                    ["PrintSelected"] = "打印选中",
                    ["Refresh"] = "刷新",
                    ["AutoPrint"] = "自动打印",
                    ["EnablePrintCount"] = "启用打印次数统计",
                    ["PrintPreview"] = "打印预览",
                    
                    // 系统配置页面
                    ["DatabasePath"] = "数据库路径:",
                    ["Browse"] = "浏览...",
                    ["TestConnection"] = "测试连接",
                    ["SelectDatabase"] = "选择数据库",
                    ["SelectedPrinter"] = "选择打印机:",
                    ["AutoStartMonitoring"] = "程序启动时自动开始监控",
                    ["MinimizeToTray"] = "最小化到系统托盘",
                    ["PollInterval"] = "轮询间隔 (毫秒):",
                    ["StartMonitoring"] = "开始监控",
                    ["StopMonitoring"] = "停止监控",
                    ["SaveConfiguration"] = "保存配置",
                    ["Language"] = "界面语言:",
                    ["LanguageConfig"] = "语言配置",
                    ["PrinterConfig"] = "打印机配置",
                    ["ApplicationConfig"] = "应用程序配置",
                    ["TestPrint"] = "测试打印",
                    ["PrinterStatus"] = "打印机状态: 未知",
                    ["PrinterStatusOK"] = "打印机状态: 正常",
                    ["PrinterStatusError"] = "打印机状态: 错误",
                    ["PrinterStatusOffline"] = "打印机状态: 离线",
                    ["GetPrinterListFailed"] = "获取打印机列表失败",
                    
                    // 打印模板页面
                    ["PrintFormat"] = "打印格式:",
                    ["TemplateList"] = "模板列表",
                    ["TemplateEditor"] = "模板编辑",
                    ["TemplatePreview"] = "可用字段和预览",
                    ["TemplateName"] = "模板名称:",
                    ["TemplateContent"] = "模板内容:",
                    ["AvailableFields"] = "可用字段:",
                    ["FieldDescription"] = "字段说明:",
                    ["PreviewTemplate"] = "预览模板",
                    ["SaveTemplate"] = "保存模板",
                    ["ResetTemplate"] = "重置模板",
                    ["NewTemplate"] = "新建模板",
                    ["DeleteTemplate"] = "删除模板",
                    ["VisualDesigner"] = "可视化设计器",
                    ["PrePrintedLabelMode"] = "预印刷标签模式",
                    ["FieldPositionSetting"] = "字段位置设置",
                    ["SelectField"] = "选择字段:",
                    ["PosX"] = "X:",
                    ["PosY"] = "Y:",
                    ["Width"] = "宽度:",
                    ["Alignment"] = "对齐:",
                    ["ValueOnly"] = "仅数值",
                    ["AddField"] = "添加字段",
                    ["UpdateField"] = "更新字段",
                    ["DeleteField"] = "删除字段",
                    ["ClearContent"] = "清空内容",
                    ["ImportTemplate"] = "导入模板",
                    ["PreviewLabel"] = "预览:",
                    ["ButtonsInfo"] = "提示：蓝色保存、绿色预览、黄色清空、灰色导入",
                    ["SavingTemplate"] = "保存中...",
                    ["NewTemplateText"] = "新模板",
                    
                    // 运行日志页面
                    ["RuntimeLogs"] = "运行日志",
                    ["ClearLogs"] = "清空日志",
                    ["SaveLogs"] = "保存日志",
                    
                    // 打印预览窗体
                    ["PrintPreviewTitle"] = "🖨️ 打印预览",
                    ["ConfirmPrint"] = "🖨️ 确认打印",
                    ["ShowMainWindow"] = "📋 显示主界面",
                    ["Close"] = "✖ 关闭",
                    ["LoadingContent"] = "正在加载打印内容...",
                    ["NoPreviewData"] = "没有可预览的数据",
                    ["AutoPrintEnabled"] = "自动打印已启用",
                    ["Loading"] = "加载中...",
                    
                    // 可视化设计器窗体
                    ["VisualDesignerTitle"] = "模板可视化设计器",
                    ["AvailableFieldsDesc"] = "可用字段 (只引用数据内容，项目名称请自己输入)",
                    ["CustomTextPlaceholder"] = "输入自定义文本",
                    ["AddCustomText"] = "添加自定义文本",
                    ["ClearDesignArea"] = "清空设计面板",
                    ["DesignArea"] = "设计区域",
                    ["PropertiesAndPreview"] = "属性和预览",
                    ["TemplateNameProp"] = "模板名称:",
                    ["OutputFormat"] = "输出格式:",
                    ["Preview"] = "预览",
                    ["Clear"] = "清空",
                    ["SaveTemplateBtn"] = "保存模板",
                    ["LoadTemplate"] = "加载模板",
                    ["Cancel"] = "取消",
                    ["SelectTemplate"] = "选择模板",
                    ["OK"] = "确定",
                    
                    // 按钮和消息
                    ["Print"] = "打印",
                    ["Exit"] = "退出",
                    ["ExitProgram"] = "退出程序",
                    
                    // 状态信息
                    ["Status"] = "状态:",
                    ["Ready"] = "就绪",
                    ["Monitoring"] = "监控中",
                    ["Stopped"] = "已停止",
                    ["Connected"] = "已连接",
                    ["Disconnected"] = "未连接",
                    ["MonitoringStatus"] = "监控状态:",
                    ["MonitoringStatusStopped"] = "监控状态: 已停止",
                    ["MonitoringStatusRunning"] = "监控状态: 监控中",
                    ["TotalRecords"] = "处理记录: 0",
                    ["TotalPrints"] = "打印任务: 0",
                    ["LastRecord"] = "最后记录: N/A",
                    ["TotalPrintJobs"] = "打印任务数:",
                    
                    // 日志信息
                    ["LogLevel"] = "日志级别",
                    ["LogTime"] = "时间",
                    ["LogMessage"] = "消息",
                    
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
                    
                    // 状态和处理信息
                    ["ProcessedRecords"] = "处理记录",
                    ["PrintJobs"] = "打印任务",
                    ["LastRecord"] = "最后记录",
                    ["CurrentPrinter"] = "当前打印机",
                    ["LanguageSwitched"] = "语言已切换到",
                    ["PollIntervalChanged"] = "轮询间隔已更改为",
                    ["LogSaved"] = "日志已保存到",
                    ["Enabled"] = "启用",
                    ["Disabled"] = "禁用",
                    ["FunctionEnabled"] = "功能启用",
                    ["FunctionDisabled"] = "功能禁用",
                    ["PrintCountStatistic"] = "打印次数统计已",
                    ["PrintCountEnabledMessage"] = "打印次数统计已启用。",
                    ["PrintOperationWillUpdate"] = "新的打印操作将更新数据库中的TR_Print字段。",
                    ["PrintCountDisabledMessage"] = "打印次数统计已禁用。",
                    ["PrintOperationWillNotUpdate"] = "打印操作将不会更新数据库中的TR_Print字段，保持数据库兼容性。",
                    
                    // 模板相关
                    ["TemplateNameRequired"] = "请输入模板名称",
                    ["TemplateSaved"] = "模板已保存",
                    ["TemplateSaveFailed"] = "模板保存失败",
                    ["SelectTemplateToDelete"] = "请选择要删除的模板",
                    ["ConfirmDeleteTemplate"] = "确定要删除模板 '{0}' 吗？",
                    ["TemplateDeleted"] = "模板已删除",
                    ["TemplateDeleteFailed"] = "模板删除失败",
                    ["TemplateImported"] = "模板导入成功",
                    ["ImportTemplateFailed"] = "导入模板失败",
                    ["VisualDesignerError"] = "启动可视化设计器失败",
                    ["SaveTemplateToolTip"] = "保存当前模板 (Ctrl+S)",
                    
                    // 打印相关
                    ["PrintRecordSuccess"] = "成功打印记录",
                    ["PrintRecordFailed"] = "打印记录失败",
                    ["PrintTaskSent"] = "打印任务已发送",
                    ["TestPrintSent"] = "测试打印已发送到",
                    ["TestPrintFailed"] = "测试打印失败",
                    ["TestPrintError"] = "测试打印过程中发生错误",
                    ["PrintError"] = "打印过程中发生错误",
                    ["PrintPreviewError"] = "打印预览失败",
                    ["PrintThisRecord"] = "打印此记录",
                    ["InvalidRecordData"] = "无效的记录数据",
                    
                    // 记录详情
                    ["RecordDetails"] = "记录详细信息",
                    ["NA"] = "N/A",
                    ["ShortCircuitCurrent"] = "短路电流",
                    ["Isc"] = "Isc",
                    ["A"] = "A",
                    ["OpenCircuitVoltage"] = "开路电压",
                    ["Voc"] = "Voc",
                    ["V"] = "V",
                    ["MaximumPowerPointVoltage"] = "最大功率点电压",
                    ["Vpm"] = "Vpm",
                    ["MaximumPowerPointCurrent"] = "最大功率点电流",
                    ["Ipm"] = "Ipm",
                    ["MaximumPower"] = "最大功率",
                    ["Pm"] = "Pm",
                    ["W"] = "W",
                    ["Efficiency"] = "效率",
                    ["FillFactor"] = "填充因子",
                    ["FF"] = "FF",
                    ["Grade"] = "等级",
                    
                    // 打印机相关提示
                    ["NoPrinterFound"] = "系统中没有找到任何打印机，请先安装打印机。",
                    ["NoPrinterInstalled"] = "系统中没有安装打印机。\n\n请按照以下步骤安装打印机：\n1. 打开\"设置\" > \"打印机和扫描仪\"\n2. 点击\"添加打印机或扫描仪\"\n3. 选择您的打印机并按照提示完成安装\n\n安装完成后，请重新启动本软件。",
                    ["NoPrinterTitle"] = "未安装打印机",
                    ["ReprintWarningTitle"] = "重复打印确认",
                    ["ReprintWarningMessage"] = "此记录已打印 {0} 次。\n\n确认要继续打印吗？",
                    ["Confirm"] = "确认",
                    ["PrintAgain"] = "继续打印",
                    
                    // 托盘提示
                    ["TrayNotificationTitle"] = "程序已最小化到系统托盘",
                    ["TrayNotificationMessage"] = "双击托盘图标可恢复窗口",
                    
                    // 文件对话框
                    ["LogFileFilter"] = "文本文件 (*.txt)|*.txt",
                    ["AllFiles"] = "所有文件 (*.*)|*.*",
                    
                    // 对话框标题
                    ["SaveLogDialogTitle"] = "保存日志文件",
                    ["SelectDatabaseDialogTitle"] = "选择数据库文件",
                    ["DatabaseFileFilter"] = "数据库文件 (*.db)|*.db|所有文件 (*.*)|*.*",
                    
                    // 新增的语言资源
                    ["LoadPreviewError"] = "加载预览失败",
                    ["NoTemplatesAvailable"] = "没有可用的模板",
                    ["Information"] = "信息",
                    
                    // 模板设计器专用翻译
                    ["FieldsListTitle"] = "可用字段列表",
                    ["CustomTextTitle"] = "自定义文本",
                    ["EnterCustomTextHint"] = "请输入自定义文本。",
                    ["FieldPositionX"] = "X位置:",
                    ["FieldPositionY"] = "Y位置:",
                    ["FieldWidth"] = "字段宽度:",
                    ["FieldAlignment"] = "字段对齐:",
                    ["AlignLeft"] = "左对齐",
                    ["AlignCenter"] = "居中",
                    ["AlignRight"] = "右对齐",
                    ["ValueOnlyMode"] = "仅显示值",
                    ["AddFieldBtn"] = "添加字段",
                    ["UpdateFieldBtn"] = "更新字段",
                    ["RemoveFieldBtn"] = "移除字段",
                    ["ClearAllFields"] = "清空所有字段",
                    ["DesignCanvas"] = "设计画布",
                    ["TemplateProperties"] = "模板属性",
                    ["SaveCurrentTemplate"] = "保存当前模板",
                    ["LoadExistingTemplate"] = "加载现有模板",
                    ["CloseDesigner"] = "关闭设计器",
                    ["SelectTemplatePrompt"] = "请选择要加载的模板:",
                    ["ConfirmAction"] = "确认操作",
                    ["DesignInstructions"] = "操作说明：从左侧拖拽字段到设计区域，或输入自定义文本后点击添加。",
                    ["NoFieldSelected"] = "未选择字段"
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
                    ["DatabaseConfig"] = "Database Configuration",
                    ["MonitorControl"] = "Monitor Control",
                    ["StatusInfo"] = "Status Information",
                    ["RecentRecords"] = "Recent Records",
                    ["SerialNumber"] = "Serial Number",
                    ["TestDateTime"] = "Test Date Time",
                    ["Current"] = "Current(A)",
                    ["Voltage"] = "Voltage(V)",
                    ["VoltageVpm"] = "Vpm Voltage(V)",
                    ["Power"] = "Power(W)",
                    ["PrintCount"] = "Print Count",
                    ["Operation"] = "Operation",
                    ["DoubleClickToPrint"] = "Double Click to Print",
                    ["ViewDetails"] = "View Details",
                    ["PrintSelected"] = "Print Selected",
                    ["Refresh"] = "Refresh",
                    ["AutoPrint"] = "Auto Print",
                    ["EnablePrintCount"] = "Enable Print Count Statistics",
                    ["PrintPreview"] = "Print Preview",
                    
                    // System Configuration Tab
                    ["DatabasePath"] = "Database Path:",
                    ["Browse"] = "Browse...",
                    ["TestConnection"] = "Test Connection",
                    ["SelectDatabase"] = "Select Database",
                    ["SelectedPrinter"] = "Selected Printer:",
                    ["AutoStartMonitoring"] = "Auto Start Monitoring",
                    ["MinimizeToTray"] = "Minimize to System Tray",
                    ["PollInterval"] = "Poll Interval(ms):",
                    ["StartMonitoring"] = "Start Monitoring",
                    ["StopMonitoring"] = "Stop Monitoring",
                    ["SaveConfiguration"] = "Save Configuration",
                    ["Language"] = "Language:",
                    ["LanguageConfig"] = "Language Configuration",
                    ["PrinterConfig"] = "Printer Configuration",
                    ["ApplicationConfig"] = "Application Configuration",
                    ["TestPrint"] = "Test Print",
                    ["PrinterStatus"] = "Printer Status: Unknown",
                    ["PrinterStatusOK"] = "Printer Status: OK",
                    ["PrinterStatusError"] = "Printer Status: Error",
                    ["PrinterStatusOffline"] = "Printer Status: Offline",
                    ["GetPrinterListFailed"] = "Failed to get printer list",
                    
                    // Print Template Tab
                    ["PrintFormat"] = "Print Format:",
                    ["TemplateList"] = "Template List",
                    ["TemplateEditor"] = "Template Editor",
                    ["TemplatePreview"] = "Available Fields and Preview",
                    ["TemplateName"] = "Template Name:",
                    ["TemplateContent"] = "Template Content:",
                    ["AvailableFields"] = "Available Fields:",
                    ["FieldDescription"] = "Field Description:",
                    ["PreviewTemplate"] = "Preview Template",
                    ["SaveTemplate"] = "Save Template",
                    ["ResetTemplate"] = "Reset Template",
                    ["NewTemplate"] = "New Template",
                    ["DeleteTemplate"] = "Delete Template",
                    ["VisualDesigner"] = "Visual Designer",
                    ["PrePrintedLabelMode"] = "Pre-printed Label Mode",
                    ["FieldPositionSetting"] = "Field Position Setting",
                    ["SelectField"] = "Select Field:",
                    ["PosX"] = "X:",
                    ["PosY"] = "Y:",
                    ["Width"] = "Width:",
                    ["Alignment"] = "Alignment:",
                    ["ValueOnly"] = "Value Only",
                    ["AddField"] = "Add Field",
                    ["UpdateField"] = "Update Field",
                    ["DeleteField"] = "Delete Field",
                    ["ClearContent"] = "Clear Content",
                    ["ImportTemplate"] = "Import Template",
                    ["PreviewLabel"] = "Preview:",
                    ["ButtonsInfo"] = "Tip: Blue Save, Green Preview, Yellow Clear, Gray Import",
                    ["SavingTemplate"] = "Saving...",
                    ["NewTemplateText"] = "New Template",
                    
                    // Runtime Logs Tab
                    ["RuntimeLogs"] = "Runtime Logs",
                    ["ClearLogs"] = "Clear Logs",
                    ["SaveLogs"] = "Save Logs",
                    
                    // Print Preview Form
                    ["PrintPreviewTitle"] = "🖨️ Print Preview",
                    ["ConfirmPrint"] = "🖨️ Confirm Print",
                    ["ShowMainWindow"] = "📋 Show Main Window",
                    ["Close"] = "✖ Close",
                    ["LoadingContent"] = "Loading print content...",
                    ["NoPreviewData"] = "No preview data available",
                    ["AutoPrintEnabled"] = "Auto Print Enabled",
                    ["Loading"] = "Loading...",
                    
                    // Visual Designer Form
                    ["VisualDesignerTitle"] = "Template Visual Designer",
                    ["AvailableFieldsDesc"] = "Available Fields (Only reference data content, please enter project name yourself)",
                    ["CustomTextPlaceholder"] = "Enter custom text",
                    ["AddCustomText"] = "Add Custom Text",
                    ["ClearDesignArea"] = "Clear Design Area",
                    ["DesignArea"] = "Design Area",
                    ["PropertiesAndPreview"] = "Properties and Preview",
                    ["TemplateNameProp"] = "Template Name:",
                    ["OutputFormat"] = "Output Format:",
                    ["Preview"] = "Preview",
                    ["Clear"] = "Clear",
                    ["SaveTemplateBtn"] = "Save Template",
                    ["LoadTemplate"] = "Load Template",
                    ["Cancel"] = "Cancel",
                    ["SelectTemplate"] = "Select Template",
                    ["OK"] = "OK",
                    
                    // Buttons and Messages
                    ["Print"] = "Print",
                    ["Exit"] = "Exit",
                    ["ExitProgram"] = "Exit Program",
                    
                    // Status Information
                    ["Status"] = "Status:",
                    ["Ready"] = "Ready",
                    ["Monitoring"] = "Monitoring",
                    ["Stopped"] = "Stopped",
                    ["Connected"] = "Connected",
                    ["Disconnected"] = "Disconnected",
                    ["MonitoringStatus"] = "Monitoring Status:",
                    ["MonitoringStatusStopped"] = "Monitoring Status: Stopped",
                    ["MonitoringStatusRunning"] = "Monitoring Status: Running",
                    ["TotalRecords"] = "Total Records: 0",
                    ["TotalPrints"] = "Total Prints: 0",
                    ["LastRecord"] = "Last Record: N/A",
                    ["TotalPrintJobs"] = "Total Print Jobs:",
                    
                    // Log Information
                    ["LogLevel"] = "Log Level",
                    ["LogTime"] = "Time",
                    ["LogMessage"] = "Message",
                    
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
                    
                    // Status and Processing Information
                    ["ProcessedRecords"] = "Processed Records",
                    ["PrintJobs"] = "Print Jobs",
                    ["LastRecord"] = "Last Record",
                    ["CurrentPrinter"] = "Current Printer",
                    ["LanguageSwitched"] = "Language switched to",
                    ["PollIntervalChanged"] = "Poll interval changed to",
                    ["LogSaved"] = "Log saved to",
                    ["Enabled"] = "Enabled",
                    ["Disabled"] = "Disabled",
                    ["FunctionEnabled"] = "Function Enabled",
                    ["FunctionDisabled"] = "Function Disabled",
                    ["PrintCountStatistic"] = "Print count statistics",
                    ["PrintCountEnabledMessage"] = "Print count statistics enabled.",
                    ["PrintOperationWillUpdate"] = "New print operations will update the TR_Print field in the database.",
                    ["PrintCountDisabledMessage"] = "Print count statistics disabled.",
                    ["PrintOperationWillNotUpdate"] = "Print operations will not update the TR_Print field in the database, maintaining database compatibility.",
                    
                    // Template Related
                    ["TemplateNameRequired"] = "Please enter template name",
                    ["TemplateSaved"] = "Template saved",
                    ["TemplateSaveFailed"] = "Template save failed",
                    ["SelectTemplateToDelete"] = "Please select template to delete",
                    ["ConfirmDeleteTemplate"] = "Are you sure you want to delete template '{0}'?",
                    ["TemplateDeleted"] = "Template deleted",
                    ["TemplateDeleteFailed"] = "Template delete failed",
                    ["TemplateImported"] = "Template imported successfully",
                    ["ImportTemplateFailed"] = "Import template failed",
                    ["VisualDesignerError"] = "Failed to start visual designer",
                    ["SaveTemplateToolTip"] = "Save current template (Ctrl+S)",
                    
                    // Print Related
                    ["PrintRecordSuccess"] = "Successfully printed record",
                    ["PrintRecordFailed"] = "Failed to print record",
                    ["PrintTaskSent"] = "Print task sent",
                    ["TestPrintSent"] = "Test print sent to",
                    ["TestPrintFailed"] = "Test print failed",
                    ["TestPrintError"] = "Error occurred during test print",
                    ["PrintError"] = "Error occurred during printing",
                    ["PrintPreviewError"] = "Print preview failed",
                    ["PrintThisRecord"] = "Print this record",
                    ["InvalidRecordData"] = "Invalid record data",
                    
                    // Record Details
                    ["RecordDetails"] = "Record Details",
                    ["NA"] = "N/A",
                    ["ShortCircuitCurrent"] = "Short Circuit Current",
                    ["Isc"] = "Isc",
                    ["A"] = "A",
                    ["OpenCircuitVoltage"] = "Open Circuit Voltage",
                    ["Voc"] = "Voc",
                    ["V"] = "V",
                    ["MaximumPowerPointVoltage"] = "Maximum Power Point Voltage",
                    ["Vpm"] = "Vpm",
                    ["MaximumPowerPointCurrent"] = "Maximum Power Point Current",
                    ["Ipm"] = "Ipm",
                    ["MaximumPower"] = "Maximum Power",
                    ["Pm"] = "Pm",
                    ["W"] = "W",
                    ["Efficiency"] = "Efficiency",
                    ["FillFactor"] = "Fill Factor",
                    ["FF"] = "FF",
                    ["Grade"] = "Grade",
                    
                    // Printer related messages
                    ["NoPrinterFound"] = "No printer found in the system. Please install a printer first.",
                    ["NoPrinterInstalled"] = "No printer is installed in the system.\n\nPlease follow these steps to install a printer:\n1. Open \"Settings\" > \"Printers & scanners\"\n2. Click \"Add a printer or scanner\"\n3. Select your printer and follow the prompts to complete installation\n\nPlease restart this software after installation.",
                    ["NoPrinterTitle"] = "No Printer Installed",
                    ["ReprintWarningTitle"] = "Reprint Confirmation",
                    ["ReprintWarningMessage"] = "This record has been printed {0} times.\n\nDo you want to continue printing?",
                    ["Confirm"] = "Confirm",
                    ["PrintAgain"] = "Print Again",
                    
                    // Tray notifications
                    ["TrayNotificationTitle"] = "Program minimized to system tray",
                    ["TrayNotificationMessage"] = "Double-click tray icon to restore window",
                    
                    // File dialogs
                    ["LogFileFilter"] = "Text files (*.txt)|*.txt",
                    ["AllFiles"] = "All files (*.*)|*.*",
                    
                    // Dialog titles
                    ["SaveLogDialogTitle"] = "Save Log File",
                    ["SelectDatabaseDialogTitle"] = "Select Database File",
                    ["DatabaseFileFilter"] = "Database files (*.db)|*.db|All files (*.*)|*.*",
                    
                    // 新增的语言资源
                    ["LoadPreviewError"] = "Failed to load preview",
                    ["NoTemplatesAvailable"] = "No templates available",
                    ["Information"] = "Information",
                    
                    // 模板设计器专用翻译
                    ["FieldsListTitle"] = "Available Fields List",
                    ["CustomTextTitle"] = "Custom Text",
                    ["EnterCustomTextHint"] = "Please enter custom text.",
                    ["FieldPositionX"] = "X Position:",
                    ["FieldPositionY"] = "Y Position:",
                    ["FieldWidth"] = "Field Width:",
                    ["FieldAlignment"] = "Field Alignment:",
                    ["AlignLeft"] = "Align Left",
                    ["AlignCenter"] = "Align Center",
                    ["AlignRight"] = "Align Right",
                    ["ValueOnlyMode"] = "Value Only",
                    ["AddFieldBtn"] = "Add Field",
                    ["UpdateFieldBtn"] = "Update Field",
                    ["RemoveFieldBtn"] = "Remove Field",
                    ["ClearAllFields"] = "Clear All Fields",
                    ["DesignCanvas"] = "Design Canvas",
                    ["TemplateProperties"] = "Template Properties",
                    ["SaveCurrentTemplate"] = "Save Current Template",
                    ["LoadExistingTemplate"] = "Load Existing Template",
                    ["CloseDesigner"] = "Close Designer",
                    ["SelectTemplatePrompt"] = "Please select a template to load:",
                    ["ConfirmAction"] = "Confirm Action",
                    ["DesignInstructions"] = "Operation Instructions: Drag fields from the left to the design area, or add after entering custom text.",
                    ["NoFieldSelected"] = "No field selected"
                }
            };
        }
    }
} 