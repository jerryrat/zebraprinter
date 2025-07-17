using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;
using System.Linq; // Added for FirstOrDefault and Where

namespace ZebraPrinterMonitor.Services
{
    public class PrintTemplate
    {
        public string Name { get; set; } = "";
        public string Content { get; set; } = "";
        public PrintFormat Format { get; set; } = PrintFormat.Text;
        public bool IsDefault { get; set; } = false;
        public int FontSize { get; set; } = 10; // 默认字体大小为10
        public string FontName { get; set; } = "Arial"; // 默认字体名称
        
        // 页眉页脚设置
        public string HeaderText { get; set; } = "";
        public string FooterText { get; set; } = "";
        public string HeaderImagePath { get; set; } = "";
        public string FooterImagePath { get; set; } = "";
        public bool ShowHeader { get; set; } = false;
        public bool ShowFooter { get; set; } = false;
        
        // 预印刷标签功能已删除
    }

    // FieldPosition 和 TextAlignment 已删除

    public static class PrintTemplateManager
    {
        private static List<PrintTemplate> _templates = new();
        private static string _templatesFilePath = "print_templates.json";

        static PrintTemplateManager()
        {
            LoadTemplates();
        }

        public static List<PrintTemplate> GetTemplates() => new(_templates);

        public static PrintTemplate GetDefaultTemplate()
        {
            return _templates.FirstOrDefault(t => t.IsDefault) ?? GetBuiltInTemplates().First();
        }

        public static PrintTemplate? GetTemplate(string name)
        {
            return _templates.FirstOrDefault(t => t.Name == name);
        }

        public static void SaveTemplate(PrintTemplate template)
        {
            var existingTemplate = _templates.FirstOrDefault(t => t.Name == template.Name);
            if (existingTemplate != null)
            {
                existingTemplate.Content = template.Content;
                existingTemplate.Format = template.Format;
                existingTemplate.IsDefault = template.IsDefault;
                existingTemplate.FontSize = template.FontSize;
                existingTemplate.FontName = template.FontName;
                existingTemplate.HeaderText = template.HeaderText;
                existingTemplate.FooterText = template.FooterText;
                existingTemplate.HeaderImagePath = template.HeaderImagePath;
                existingTemplate.FooterImagePath = template.FooterImagePath;
                existingTemplate.ShowHeader = template.ShowHeader;
                existingTemplate.ShowFooter = template.ShowFooter;
            }
            else
            {
                _templates.Add(template);
            }

            // 确保只有一个默认模板
            if (template.IsDefault)
            {
                foreach (var t in _templates.Where(t => t != template))
                {
                    t.IsDefault = false;
                }
            }

            SaveTemplates();
            Logger.Info($"保存打印模板: {template.Name}");
        }

        public static void DeleteTemplate(string name)
        {
            var template = _templates.FirstOrDefault(t => t.Name == name);
            if (template != null)
            {
                _templates.Remove(template);
                SaveTemplates();
                Logger.Info($"删除打印模板: {name}");
            }
        }

        public static string ProcessTemplate(PrintTemplate template, TestRecord record)
        {
            var content = template.Content;
            
            // 先处理对齐标记（在变量替换之前）
            content = ProcessAlignment(content);
            
            // 替换模板变量 - 使用规范化的字段名称
            content = content.Replace("{SerialNumber}", record.TR_SerialNum ?? "N/A");
            content = content.Replace("{TestDateTime}", record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
            content = content.Replace("{ShortCircuitCurrent}", record.FormatNumber(record.TR_Isc));
            content = content.Replace("{OpenCircuitVoltage}", record.FormatNumber(record.TR_Voc));
            content = content.Replace("{MaxPowerVoltage}", record.FormatNumber(record.TR_Vpm));
            content = content.Replace("{MaxPower}", record.FormatNumber(record.TR_Pm));
            content = content.Replace("{MaxPowerCurrent}", record.FormatNumber(record.TR_Ipm));
            content = content.Replace("{PrintCount}", (record.TR_Print ?? 0).ToString());

            // 保持向后兼容性 - 兼容旧的字段名称
            content = content.Replace("{Current}", record.FormatNumber(record.TR_Isc));
            content = content.Replace("{Voltage}", record.FormatNumber(record.TR_Voc));
            content = content.Replace("{VoltageVpm}", record.FormatNumber(record.TR_Vpm));
            content = content.Replace("{Power}", record.FormatNumber(record.TR_Pm));

            // 添加时间戳
            content = content.Replace("{CurrentTime}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            content = content.Replace("{CurrentDate}", DateTime.Now.ToString("yyyy-MM-dd"));

            return content;
        }

        // 预印刷标签相关方法已删除

        private static string ProcessAlignment(string content)
        {
            // 🔧 修复换行符失效问题：更好地保持原始换行符格式
            // 1. 如果行只有值（如 {Voltage}V），则标记为右对齐（添加RIGHT_ALIGN标记）
            // 2. 如果有项目名称和值（如 Open Circuit Voltage(Voc): {Voltage}V），则保持原样用于左右对齐处理
            
            // 保持原始换行符格式，不要过度处理
            var lines = content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            var processedLines = new List<string>();

            foreach (var line in lines)
            {
                // 🔧 不要使用Trim()，保持原始的空白字符和格式
                var currentLine = line;
                
                // 保留完全空行（包含空白字符的行）
                if (string.IsNullOrWhiteSpace(currentLine))
                {
                    processedLines.Add(currentLine); // 保持原样，不改为空字符串
                    continue;
                }

                // 检查是否只包含值（没有冒号且包含{}变量）
                var trimmedForCheck = currentLine.Trim(); // 只用于检查，不影响原始内容
                bool hasColon = trimmedForCheck.Contains(':');
                bool hasVariable = trimmedForCheck.Contains('{') && trimmedForCheck.Contains('}');
                
                // 跳过装饰行（下划线、等号、破折号开头的行）
                bool isDecorationLine = trimmedForCheck.StartsWith("_") || 
                                      trimmedForCheck.StartsWith("=") || 
                                      trimmedForCheck.StartsWith("-") ||
                                      trimmedForCheck.All(c => c == '_' || c == '=' || c == '-' || char.IsWhiteSpace(c));
                
                if (!hasColon && hasVariable && !isDecorationLine)
                {
                    // 只有值的行，添加右对齐标记，但保持原始缩进
                    processedLines.Add($"RIGHT_ALIGN:{currentLine}");
                }
                else
                {
                    // 有项目名称的行或其他行，完全保持原样
                    processedLines.Add(currentLine);
                }
            }

            // 🔧 保持原始的换行符格式
            return string.Join("\r\n", processedLines);
        }

        public static List<string> GetAvailableFields()
        {
            return new List<string>
            {
                "{SerialNumber}",
                "{TestDateTime}",
                "{ShortCircuitCurrent}",
                "{OpenCircuitVoltage}",
                "{MaxPowerVoltage}",
                "{MaxPower}",
                "{MaxPowerCurrent}",
                "{PrintCount}",
                "{CurrentTime}",
                "{CurrentDate}"
            };
        }

        public static Dictionary<string, string> GetFieldDescriptions()
        {
            return new Dictionary<string, string>
            {
                ["{SerialNumber}"] = "序列号",
                ["{TestDateTime}"] = "测试时间",
                ["{ShortCircuitCurrent}"] = "短路电流(A)", // TR_Isc
                ["{OpenCircuitVoltage}"] = "开路电压(V)", // TR_Voc
                ["{MaxPowerVoltage}"] = "最大功率电压(V)", // TR_Vpm
                ["{MaxPower}"] = "最大功率(W)", // TR_Pm
                ["{MaxPowerCurrent}"] = "最大功率电流(A)", // TR_Ipm
                ["{PrintCount}"] = "打印次数",
                ["{CurrentTime}"] = "当前时间",
                ["{CurrentDate}"] = "当前日期"
            };
        }

        private static void LoadTemplates()
        {
            try
            {
                if (File.Exists(_templatesFilePath))
                {
                    var json = File.ReadAllText(_templatesFilePath);
                    var templates = JsonSerializer.Deserialize<List<PrintTemplate>>(json);
                    if (templates != null)
                    {
                        _templates = templates;
                        Logger.Info($"加载了 {_templates.Count} 个打印模板");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"加载打印模板失败: {ex.Message}", ex);
            }

            // 如果加载失败或没有文件，使用内置模板
            _templates = GetBuiltInTemplates();
            SaveTemplates();
        }

        private static void SaveTemplates()
        {
            try
            {
                var json = JsonSerializer.Serialize(_templates, new JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
                File.WriteAllText(_templatesFilePath, json);
            }
            catch (Exception ex)
            {
                Logger.Error($"保存打印模板失败: {ex.Message}", ex);
            }
        }

        private static List<PrintTemplate> GetBuiltInTemplates()
        {
            return new List<PrintTemplate>
            {
                new PrintTemplate
                {
                    Name = "默认文本模板",
                    Content = "Module Type: {SerialNumber}\r\n" +
                             "Maximum Power(Pm): {Power}W\r\n" +
                             "Open Circuit Voltage(Voc): {Voltage}V\r\n" +
                             "Short Circuit Current(Isc): {Current}A\r\n" +
                             "Maximum Power Voltage(Vm): {VoltageVpm}V\r\n" +
                             "Maximum Power Current(Im): {Current}A\r\n" +
                             "Weight: -- kg\r\n" +
                             "Dimensions: ----×----×----\r\n" +
                             "________________________________________________________________\r\n" +
                             "Series Fuse Rating: 15A\r\n" +
                             "Tolerance of Pm: 0~+5W\r\n" +
                             "Measuring uncertainty of Pm: ±3%\r\n" +
                             "Tolerance of Voc: ±3%\r\n" +
                             "Tolerance of Isc: ±3%\r\n" +
                             "Standard Test Conditions: 1000W/m², 25°C, AM1.5\r\n" +
                             "Produced in accordance with: IEC 61215:2016 & IEC 61730:2016\r\n" +
                             "Fire Rating/Module Fire Performance: Class C\r\n" +
                             "MAX.System Voltage: 1000V\r\n" +
                             "Module Protection: Class II",
                    Format = PrintFormat.Text,
                    IsDefault = true
                },
                new PrintTemplate
                {
                    Name = "简洁文本模板",
                    Content = "序列号: {SerialNumber}\r\n" +
                             "测试时间: {TestDateTime}\r\n" +
                             "电流: {Current}A\r\n" +
                             "电压: {Voltage}V\r\n" +
                             "功率: {Power}W",
                    Format = PrintFormat.Text,
                    IsDefault = false
                },
                new PrintTemplate
                {
                    Name = "ZPL标签模板",
                    Content = "^XA\r\n" +
                             "^FO50,50^A0N,30,30^FD序列号: {SerialNumber}^FS\r\n" +
                             "^FO50,100^A0N,25,25^FD测试时间: {TestDateTime}^FS\r\n" +
                             "^FO50,150^A0N,25,25^FD电流: {Current}A^FS\r\n" +
                             "^FO50,200^A0N,25,25^FD电压: {Voltage}V^FS\r\n" +
                             "^FO50,250^A0N,25,25^FD功率: {Power}W^FS\r\n" +
                             "^FO50,300^A0N,20,20^FD打印时间: {CurrentTime}^FS\r\n" +
                             "^XZ",
                    Format = PrintFormat.ZPL,
                    IsDefault = false
                }
            };
        }
    }
} 