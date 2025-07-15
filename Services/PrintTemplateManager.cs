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

            // 处理对齐和换行
            content = ProcessAlignment(content);

            return content;
        }

        // 预印刷标签相关方法已删除

        private static string ProcessAlignment(string content)
        {
            // 新的对齐逻辑：
            // 1. 如果行只有值（如 {Voltage}V），则标记为右对齐（添加RIGHT_ALIGN标记）
            // 2. 如果有项目名称和值（如 Open Circuit Voltage(Voc): {Voltage}V），则保持原样用于左右对齐处理
            var lines = content.Split('\n');
            var processedLines = new List<string>();

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                if (string.IsNullOrEmpty(trimmedLine))
                    continue;

                // 检查是否只包含值（没有冒号且包含{}变量）
                bool hasColon = trimmedLine.Contains(':');
                bool hasVariable = trimmedLine.Contains('{') && trimmedLine.Contains('}');
                
                if (!hasColon && hasVariable)
                {
                    // 只有值的行，添加右对齐标记
                    processedLines.Add($"RIGHT_ALIGN:{trimmedLine}");
                }
                else
                {
                    // 有项目名称的行或其他行，保持原样
                    processedLines.Add(trimmedLine);
                }
            }

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
                    Content = @"Module Type: {SerialNumber}
Maximum Power(Pm): {Power}W
Open Circuit Voltage(Voc): {Voltage}V
Short Circuit Current(Isc): {Current}A
Maximum Power Voltage(Vm): {VoltageVpm}V
Maximum Power Current(Im): {Current}A
Weight: -- kg
Dimensions: ----×----×----
________________________________________________________________
Series Fuse Rating: 15A
Tolerance of Pm: 0~+5W
Measuring uncertainty of Pm: ±3%
Tolerance of Voc: ±3%
Tolerance of Isc: ±3%
Standard Test Conditions: 1000W/m², 25°C, AM1.5
Produced in accordance with: IEC 61215:2016 & IEC 61730:2016
Fire Rating/Module Fire Performance: Class C
MAX.System Voltage: 1000V
Module Protection: Class II",
                    Format = PrintFormat.Text,
                    IsDefault = true
                },
                new PrintTemplate
                {
                    Name = "简洁文本模板",
                    Content = @"序列号: {SerialNumber}
测试时间: {TestDateTime}
电流: {Current}A
电压: {Voltage}V
功率: {Power}W",
                    Format = PrintFormat.Text,
                    IsDefault = false
                },
                new PrintTemplate
                {
                    Name = "ZPL标签模板",
                    Content = @"^XA
^FO50,50^A0N,30,30^FD序列号: {SerialNumber}^FS
^FO50,100^A0N,25,25^FD测试时间: {TestDateTime}^FS
^FO50,150^A0N,25,25^FD电流: {Current}A^FS
^FO50,200^A0N,25,25^FD电压: {Voltage}V^FS
^FO50,250^A0N,25,25^FD功率: {Power}W^FS
^FO50,300^A0N,20,20^FD打印时间: {CurrentTime}^FS
^XZ",
                    Format = PrintFormat.ZPL,
                    IsDefault = false
                }
            };
        }
    }
} 