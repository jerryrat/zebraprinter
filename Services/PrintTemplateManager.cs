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
    }

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
            
            // 替换模板变量
            content = content.Replace("{SerialNumber}", record.TR_SerialNum ?? "N/A");
            content = content.Replace("{TestDateTime}", record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
            content = content.Replace("{Current}", record.FormatNumber(record.TR_Isc));
            content = content.Replace("{Voltage}", record.FormatNumber(record.TR_Voc));
            content = content.Replace("{VoltageVpm}", record.FormatNumber(record.TR_Vpm));
            content = content.Replace("{Power}", record.FormatNumber(record.TR_Pm));
            content = content.Replace("{PrintCount}", (record.TR_Print ?? 0).ToString());

            // 添加时间戳
            content = content.Replace("{CurrentTime}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            content = content.Replace("{CurrentDate}", DateTime.Now.ToString("yyyy-MM-dd"));

            // 处理对齐和换行
            content = ProcessAlignment(content);

            return content;
        }

        private static string ProcessAlignment(string content)
        {
            var lines = content.Split('\n');
            var processedLines = new List<string>();

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                
                // 处理包含冒号的行，实现左对齐和右对齐
                if (trimmedLine.Contains(':') && !trimmedLine.StartsWith("_") && !trimmedLine.StartsWith("-"))
                {
                    var parts = trimmedLine.Split(':', 2);
                    if (parts.Length == 2)
                    {
                        var label = parts[0].Trim();
                        var value = parts[1].Trim();
                        
                        // 如果标签或值太长，进行处理
                        var processedLine = FormatAlignedLine(label, value);
                        processedLines.Add(processedLine);
                    }
                    else
                    {
                        processedLines.Add(trimmedLine);
                    }
                }
                else
                {
                    processedLines.Add(trimmedLine);
                }
            }

            return string.Join("\r\n", processedLines);
        }

        private static string FormatAlignedLine(string label, string value, int totalWidth = 60)
        {
            // 清理标签和值
            label = label.Trim();
            value = value.Trim();

            // 处理超长标签
            if (label.Length > totalWidth - 10)
            {
                // 如果标签太长，分成两行
                var wrappedLabel = WrapText(label, totalWidth - 4);
                var lines = wrappedLabel.Split('\n');
                var result = new List<string>();
                
                for (int i = 0; i < lines.Length; i++)
                {
                    if (i == lines.Length - 1)
                    {
                        // 最后一行，添加值
                        result.Add(FormatSingleLine(lines[i], value, totalWidth));
                    }
                    else
                    {
                        // 其他行只有标签
                        result.Add(lines[i]);
                    }
                }
                return string.Join("\r\n", result);
            }
            else
            {
                return FormatSingleLine(label, value, totalWidth);
            }
        }

        private static string FormatSingleLine(string label, string value, int totalWidth)
        {
            // 如果值太长，也需要处理
            if (value.Length > totalWidth - label.Length - 3)
            {
                // 值太长，分成两行
                var availableWidth = totalWidth - 4;
                var wrappedValue = WrapText(value, availableWidth);
                var valueLines = wrappedValue.Split('\n');
                
                var result = new List<string>();
                // 第一行：标签 + 第一部分值
                result.Add(label.PadRight(totalWidth - valueLines[0].Length) + valueLines[0]);
                
                // 后续行：只有值，右对齐
                for (int i = 1; i < valueLines.Length; i++)
                {
                    result.Add(valueLines[i].PadLeft(totalWidth));
                }
                
                return string.Join("\r\n", result);
            }
            else
            {
                // 正常情况：标签左对齐，值右对齐
                return label.PadRight(totalWidth - value.Length) + value;
            }
        }

        private static string WrapText(string text, int width)
        {
            if (string.IsNullOrEmpty(text) || width <= 0)
                return text;

            var words = text.Split(' ');
            var lines = new List<string>();
            var currentLine = "";

            foreach (var word in words)
            {
                if (string.IsNullOrEmpty(currentLine))
                {
                    currentLine = word;
                }
                else if (currentLine.Length + word.Length + 1 <= width)
                {
                    currentLine += " " + word;
                }
                else
                {
                    lines.Add(currentLine);
                    currentLine = word;
                }
            }

            if (!string.IsNullOrEmpty(currentLine))
            {
                lines.Add(currentLine);
            }

            return string.Join("\n", lines);
        }

        public static List<string> GetAvailableFields()
        {
            return new List<string>
            {
                "{SerialNumber}",
                "{TestDateTime}",
                "{Current}",
                "{Voltage}",
                "{VoltageVpm}",
                "{Power}",
                "{PrintCount}",
                "{CurrentTime}",
                "{CurrentDate}"
            };
        }

        public static Dictionary<string, string> GetFieldDescriptions()
        {
            return new Dictionary<string, string>
            {
                ["{SerialNumber}"] = LanguageManager.GetString("SerialNumber", "序列号"),
                ["{TestDateTime}"] = LanguageManager.GetString("TestDateTime", "测试时间"),
                ["{Current}"] = LanguageManager.GetString("Current", "电流(A)"),
                ["{Voltage}"] = LanguageManager.GetString("Voltage", "电压(V)"),
                ["{VoltageVpm}"] = LanguageManager.GetString("VoltageVpm", "Vpm电压(V)"),
                ["{Power}"] = LanguageManager.GetString("Power", "功率(W)"),
                ["{PrintCount}"] = LanguageManager.GetString("PrintCount", "打印次数"),
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