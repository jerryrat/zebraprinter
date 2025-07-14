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
            content = content.Replace("{CurrentImp}", record.FormatNumber(record.TR_Ipm));
            content = content.Replace("{Voltage}", record.FormatNumber(record.TR_Voc));
            content = content.Replace("{VoltageVpm}", record.FormatNumber(record.TR_Vpm));
            content = content.Replace("{Power}", record.FormatNumber(record.TR_Pm));
            content = content.Replace("{PrintCount}", (record.TR_Print ?? 0).ToString());

            // 添加时间戳
            content = content.Replace("{CurrentTime}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            content = content.Replace("{CurrentDate}", DateTime.Now.ToString("yyyy-MM-dd"));

            return content;
        }

        public static List<string> GetAvailableFields()
        {
            return new List<string>
            {
                "{SerialNumber}",
                "{TestDateTime}",
                "{Current}",
                "{CurrentImp}",
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
                ["{Current}"] = LanguageManager.GetString("Current", "短路电流Isc(A)"),
                ["{CurrentImp}"] = LanguageManager.GetString("CurrentImp", "最大功率电流Imp(A)"),
                ["{Voltage}"] = LanguageManager.GetString("Voltage", "开路电压Voc(V)"),
                ["{VoltageVpm}"] = LanguageManager.GetString("VoltageVpm", "最大功率电压Vmp(V)"),
                ["{Power}"] = LanguageManager.GetString("Power", "最大功率Pm(W)"),
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
                    Content = @"+---------------------------------------------------------------+
|                    SOLAR PANEL SPECIFICATION                 |
+--------------------------------------+----------------------------+
| Model Type                           | SKT 600 M12/120HB         |
+--------------------------------------+----------------------------+
| Rated Maximum Power      (Pmax)      | {Power} W                  |
+--------------------------------------+----------------------------+
| Power Sorting                        | 0~+5W                      |
+--------------------------------------+----------------------------+
| Voltage at Pmax          (Vmp)       | {VoltageVpm} V             |
+--------------------------------------+----------------------------+
| Current at Pmax          (Imp)       | {CurrentImp} A             |
+--------------------------------------+----------------------------+
| Open-Circuit Voltage     (Voc)       | {Voltage} V                |
+--------------------------------------+----------------------------+
| Short-Circuit Current    (Isc)       | {Current} A                |
+--------------------------------------+----------------------------+
| PV Module Classification             | CLASS II                   |
+--------------------------------------+----------------------------+
| Maximum System Voltage               | 1500 V                     |
+--------------------------------------+----------------------------+
| Maximum Series Fuse Rating           | 35 A                       |
+--------------------------------------+----------------------------+
| Operating Temperature                | -40~85°C                   |
+--------------------------------------+----------------------------+
| Dimensions(mm)                       | 2172x1303x40(mm)           |
+--------------------------------------+----------------------------+
| Pmax/Voc/Isc Tolerance               | ±3%                        |
+--------------------------------------+----------------------------+

Tested at STC:1000W/m²; AM1.5; Cell temperature 25°C",
                    Format = PrintFormat.Text,
                    IsDefault = true
                },


                new PrintTemplate
                {
                    Name = "紧凑表格格式",
                    Content = @"================================================================
                     SOLAR PANEL SPECIFICATION
================================================================
Model Type                           | SKT 600 M12/120HB
Rated Maximum Power      (Pmax)      | {Power} W
Power Sorting                        | 0~+5W
Voltage at Pmax          (Vmp)       | {VoltageVpm} V
Current at Pmax          (Imp)       | {CurrentImp} A
Open-Circuit Voltage     (Voc)       | {Voltage} V
Short-Circuit Current    (Isc)       | {Current} A
PV Module Classification             | CLASS II
Maximum System Voltage               | 1500 V
Maximum Series Fuse Rating           | 35 A
Operating Temperature                | -40~85°C
Dimensions(mm)                       | 2172x1303x40(mm)
Pmax/Voc/Isc Tolerance               | ±3%
================================================================
Tested at STC:1000W/m²; AM1.5; Cell temperature 25°C",
                    Format = PrintFormat.Text,
                    IsDefault = false
                },
                new PrintTemplate
                {
                    Name = "ZPL标签模板",
                    Content = @"^XA
^FO50,50^A0N,30,30^FD序列号: {SerialNumber}^FS
^FO50,100^A0N,25,25^FD测试时间: {TestDateTime}^FS
^FO50,150^A0N,25,25^FD短路电流(Isc): {Current}A^FS
^FO50,180^A0N,25,25^FD最大功率电流(Imp): {CurrentImp}A^FS
^FO50,210^A0N,25,25^FD开路电压(Voc): {Voltage}V^FS
^FO50,240^A0N,25,25^FD最大功率电压(Vmp): {VoltageVpm}V^FS
^FO50,270^A0N,25,25^FD最大功率(Pm): {Power}W^FS
^FO50,300^A0N,20,20^FD打印时间: {CurrentTime}^FS
^XZ",
                    Format = PrintFormat.ZPL,
                    IsDefault = false
                }
            };
        }
    }
} 