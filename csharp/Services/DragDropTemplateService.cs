using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Drawing.Printing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Services
{
    /// <summary>
    /// 拖拽式打印模板服务 - 使用HTML/JavaScript实现可视化编辑
    /// </summary>
    public class DragDropTemplateService : IDisposable
    {
        private readonly string _templatesPath;
        private readonly string _previewPath;
        private Dictionary<string, PrintElementTemplate> _currentTemplate;

        public DragDropTemplateService()
        {
            _templatesPath = Path.Combine(Environment.CurrentDirectory, "PrintTemplates");
            _previewPath = Path.Combine(Environment.CurrentDirectory, "TemplatePreview");
            
            if (!Directory.Exists(_templatesPath))
            {
                Directory.CreateDirectory(_templatesPath);
            }
            
            if (!Directory.Exists(_previewPath))
            {
                Directory.CreateDirectory(_previewPath);
            }

            _currentTemplate = new Dictionary<string, PrintElementTemplate>();
            InitializeDefaultTemplate();
            Logger.Info("拖拽式模板编辑器服务初始化完成");
        }

        /// <summary>
        /// 打印元素模板类
        /// </summary>
        public class PrintElementTemplate
        {
            public string Id { get; set; } = "";
            public string Type { get; set; } = ""; // text, image, barcode, line, rectangle
            public double X { get; set; } = 0;
            public double Y { get; set; } = 0;
            public double Width { get; set; } = 100;
            public double Height { get; set; } = 30;
            public string Content { get; set; } = "";
            public Dictionary<string, object> Properties { get; set; } = new();
        }

        /// <summary>
        /// 模板数据类
        /// </summary>
        public class TemplateData
        {
            public string Name { get; set; } = "默认模板";
            public string Description { get; set; } = "";
            public double PageWidth { get; set; } = 210; // A4宽度 mm
            public double PageHeight { get; set; } = 297; // A4高度 mm
            public List<PrintElementTemplate> Elements { get; set; } = new();
            public DateTime CreatedDate { get; set; } = DateTime.Now;
            public DateTime ModifiedDate { get; set; } = DateTime.Now;
        }

        /// <summary>
        /// 初始化默认模板
        /// </summary>
        private void InitializeDefaultTemplate()
        {
            try
            {
                string defaultTemplatePath = Path.Combine(_templatesPath, "SolarCellTest.json");
                
                if (!File.Exists(defaultTemplatePath))
                {
                    CreateDefaultTemplate();
                    SaveTemplate("SolarCellTest", _currentTemplate);
                    Logger.Info("创建默认太阳能电池测试打印模板");
                }
                else
                {
                    LoadTemplate("SolarCellTest");
                    Logger.Info("加载现有打印模板");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"初始化默认模板失败: {ex.Message}", ex);
                CreateDefaultTemplate(); // 创建基本模板作为后备
            }
        }

        /// <summary>
        /// 创建默认的太阳能电池测试打印模板
        /// </summary>
        private void CreateDefaultTemplate()
        {
            _currentTemplate.Clear();

            // 主标题 - SKT 600 M12/120HB
            _currentTemplate["mainTitle"] = new PrintElementTemplate
            {
                Id = "mainTitle",
                Type = "text",
                X = 10,
                Y = 5,
                Width = 350,
                Height = 20,
                Content = "SKT 600 M12/120HB",
                Properties = new Dictionary<string, object>
                {
                    ["fontSize"] = 18,
                    ["fontWeight"] = "bold",
                    ["textAlign"] = "center",
                    ["fontFamily"] = "Arial",
                    ["backgroundColor"] = "#f5f5f5"
                }
            };

            // 表格外边框
            _currentTemplate["tableBorder"] = new PrintElementTemplate
            {
                Id = "tableBorder",
                Type = "rectangle",
                X = 10,
                Y = 25,
                Width = 350,
                Height = 220,
                Content = "",
                Properties = new Dictionary<string, object>
                {
                    ["strokeWidth"] = 2,
                    ["strokeColor"] = "#000000",
                    ["fillColor"] = "transparent"
                }
            };

            // 表格行数据 - 13行参数
            var tableData = new[]
            {
                new { Label = "Model Type", Value = "SKT 600 M12/120HB", Unit = "", Y = 30 },
                new { Label = "Rated Maximum Power", Value = "600 W", Unit = "(Pmax)", Y = 47 },
                new { Label = "Power Sorting", Value = "0~+5W", Unit = "", Y = 64 },
                new { Label = "Voltage at Pmax", Value = "34.60 V", Unit = "(Vmp)", Y = 81 },
                new { Label = "Current at Pmax", Value = "17.34 A", Unit = "(Imp)", Y = 98 },
                new { Label = "Open-Circuit Voltage", Value = "41.70 V", Unit = "(Voc)", Y = 115 },
                new { Label = "Short-Circuit Current", Value = "18.42 A", Unit = "(Isc)", Y = 132 },
                new { Label = "PV Module Classification", Value = "CLASS II", Unit = "", Y = 149 },
                new { Label = "Maximum System Voltage", Value = "1500 V", Unit = "", Y = 166 },
                new { Label = "Maximum Series Fuse Rating", Value = "35 A", Unit = "", Y = 183 },
                new { Label = "Operating Temperature", Value = "-40~85°C", Unit = "", Y = 200 },
                new { Label = "Dimensions(mm)", Value = "2172x1303x40(mm)", Unit = "", Y = 217 },
                new { Label = "Pmax/Voc/Isc Tolerance", Value = "±3%", Unit = "", Y = 234 }
            };

            // 创建表格行
            for (int i = 0; i < tableData.Length; i++)
            {
                var row = tableData[i];
                
                // 行背景色（隔行变色）
                if (i % 2 == 0)
                {
                    _currentTemplate[$"rowBg_{i}"] = new PrintElementTemplate
                    {
                        Id = $"rowBg_{i}",
                        Type = "rectangle",
                        X = 10,
                        Y = row.Y,
                        Width = 350,
                        Height = 17,
                        Content = "",
                        Properties = new Dictionary<string, object>
                        {
                            ["fillColor"] = "#f9f9f9",
                            ["strokeWidth"] = 0
                        }
                    };
                }

                // 参数名称
                _currentTemplate[$"label_{i}"] = new PrintElementTemplate
                {
                    Id = $"label_{i}",
                    Type = "text",
                    X = 15,
                    Y = row.Y + 2,
                    Width = 180,
                    Height = 13,
                    Content = row.Label,
                    Properties = new Dictionary<string, object>
                    {
                        ["fontSize"] = 11,
                        ["fontFamily"] = "Arial",
                        ["textAlign"] = "left",
                        ["fontWeight"] = "normal"
                    }
                };

                // 单位（如果有）
                if (!string.IsNullOrEmpty(row.Unit))
                {
                    _currentTemplate[$"unit_{i}"] = new PrintElementTemplate
                    {
                        Id = $"unit_{i}",
                        Type = "text",
                        X = 200,
                        Y = row.Y + 2,
                        Width = 50,
                        Height = 13,
                        Content = row.Unit,
                        Properties = new Dictionary<string, object>
                        {
                            ["fontSize"] = 11,
                            ["fontFamily"] = "Arial",
                            ["textAlign"] = "center",
                            ["fontWeight"] = "normal"
                        }
                    };
                }

                // 参数值
                _currentTemplate[$"value_{i}"] = new PrintElementTemplate
                {
                    Id = $"value_{i}",
                    Type = "text",
                    X = 255,
                    Y = row.Y + 2,
                    Width = 100,
                    Height = 13,
                    Content = row.Value,
                    Properties = new Dictionary<string, object>
                    {
                        ["fontSize"] = 11,
                        ["fontFamily"] = "Arial",
                        ["textAlign"] = "right",
                        ["fontWeight"] = "normal"
                    }
                };

                // 行分割线
                _currentTemplate[$"rowLine_{i}"] = new PrintElementTemplate
                {
                    Id = $"rowLine_{i}",
                    Type = "line",
                    X = 10,
                    Y = row.Y + 17,
                    Width = 350,
                    Height = 1,
                    Content = "",
                    Properties = new Dictionary<string, object>
                    {
                        ["strokeWidth"] = 1,
                        ["strokeColor"] = "#cccccc"
                    }
                };
            }

            // 测试条件说明
            _currentTemplate["testConditions"] = new PrintElementTemplate
            {
                Id = "testConditions",
                Type = "text",
                X = 10,
                Y = 255,
                Width = 350,
                Height = 15,
                Content = "Tested at STC:1000W/m² ; AM1.5 ; Cell temperature 25°C",
                Properties = new Dictionary<string, object>
                {
                    ["fontSize"] = 10,
                    ["fontFamily"] = "Arial",
                    ["textAlign"] = "left",
                    ["fontWeight"] = "normal",
                    ["backgroundColor"] = "#f0f0f0"
                }
            };

            // 右侧二维码区域
            _currentTemplate["qrCodeArea"] = new PrintElementTemplate
            {
                Id = "qrCodeArea",
                Type = "rectangle",
                X = 370,
                Y = 50,
                Width = 60,
                Height = 120,
                Content = "",
                Properties = new Dictionary<string, object>
                {
                    ["fillColor"] = "#000000",
                    ["strokeWidth"] = 1,
                    ["strokeColor"] = "#000000"
                }
            };

            // 二维码标识文本
            _currentTemplate["qrCodeText"] = new PrintElementTemplate
            {
                Id = "qrCodeText",
                Type = "text",
                X = 370,
                Y = 175,
                Width = 60,
                Height = 10,
                Content = "QR Code",
                Properties = new Dictionary<string, object>
                {
                    ["fontSize"] = 8,
                    ["fontFamily"] = "Arial",
                    ["textAlign"] = "center",
                    ["color"] = "#666666"
                }
            };
        }

        /// <summary>
        /// 生成打印内容HTML
        /// </summary>
        public string GeneratePrintHtml(TestRecord testRecord)
        {
            try
            {
                var html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        @page { 
            size: A4; 
            margin: 0; 
        }
        body { 
            margin: 0; 
            padding: 0; 
            font-family: 'Microsoft YaHei', sans-serif; 
        }
        .print-container {
            position: relative;
            width: 210mm;
            height: 297mm;
            background: white;
        }
        .element {
            position: absolute;
        }
        .barcode {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            border: 1px solid #000;
            font-family: 'Courier New', monospace;
        }
        .line {
            border-top: 1px solid #000;
        }
    </style>
</head>
<body>
    <div class='print-container'>";

                foreach (var element in _currentTemplate.Values)
                {
                    html += GenerateElementHtml(element, testRecord);
                }

                html += @"
    </div>
</body>
</html>";

                return html;
            }
            catch (Exception ex)
            {
                Logger.Error($"生成打印HTML失败: {ex.Message}", ex);
                return "";
            }
        }

        /// <summary>
        /// 生成单个元素的HTML
        /// </summary>
        private string GenerateElementHtml(PrintElementTemplate element, TestRecord testRecord)
        {
            string content = ReplaceVariables(element.Content, testRecord);
            string style = $"left: {element.X}mm; top: {element.Y}mm; width: {element.Width}mm; height: {element.Height}mm;";

            // 添加元素特定样式
            if (element.Properties.ContainsKey("fontSize"))
                style += $" font-size: {element.Properties["fontSize"]}pt;";
            if (element.Properties.ContainsKey("fontWeight"))
                style += $" font-weight: {element.Properties["fontWeight"]};";
            if (element.Properties.ContainsKey("textAlign"))
                style += $" text-align: {element.Properties["textAlign"]};";
            if (element.Properties.ContainsKey("fontFamily"))
                style += $" font-family: '{element.Properties["fontFamily"]}';";

            switch (element.Type.ToLower())
            {
                case "text":
                    return $"<div class='element' style='{style}'>{content}</div>";
                
                case "barcode":
                    return $"<div class='element barcode' style='{style}'>|||{content}|||</div>";
                
                case "line":
                    return $"<div class='element line' style='{style}'></div>";
                
                case "rectangle":
                    return $"<div class='element' style='{style} border: 1px solid #000;'></div>";
                
                default:
                    return $"<div class='element' style='{style}'>{content}</div>";
            }
        }

        /// <summary>
        /// 替换模板变量
        /// </summary>
        private string ReplaceVariables(string template, TestRecord testRecord)
        {
            if (string.IsNullOrEmpty(template)) return "";

            return template
                .Replace("{TestTime}", testRecord.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A")
                .Replace("{SerialNumber}", testRecord.TR_SerialNum ?? "N/A")
                .Replace("{Current}", testRecord.TR_Ipm?.ToString("F3") ?? "N/A")
                .Replace("{Voltage}", testRecord.TR_Vpm?.ToString("F3") ?? "N/A")
                .Replace("{Power}", testRecord.TR_Pm?.ToString("F3") ?? "N/A")
                .Replace("{TestResult}", testRecord.TR_Grade ?? "N/A")
                .Replace("{Remarks}", testRecord.TR_Operater ?? "N/A");
        }

        /// <summary>
        /// 打印模板
        /// </summary>
        public bool PrintTemplate(TestRecord testRecord)
        {
            try
            {
                string html = GeneratePrintHtml(testRecord);
                string htmlFile = Path.Combine(_previewPath, $"print_{DateTime.Now:yyyyMMdd_HHmmss}.html");
                
                File.WriteAllText(htmlFile, html);
                
                // 使用默认浏览器打印
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = htmlFile,
                    UseShellExecute = true
                });

                Logger.Info($"打印模板已生成: {htmlFile}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"打印模板失败: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// 预览模板
        /// </summary>
        public string PreviewTemplate(TestRecord testRecord)
        {
            try
            {
                string html = GeneratePrintHtml(testRecord);
                string previewFile = Path.Combine(_previewPath, "preview.html");
                
                File.WriteAllText(previewFile, html);
                Logger.Info("模板预览已生成");
                
                return previewFile;
            }
            catch (Exception ex)
            {
                Logger.Error($"预览模板失败: {ex.Message}", ex);
                return "";
            }
        }

        /// <summary>
        /// 保存模板
        /// </summary>
        public bool SaveTemplate(string templateName, Dictionary<string, PrintElementTemplate> template)
        {
            try
            {
                var templateData = new TemplateData
                {
                    Name = templateName,
                    Description = "太阳能电池测试打印模板",
                    Elements = new List<PrintElementTemplate>(template.Values),
                    ModifiedDate = DateTime.Now
                };

                string json = JsonSerializer.Serialize(templateData, new JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
                
                string filePath = Path.Combine(_templatesPath, $"{templateName}.json");
                File.WriteAllText(filePath, json);
                
                Logger.Info($"模板保存成功: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"保存模板失败: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// 加载模板
        /// </summary>
        public bool LoadTemplate(string templateName)
        {
            try
            {
                string filePath = Path.Combine(_templatesPath, $"{templateName}.json");
                if (!File.Exists(filePath))
                {
                    Logger.Warning($"模板文件不存在: {filePath}");
                    return false;
                }

                string json = File.ReadAllText(filePath);
                var templateData = JsonSerializer.Deserialize<TemplateData>(json);
                
                if (templateData?.Elements != null)
                {
                    _currentTemplate.Clear();
                    foreach (var element in templateData.Elements)
                    {
                        _currentTemplate[element.Id] = element;
                    }
                    
                    Logger.Info($"模板加载成功: {templateName}");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"加载模板失败: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// 获取可用模板列表
        /// </summary>
        public string[] GetAvailableTemplates()
        {
            try
            {
                var files = Directory.GetFiles(_templatesPath, "*.json");
                var templateNames = new string[files.Length];
                
                for (int i = 0; i < files.Length; i++)
                {
                    templateNames[i] = Path.GetFileNameWithoutExtension(files[i]);
                }
                
                return templateNames;
            }
            catch (Exception ex)
            {
                Logger.Error($"获取模板列表失败: {ex.Message}", ex);
                return new string[0];
            }
        }

        /// <summary>
        /// 获取当前模板数据（用于编辑器）
        /// </summary>
        public Dictionary<string, PrintElementTemplate> GetCurrentTemplate()
        {
            return new Dictionary<string, PrintElementTemplate>(_currentTemplate);
        }

        /// <summary>
        /// 更新模板元素
        /// </summary>
        public void UpdateTemplate(Dictionary<string, PrintElementTemplate> template)
        {
            _currentTemplate = new Dictionary<string, PrintElementTemplate>(template);
        }

        /// <summary>
        /// 打开模板编辑器
        /// </summary>
        public void OpenTemplateEditor()
        {
            try
            {
                // 暂时禁用模板编辑器，专注于核心模板功能
                Logger.Info("太阳能电池板规格表模板已1:1还原图片样式");
                
                // TODO: 重新启用模板编辑器
                // var editorForm = new TemplateEditorForm(this);
                // editorForm.ShowDialog();
            }
            catch (Exception ex)
            {
                Logger.Error($"打开模板编辑器失败: {ex.Message}", ex);
                MessageBox.Show($"打开模板编辑器失败: {ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Dispose()
        {
            _currentTemplate?.Clear();
            Logger.Info("拖拽式模板服务已释放");
        }
    }
} 