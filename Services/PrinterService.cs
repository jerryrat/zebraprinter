using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Services
{
    public class PrinterService
    {
        private string _currentPrinterName;
        private PrintDocument _printDocument;
        private int _currentFontSize = 10; // 默认字体大小
        private string _currentFontName = "Arial"; // 默认字体名称
        private PrintTemplate? _currentTemplate; // 当前打印模板

        public PrinterService()
        {
            _currentPrinterName = ConfigurationManager.Config.Printer.PrinterName;
            _printDocument = new PrintDocument();
            _printDocument.PrintPage += OnPrintPage;
        }

        private string? _currentPrintContent;

        public List<string> GetAvailablePrinters()
        {
            var printers = new List<string>();

            try
            {
                // 方法1: 使用 PrinterSettings
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    printers.Add(printer);
                }

                Logger.Info($"找到 {printers.Count} 个打印机");
                return printers;
            }
            catch (Exception ex)
            {
                Logger.Error($"获取打印机列表失败: {ex.Message}", ex);

                // 方法2: 使用 WMI 作为备用
                try
                {
                    printers.Clear();
                    using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
                    
                    foreach (ManagementObject printer in searcher.Get())
                    {
                        var name = printer["Name"]?.ToString();
                        if (!string.IsNullOrEmpty(name))
                        {
                            printers.Add(name);
                        }
                    }

                    Logger.Info($"通过WMI找到 {printers.Count} 个打印机");
                }
                catch (Exception wmiEx)
                {
                    Logger.Error($"WMI获取打印机失败: {wmiEx.Message}", wmiEx);
                }

                return printers;
            }
        }

        public bool IsPrinterAvailable(string printerName)
        {
            var availablePrinters = GetAvailablePrinters();
            return availablePrinters.Contains(printerName);
        }

        public bool HasAnyPrinter()
        {
            var printers = GetAvailablePrinters();
            return printers.Count > 0;
        }

        public string GetNoPrinterMessage()
        {
            return LanguageManager.GetString("NoPrinterInstalled", 
                "系统中没有安装打印机。\n\n请按照以下步骤安装打印机：\n" +
                "1. 打开\"设置\" > \"打印机和扫描仪\"\n" +
                "2. 点击\"添加打印机或扫描仪\"\n" +
                "3. 选择您的打印机并按照提示完成安装\n\n" +
                "安装完成后，请重新启动本软件。");
        }

        public void UpdatePrinterName(string printerName)
        {
            _currentPrinterName = printerName;
            Logger.Info($"打印机设置已更新: {printerName}");
        }

        public PrintResult PrintRecord(TestRecord record, string format = "Text", string? templateName = null)
        {
            try
            {
                // 检查是否有任何打印机
                if (!HasAnyPrinter())
                {
                    return new PrintResult
                    {
                        Success = false,
                        ErrorMessage = LanguageManager.GetString("NoPrinterFound", "系统中没有找到任何打印机，请先安装打印机。")
                    };
                }

                if (!IsPrinterAvailable(_currentPrinterName))
                {
                    return new PrintResult
                    {
                        Success = false,
                        ErrorMessage = $"打印机 '{_currentPrinterName}' 不可用"
                    };
                }

                Logger.Info($"开始打印记录: {record.TR_SerialNum}, 格式: {format}, 模板: {templateName ?? "默认"}");

                // 如果指定了模板名称，使用模板打印
                if (!string.IsNullOrEmpty(templateName))
                {
                    return PrintWithTemplate(record, templateName);
                }

                // 否则使用默认格式
                switch (format.ToLower())
                {
                    case "zpl":
                        return PrintZPL(record);
                    case "code128":
                        return PrintCode128(record);
                    case "qr":
                        return PrintQRCode(record);
                    default:
                        return PrintText(record);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"打印记录失败: {ex.Message}", ex);
                return new PrintResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private PrintResult PrintWithTemplate(TestRecord record, string templateName)
        {
            var template = PrintTemplateManager.GetTemplate(templateName);
            if (template == null)
            {
                // 如果找不到模板，使用默认模板
                template = PrintTemplateManager.GetDefaultTemplate();
            }

            var content = PrintTemplateManager.ProcessTemplate(template, record);
            
            // 设置当前字体大小和名称
            _currentFontSize = template.FontSize;
            _currentFontName = template.FontName;
            _currentTemplate = template; // 保存当前模板以供页眉页脚使用

            return template.Format switch
            {
                PrintFormat.Text => PrintCustomText(content),
                PrintFormat.ZPL => SendRawDataToPrinter(content, "ZPL模板"),
                PrintFormat.Code128 => SendRawDataToPrinter(content, "Code128模板"),
                PrintFormat.QRCode => SendRawDataToPrinter(content, "QRCode模板"),
                _ => PrintCustomText(content)
            };
        }

        private PrintResult PrintCustomText(string content)
        {
            try
            {
                _currentPrintContent = content;
                _printDocument.PrinterSettings.PrinterName = _currentPrinterName;
                _printDocument.Print();

                Logger.Info($"文本打印任务已发送到: {_currentPrinterName}");
                return new PrintResult
                {
                    Success = true,
                    Method = "自定义文本模板",
                    PrinterUsed = _currentPrinterName
                };
            }
            catch (Exception ex)
            {
                var error = $"文本打印失败: {ex.Message}";
                Logger.Error(error, ex);
                return new PrintResult
                {
                    Success = false,
                    ErrorMessage = error,
                    Method = "自定义文本模板",
                    PrinterUsed = _currentPrinterName
                };
            }
        }

        private PrintResult PrintText(TestRecord record)
        {
            try
            {
                _currentPrintContent = GenerateTextReceipt(record);
                
                _printDocument.PrinterSettings.PrinterName = _currentPrinterName;
                _printDocument.DocumentName = $"测试标签_{record.TR_SerialNum}_{DateTime.Now:yyyyMMddHHmmss}";

                _printDocument.Print();

                Logger.Info($"文本打印任务已发送到: {_currentPrinterName}");

                return new PrintResult
                {
                    Success = true,
                    Method = "Text",
                    JobId = $"TXT_{DateTime.Now:yyyyMMddHHmmss}",
                    PrinterUsed = _currentPrinterName
                };
            }
            catch (Exception ex)
            {
                Logger.Error($"文本打印失败: {ex.Message}", ex);
                throw;
            }
        }

        private PrintResult PrintZPL(TestRecord record)
        {
            try
            {
                var zplCode = GenerateZPLCode(record);
                return SendRawDataToPrinter(zplCode, "ZPL");
            }
            catch (Exception ex)
            {
                Logger.Error($"ZPL打印失败: {ex.Message}", ex);
                throw;
            }
        }

        private PrintResult PrintCode128(TestRecord record)
        {
            try
            {
                var zplCode = GenerateCode128ZPL(record);
                return SendRawDataToPrinter(zplCode, "Code128");
            }
            catch (Exception ex)
            {
                Logger.Error($"Code128打印失败: {ex.Message}", ex);
                throw;
            }
        }

        private PrintResult PrintQRCode(TestRecord record)
        {
            try
            {
                var zplCode = GenerateQRCodeZPL(record);
                return SendRawDataToPrinter(zplCode, "QRCode");
            }
            catch (Exception ex)
            {
                Logger.Error($"QR码打印失败: {ex.Message}", ex);
                throw;
            }
        }

        private void OnPrintPage(object sender, PrintPageEventArgs e)
        {
            if (e.Graphics == null || string.IsNullOrEmpty(_currentPrintContent))
            {
                e.HasMorePages = false;
                return;
            }

            // 使用自定义字体
            var font = new Font(_currentFontName, _currentFontSize);
            var brush = new SolidBrush(Color.Black);
            var leftMargin = 50;
            var topMargin = 50;
            var bottomMargin = 50;
            var lineHeight = _currentFontSize + 8; // 动态行高，基于字体大小
            var pageWidth = e.PageBounds.Width - leftMargin * 2;
            var pageHeight = e.PageBounds.Height - topMargin - bottomMargin;
            
            var currentY = (float)topMargin;

            // 绘制页眉
            if (_currentTemplate?.ShowHeader == true)
            {
                currentY = DrawHeader(e.Graphics, font, brush, leftMargin, currentY, pageWidth);
                currentY += lineHeight; // 页眉与内容之间的间距
            }

            // 计算内容区域的可用高度
            var footerHeight = 0;
            if (_currentTemplate?.ShowFooter == true)
            {
                // 预估页脚高度
                footerHeight = CalculateFooterHeight(e.Graphics, font, pageWidth) + lineHeight;
            }

            var contentAreaHeight = pageHeight - (currentY - topMargin) - footerHeight;
            var startY = currentY;

            var lines = _currentPrintContent.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var y = startY + (i * lineHeight);
                
                // 检查是否超出内容区域
                if (y + lineHeight > startY + contentAreaHeight)
                {
                    break; // 超出可用区域，停止绘制
                }
                
                // 检查是否为右对齐标记的行
                if (line.StartsWith("RIGHT_ALIGN:"))
                {
                    var content = line.Substring("RIGHT_ALIGN:".Length);
                    var textSize = e.Graphics.MeasureString(content, font);
                    var x = leftMargin + pageWidth - textSize.Width;
                    e.Graphics.DrawString(content, font, brush, x, y);
                }
                // 检查是否为需要左右对齐的行（包含冒号且不是装饰行）
                else if (line.Contains(':') && !line.Trim().StartsWith("_") && !line.Trim().StartsWith("-") && !line.Trim().StartsWith("="))
                {
                    DrawAlignedLine(e.Graphics, line, font, brush, leftMargin, y, pageWidth);
                }
                else
                {
                    // 普通行，居中或左对齐
                    if (line.Trim().StartsWith("=") || line.Trim().Contains("测试标签") || line.Trim().Contains("测试参数"))
                    {
                        // 标题行居中
                        var textSize = e.Graphics.MeasureString(line, font);
                        var x = leftMargin + (pageWidth - textSize.Width) / 2;
                        e.Graphics.DrawString(line, font, brush, x, y);
                    }
                    else
                    {
                        // 普通行左对齐
                        e.Graphics.DrawString(line, font, brush, leftMargin, y);
                    }
                }
            }

            // 绘制页脚
            if (_currentTemplate?.ShowFooter == true)
            {
                var footerY = e.PageBounds.Height - bottomMargin;
                DrawFooter(e.Graphics, font, brush, leftMargin, footerY, pageWidth);
            }

            e.HasMorePages = false;
        }

        private float DrawHeader(Graphics graphics, Font font, Brush brush, float leftMargin, float startY, float pageWidth)
        {
            var currentY = startY;
            var lineHeight = _currentFontSize + 8;

            try
            {
                // 绘制页眉图片
                if (!string.IsNullOrEmpty(_currentTemplate?.HeaderImagePath) && File.Exists(_currentTemplate.HeaderImagePath))
                {
                    try
                    {
                        using var headerImage = Image.FromFile(_currentTemplate.HeaderImagePath);
                        var imageHeight = 60; // 固定页眉图片高度
                        var imageWidth = (int)(headerImage.Width * ((float)imageHeight / headerImage.Height));
                        
                        // 居中显示图片
                        var imageX = leftMargin + (pageWidth - imageWidth) / 2;
                        graphics.DrawImage(headerImage, imageX, currentY, imageWidth, imageHeight);
                        currentY += imageHeight + lineHeight / 2;
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"无法加载页眉图片: {ex.Message}");
                    }
                }

                // 绘制页眉文本
                if (!string.IsNullOrEmpty(_currentTemplate?.HeaderText))
                {
                    var headerFont = new Font(_currentFontName, _currentFontSize + 2, FontStyle.Bold); // 页眉字体稍大
                    var textSize = graphics.MeasureString(_currentTemplate.HeaderText, headerFont);
                    var textX = leftMargin + (pageWidth - textSize.Width) / 2; // 居中
                    graphics.DrawString(_currentTemplate.HeaderText, headerFont, brush, textX, currentY);
                    currentY += textSize.Height + lineHeight / 2;
                }

                // 绘制分隔线
                if (!string.IsNullOrEmpty(_currentTemplate?.HeaderText) || !string.IsNullOrEmpty(_currentTemplate?.HeaderImagePath))
                {
                    using var pen = new Pen(Color.Gray, 1);
                    graphics.DrawLine(pen, leftMargin, currentY, leftMargin + pageWidth, currentY);
                    currentY += lineHeight / 2;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"绘制页眉时出错: {ex.Message}", ex);
            }

            return currentY;
        }

        private int CalculateFooterHeight(Graphics graphics, Font font, float pageWidth)
        {
            var height = 0;
            var lineHeight = _currentFontSize + 8;

            try
            {
                // 计算页脚图片高度
                if (!string.IsNullOrEmpty(_currentTemplate?.FooterImagePath) && File.Exists(_currentTemplate.FooterImagePath))
                {
                    height += 60 + (lineHeight / 2); // 固定页脚图片高度
                }

                // 计算页脚文本高度
                if (!string.IsNullOrEmpty(_currentTemplate?.FooterText))
                {
                    var footerFont = new Font(_currentFontName, _currentFontSize);
                    var textSize = graphics.MeasureString(_currentTemplate.FooterText, footerFont);
                    height += (int)Math.Ceiling(textSize.Height) + (lineHeight / 2);
                }

                // 分隔线高度
                if (height > 0)
                {
                    height += (lineHeight / 2);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"计算页脚高度时出错: {ex.Message}", ex);
            }

            return height;
        }

        private void DrawFooter(Graphics graphics, Font font, Brush brush, float leftMargin, float footerY, float pageWidth)
        {
            var currentY = footerY;
            var lineHeight = _currentFontSize + 8;

            try
            {
                // 绘制分隔线
                if (!string.IsNullOrEmpty(_currentTemplate?.FooterText) || !string.IsNullOrEmpty(_currentTemplate?.FooterImagePath))
                {
                    using var pen = new Pen(Color.Gray, 1);
                    currentY -= lineHeight / 2;
                    graphics.DrawLine(pen, leftMargin, currentY, leftMargin + pageWidth, currentY);
                    currentY -= lineHeight / 2;
                }

                // 绘制页脚文本
                if (!string.IsNullOrEmpty(_currentTemplate?.FooterText))
                {
                    var footerFont = new Font(_currentFontName, _currentFontSize);
                    var textSize = graphics.MeasureString(_currentTemplate.FooterText, footerFont);
                    var textX = leftMargin + (pageWidth - textSize.Width) / 2; // 居中
                    currentY -= textSize.Height;
                    graphics.DrawString(_currentTemplate.FooterText, footerFont, brush, textX, currentY);
                    currentY -= lineHeight / 2;
                }

                // 绘制页脚图片
                if (!string.IsNullOrEmpty(_currentTemplate?.FooterImagePath) && File.Exists(_currentTemplate.FooterImagePath))
                {
                    try
                    {
                        using var footerImage = Image.FromFile(_currentTemplate.FooterImagePath);
                        var imageHeight = 60; // 固定页脚图片高度
                        var imageWidth = (int)(footerImage.Width * ((float)imageHeight / footerImage.Height));
                        
                        // 居中显示图片
                        var imageX = leftMargin + (pageWidth - imageWidth) / 2;
                        currentY -= imageHeight;
                        graphics.DrawImage(footerImage, imageX, currentY, imageWidth, imageHeight);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"无法加载页脚图片: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"绘制页脚时出错: {ex.Message}", ex);
            }
        }

        public string GeneratePrintContent(TestRecord record, string templateName = "Default")
        {
            try
            {
                var template = PrintTemplateManager.GetTemplate(templateName) ?? PrintTemplateManager.GetDefaultTemplate();
                
                if (template == null)
                {
                    Logger.Warning($"未找到模板 '{templateName}'，使用默认内容");
                    return GenerateDefaultContent(record);
                }

                var content = template.Content;
                
                // 替换模板中的占位符（注意：单位已在模板中定义，这里只替换数值）
                content = content.Replace("{SerialNumber}", record.TR_SerialNum ?? "N/A");
                content = content.Replace("{TestDateTime}", record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
                content = content.Replace("{Current}", record.FormatNumber(record.TR_Isc));
                content = content.Replace("{Voltage}", record.FormatNumber(record.TR_Voc));
                content = content.Replace("{VoltageVpm}", record.FormatNumber(record.TR_Vpm));
                content = content.Replace("{Power}", record.FormatNumber(record.TR_Pm)); // 修正：应该是Pm而不是Ipm
                content = content.Replace("{PrintCount}", record.TR_Print.ToString());
                content = content.Replace("{CurrentTime}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                
                return content;
            }
            catch (Exception ex)
            {
                Logger.Error($"生成打印内容失败: {ex.Message}", ex);
                return GenerateDefaultContent(record);
            }
        }

        private string GenerateDefaultContent(TestRecord record)
        {
            return $@"
====================================
        太阳能电池测试标签
====================================
序列号: {record.TR_SerialNum ?? "N/A"}
测试时间: {record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A"}
短路电流: {record.FormatNumber(record.TR_Isc)} A
开路电压: {record.FormatNumber(record.TR_Voc)} V
最大功率电压: {record.FormatNumber(record.TR_Vpm)} V
 最大功率: {record.FormatNumber(record.TR_Ipm)} W
打印次数: {record.TR_Print}
====================================
";
        }

        // ProcessPrePrintedLabelPrint 方法已删除

        private void DrawAlignedLine(Graphics graphics, string line, Font font, Brush brush, float leftMargin, float y, float pageWidth)
        {
            var parts = line.Split(':', 2);
            if (parts.Length != 2) 
            {
                graphics.DrawString(line, font, brush, leftMargin, y);
                return;
            }

            var label = parts[0].Trim();
            var value = parts[1].Trim();

            // 测量字符串宽度
            var labelSize = graphics.MeasureString(label + ":", font);
            var valueSize = graphics.MeasureString(value, font);

            // 如果内容太长，进行换行处理
            if (labelSize.Width + valueSize.Width > pageWidth)
            {
                // 标签太长，换行
                if (labelSize.Width > pageWidth * 0.6)
                {
                    var wrappedLabel = WrapTextToWidth(graphics, label, font, pageWidth * 0.6f);
                    var labelLines = wrappedLabel.Split('\n');
                    
                    // 绘制标签行
                    var dynamicLineHeight = _currentFontSize + 8;
                    for (int i = 0; i < labelLines.Length; i++)
                    {
                        graphics.DrawString(labelLines[i], font, brush, leftMargin, y + i * dynamicLineHeight);
                    }
                    
                    // 在最后一行绘制冒号和值
                    var lastLineWidth = graphics.MeasureString(labelLines[labelLines.Length - 1], font).Width;
                    graphics.DrawString(":", font, brush, leftMargin + lastLineWidth, y + (labelLines.Length - 1) * dynamicLineHeight);
                    graphics.DrawString(value, font, brush, leftMargin + pageWidth - valueSize.Width, y + (labelLines.Length - 1) * dynamicLineHeight);
                }
                else
                {
                    // 值太长，换行
                    graphics.DrawString(label + ":", font, brush, leftMargin, y);
                    var wrappedValue = WrapTextToWidth(graphics, value, font, pageWidth * 0.7f);
                    var valueLines = wrappedValue.Split('\n');
                    var dynamicLineHeight = _currentFontSize + 8;
                    
                    for (int i = 0; i < valueLines.Length; i++)
                    {
                        var valueLineSize = graphics.MeasureString(valueLines[i], font);
                        graphics.DrawString(valueLines[i], font, brush, leftMargin + pageWidth - valueLineSize.Width, y + (i + 1) * dynamicLineHeight);
                    }
                }
            }
            else
            {
                // 正常情况：标签左对齐，值右对齐
                graphics.DrawString(label + ":", font, brush, leftMargin, y);
                graphics.DrawString(value, font, brush, leftMargin + pageWidth - valueSize.Width, y);
            }
        }

        private string WrapTextToWidth(Graphics graphics, string text, Font font, float maxWidth)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            var words = text.Split(' ');
            var lines = new List<string>();
            var currentLine = "";

            foreach (var word in words)
            {
                var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
                var testSize = graphics.MeasureString(testLine, font);

                if (testSize.Width <= maxWidth)
                {
                    currentLine = testLine;
                }
                else
                {
                    if (!string.IsNullOrEmpty(currentLine))
                    {
                        lines.Add(currentLine);
                        currentLine = word;
                    }
                    else
                    {
                        // 单词太长，强制分割
                        lines.Add(word);
                        currentLine = "";
                    }
                }
            }

            if (!string.IsNullOrEmpty(currentLine))
            {
                lines.Add(currentLine);
            }

            return string.Join("\n", lines);
        }

        private string GenerateTextReceipt(TestRecord record)
        {
            var sb = new StringBuilder();
            var date = record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var time = DateTime.Now.ToString("HH:mm:ss");

            sb.AppendLine("===============================");
            sb.AppendLine("    太阳能电池测试标签");
            sb.AppendLine("===============================");
            sb.AppendLine("");
            sb.AppendLine($"序列号: {record.TR_SerialNum ?? "N/A"}");
            sb.AppendLine($"测试日期: {date}");
            sb.AppendLine($"打印时间: {time}");
            sb.AppendLine("");
            sb.AppendLine("======== 测试参数 ========");
            sb.AppendLine($"短路电流 (Isc): {record.FormatNumber(record.TR_Isc)} A");
            sb.AppendLine($"最大功率电流 (Ipm): {record.FormatNumber(record.TR_Ipm)} A");
            sb.AppendLine($"开路电压 (Voc): {record.FormatNumber(record.TR_Voc)} V");
            sb.AppendLine($"最大功率电压 (Vpm): {record.FormatNumber(record.TR_Vpm)} V");
            sb.AppendLine($"最大功率 (Pm): {record.FormatNumber(record.TR_Pm)} W");
            sb.AppendLine("");
            sb.AppendLine($"条码: {record.TR_SerialNum ?? "N/A"}");
            sb.AppendLine("备注: 太阳能电池测试数据");
            sb.AppendLine("");
            sb.AppendLine("===============================");

            return sb.ToString();
        }

        private string GenerateZPLCode(TestRecord record)
        {
            var serialNum = record.TR_SerialNum ?? "N/A";
            var date = record.TR_DateTime?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");

            return $@"^XA
^FO20,20^A0N,25,25^FD太阳能电池测试标签^FS
^FO20,50^A0N,20,20^FD序列号: {serialNum}^FS
^FO20,80^A0N,20,20^FD日期: {date}^FS
^FO20,110^A0N,20,20^FDIsc: {record.FormatNumber(record.TR_Isc)} A^FS
^FO20,140^A0N,20,20^FDIpm: {record.FormatNumber(record.TR_Ipm)} A^FS
^FO20,170^A0N,20,20^FDVoc: {record.FormatNumber(record.TR_Voc)} V^FS
^FO20,200^A0N,20,20^FDVpm: {record.FormatNumber(record.TR_Vpm)} V^FS
^FO20,230^A0N,20,20^FDPm: {record.FormatNumber(record.TR_Pm)} W^FS
^FO20,270^BY3^BCN,70,Y,N,N^FD{serialNum}^FS
^XZ".Trim();
        }

        private string GenerateCode128ZPL(TestRecord record)
        {
            var serialNum = record.TR_SerialNum ?? "N/A";
            var date = record.TR_DateTime?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");

            return $@"^XA
^FO50,50^A0N,25,25^FD太阳能电池测试标签^FS
^FO50,80^A0N,20,20^FD序列号: {serialNum}^FS
^FO50,110^A0N,20,20^FD日期: {date}^FS
^FO50,140^A0N,20,20^FDIsc: {record.FormatNumber(record.TR_Isc)} A^FS
^FO50,170^A0N,20,20^FDIpm: {record.FormatNumber(record.TR_Ipm)} A^FS
^FO50,200^A0N,20,20^FDVoc: {record.FormatNumber(record.TR_Voc)} V^FS
^FO50,230^A0N,20,20^FDVpm: {record.FormatNumber(record.TR_Vpm)} V^FS
^FO50,260^A0N,20,20^FDPm: {record.FormatNumber(record.TR_Pm)} W^FS
^FO50,300^BY2^BCN,60,Y,N,N^FD{serialNum}^FS
^XZ".Trim();
        }

        private string GenerateQRCodeZPL(TestRecord record)
        {
            var serialNum = record.TR_SerialNum ?? "N/A";
            var date = record.TR_DateTime?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");
            var qrData = $"SN:{serialNum},Isc:{record.FormatNumber(record.TR_Isc)},Voc:{record.FormatNumber(record.TR_Voc)},Pm:{record.FormatNumber(record.TR_Pm)}";

            return $@"^XA
^FO50,50^A0N,25,25^FD太阳能电池测试标签^FS
^FO50,80^A0N,20,20^FD序列号: {serialNum}^FS
^FO50,110^A0N,20,20^FD日期: {date}^FS
^FO50,140^A0N,20,20^FDIsc: {record.FormatNumber(record.TR_Isc)} A^FS
^FO50,170^A0N,20,20^FDIpm: {record.FormatNumber(record.TR_Ipm)} A^FS
^FO50,200^A0N,20,20^FDVoc: {record.FormatNumber(record.TR_Voc)} V^FS
^FO50,230^A0N,20,20^FDVpm: {record.FormatNumber(record.TR_Vpm)} V^FS
^FO50,260^A0N,20,20^FDPm: {record.FormatNumber(record.TR_Pm)} W^FS
^FO50,300^BQN,2,4^FDQA,{qrData}^FS
^XZ".Trim();
        }

        private PrintResult SendRawDataToPrinter(string data, string method)
        {
            try
            {
                var tempFile = Path.GetTempFileName();
                File.WriteAllText(tempFile, data);

                // 对于ZPL数据，尝试直接发送到打印机
                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c copy /b \"{tempFile}\" \"\\\\localhost\\{_currentPrinterName}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = System.Diagnostics.Process.Start(processInfo);
                process?.WaitForExit(5000);

                File.Delete(tempFile);

                return new PrintResult
                {
                    Success = true,
                    Method = method,
                    JobId = $"{method}_{DateTime.Now:yyyyMMddHHmmss}",
                    PrinterUsed = _currentPrinterName
                };
            }
            catch (Exception ex)
            {
                Logger.Error($"发送原始数据到打印机失败: {ex.Message}", ex);
                throw;
            }
        }

        public PrintResult TestPrint()
        {
            var testRecord = new TestRecord
            {
                TR_SerialNum = "TEST" + DateTime.Now.ToString("yyyyMMddHHmmss"),
                TR_Isc = 5.358m,
                TR_Ipm = 4.823m,
                TR_Voc = 26.086m,
                TR_Vpm = 25.123m,
                TR_Pm = 121.168m,
                TR_DateTime = DateTime.Now
            };

            return PrintRecord(testRecord, "Text");
        }

        public static List<string> GetSystemFonts()
        {
            var fonts = new List<string>();
            
            try
            {
                using var installedFonts = new System.Drawing.Text.InstalledFontCollection();
                foreach (var fontFamily in installedFonts.Families)
                {
                    try
                    {
                        // 只添加支持常规样式的字体
                        if (fontFamily.IsStyleAvailable(FontStyle.Regular))
                        {
                            fonts.Add(fontFamily.Name);
                        }
                    }
                    catch
                    {
                        // 忽略无法访问的字体
                    }
                }
                
                // 按名称排序
                fonts.Sort();
                
                Logger.Info($"已获取 {fonts.Count} 个系统字体");
            }
            catch (Exception ex)
            {
                Logger.Error($"获取系统字体失败: {ex.Message}", ex);
                
                // 提供默认字体列表
                fonts.AddRange(new[] { "Arial", "Times New Roman", "Calibri", "Tahoma", "Verdana", "Microsoft YaHei", "SimSun" });
            }
            
            return fonts;
        }
    }
} 