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
            if (string.IsNullOrEmpty(_currentPrintContent) || e.Graphics == null)
                return;

            var font = new Font("Microsoft YaHei", 10);
            var brush = Brushes.Black;
            var startY = 50;
            var lineHeight = 20;

            var lines = _currentPrintContent.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                e.Graphics.DrawString(lines[i], font, brush, 50, startY + (i * lineHeight));
            }

            e.HasMorePages = false;
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
    }
} 