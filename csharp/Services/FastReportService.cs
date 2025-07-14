using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using FastReport;
using FastReport.Export.PdfSimple;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Services
{
    /// <summary>
    /// FastReport开源版本报表服务
    /// </summary>
    public class FastReportService
    {
        private readonly string _reportsPath;
        private Report _report;

        public FastReportService()
        {
            _reportsPath = Path.Combine(Environment.CurrentDirectory, "Reports");
            if (!Directory.Exists(_reportsPath))
            {
                Directory.CreateDirectory(_reportsPath);
            }

            _report = new Report();
            InitializeReport();
        }

        private void InitializeReport()
        {
            try
            {
                // 创建数据源
                DataSet dataSet = new DataSet();
                DataTable table = CreateTestRecordDataTable();
                dataSet.Tables.Add(table);
                
                _report.RegisterData(dataSet, "TestData");
                _report.GetDataSource("TestData").Enabled = true;
                
                Logger.Info("FastReport开源版本服务初始化完成");
            }
            catch (Exception ex)
            {
                Logger.Error($"FastReport服务初始化失败: {ex.Message}", ex);
            }
        }

        private DataTable CreateTestRecordDataTable()
        {
            DataTable table = new DataTable("TestRecord");
            
            // 添加列
            table.Columns.Add("SerialNumber", typeof(string));
            table.Columns.Add("TestDateTime", typeof(DateTime));
            table.Columns.Add("Current", typeof(decimal));
            table.Columns.Add("CurrentImp", typeof(decimal));
            table.Columns.Add("Voltage", typeof(decimal));
            table.Columns.Add("VoltageVpm", typeof(decimal));
            table.Columns.Add("Power", typeof(decimal));
            table.Columns.Add("PrintCount", typeof(int));
            table.Columns.Add("CurrentTime", typeof(DateTime));
            table.Columns.Add("CurrentDate", typeof(DateTime));

            return table;
        }

        /// <summary>
        /// 创建简单的报表模板
        /// </summary>
        public void CreateSimpleTemplate()
        {
            try
            {
                string templatePath = Path.Combine(_reportsPath, "SolarPanelTestReport.frx");
                
                if (!File.Exists(templatePath))
                {
                    // 创建一个基本的报表模板
                    string basicTemplate = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report ScriptLanguage=""CSharp"" ReportInfo.Created=""01/01/2025 00:00:00"" ReportInfo.Modified=""01/01/2025 00:00:00"" ReportInfo.CreatorVersion=""2025.2.0.0"">
  <Dictionary>
    <BusinessObjectDataSource Name=""TestData"" ReferenceName=""TestData"" DataType=""System.Int32"" Enabled=""true"">
      <Column Name=""SerialNumber"" DataType=""System.String""/>
      <Column Name=""TestDateTime"" DataType=""System.DateTime""/>
      <Column Name=""Current"" DataType=""System.Decimal""/>
      <Column Name=""CurrentImp"" DataType=""System.Decimal""/>
      <Column Name=""Voltage"" DataType=""System.Decimal""/>
      <Column Name=""VoltageVpm"" DataType=""System.Decimal""/>
      <Column Name=""Power"" DataType=""System.Decimal""/>
      <Column Name=""PrintCount"" DataType=""System.Int32""/>
      <Column Name=""CurrentTime"" DataType=""System.DateTime""/>
      <Column Name=""CurrentDate"" DataType=""System.DateTime""/>
    </BusinessObjectDataSource>
  </Dictionary>
  <ReportPage Name=""Page1"" Watermark.Font=""Arial, 60pt"">
    <ReportTitleBand Name=""ReportTitle1"" Width=""718.2"" Height=""75.6"">
      <TextObject Name=""Text1"" Left=""189"" Top=""18.9"" Width=""340.2"" Height=""37.8"" Text=""太阳能电池测试报告"" HorzAlign=""Center"" VertAlign=""Center"" Font=""Arial, 16pt, style=Bold""/>
    </ReportTitleBand>
    <PageHeaderBand Name=""PageHeader1"" Top=""79.6"" Width=""718.2"" Height=""56.7"">
      <TextObject Name=""Text2"" Left=""28.35"" Top=""18.9"" Width=""75.6"" Height=""18.9"" Text=""序列号"" Font=""Arial, 10pt, style=Bold""/>
      <TextObject Name=""Text3"" Left=""113.4"" Top=""18.9"" Width=""75.6"" Height=""18.9"" Text=""测试时间"" Font=""Arial, 10pt, style=Bold""/>
      <TextObject Name=""Text4"" Left=""198.45"" Top=""18.9"" Width=""75.6"" Height=""18.9"" Text=""短路电流"" Font=""Arial, 10pt, style=Bold""/>
      <TextObject Name=""Text5"" Left=""283.5"" Top=""18.9"" Width=""75.6"" Height=""18.9"" Text=""开路电压"" Font=""Arial, 10pt, style=Bold""/>
      <TextObject Name=""Text6"" Left=""368.55"" Top=""18.9"" Width=""75.6"" Height=""18.9"" Text=""最大功率"" Font=""Arial, 10pt, style=Bold""/>
    </PageHeaderBand>
    <DataBand Name=""Data1"" Top=""140.3"" Width=""718.2"" Height=""28.35"" DataSource=""TestData"">
      <TextObject Name=""Text7"" Left=""28.35"" Top=""9.45"" Width=""75.6"" Height=""18.9"" Text=""[TestData.SerialNumber]"" Font=""Arial, 9pt""/>
      <TextObject Name=""Text8"" Left=""113.4"" Top=""9.45"" Width=""75.6"" Height=""18.9"" Text=""[TestData.TestDateTime]"" Format=""Date"" Format.Format=""d"" Font=""Arial, 9pt""/>
      <TextObject Name=""Text9"" Left=""198.45"" Top=""9.45"" Width=""75.6"" Height=""18.9"" Text=""[TestData.Current]"" Format=""Number"" Format.UseLocale=""false"" Format.DecimalDigits=""3"" Format.DecimalSeparator=""."" Format.GroupSeparator="","" Format.NegativePattern=""1"" Font=""Arial, 9pt""/>
      <TextObject Name=""Text10"" Left=""283.5"" Top=""9.45"" Width=""75.6"" Height=""18.9"" Text=""[TestData.Voltage]"" Format=""Number"" Format.UseLocale=""false"" Format.DecimalDigits=""3"" Format.DecimalSeparator=""."" Format.GroupSeparator="","" Format.NegativePattern=""1"" Font=""Arial, 9pt""/>
      <TextObject Name=""Text11"" Left=""368.55"" Top=""9.45"" Width=""75.6"" Height=""18.9"" Text=""[TestData.Power]"" Format=""Number"" Format.UseLocale=""false"" Format.DecimalDigits=""3"" Format.DecimalSeparator=""."" Format.GroupSeparator="","" Format.NegativePattern=""1"" Font=""Arial, 9pt""/>
    </DataBand>
    <PageFooterBand Name=""PageFooter1"" Top=""172.65"" Width=""718.2"" Height=""37.8"">
      <TextObject Name=""Text12"" Left=""28.35"" Top=""9.45"" Width=""151.2"" Height=""18.9"" Text=""打印时间: [Date]"" Font=""Arial, 9pt""/>
      <TextObject Name=""Text13"" Left=""548.1"" Top=""9.45"" Width=""141.75"" Height=""18.9"" Text=""页码: [Page]"" HorzAlign=""Right"" Font=""Arial, 9pt""/>
    </PageFooterBand>
  </ReportPage>
</Report>";

                    File.WriteAllText(templatePath, basicTemplate);
                    Logger.Info("默认报表模板已创建");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"创建报表模板失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 预览报表（导出为PDF查看）
        /// </summary>
        public void PreviewReport(TestRecord record)
        {
            try
            {
                // 确保模板存在
                CreateSimpleTemplate();
                
                // 加载模板
                string templatePath = Path.Combine(_reportsPath, "SolarPanelTestReport.frx");
                _report.Load(templatePath);
                
                // 填充数据
                FillReportData(record);
                
                // 准备报表
                _report.Prepare();
                
                // 导出为临时PDF文件并打开
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"Report_Preview_{DateTime.Now:yyyyMMddHHmmss}.pdf");
                PDFSimpleExport pdfExport = new PDFSimpleExport();
                _report.Export(pdfExport, tempPdfPath);
                
                // 打开PDF文件
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = tempPdfPath,
                    UseShellExecute = true
                });
                
                Logger.Info($"报表预览已生成: {tempPdfPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"预览报表失败: {ex.Message}", ex);
                MessageBox.Show($"预览报表失败: {ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出报表为PDF
        /// </summary>
        public bool ExportToPdf(TestRecord record, string fileName)
        {
            try
            {
                // 确保模板存在
                CreateSimpleTemplate();
                
                // 加载模板
                string templatePath = Path.Combine(_reportsPath, "SolarPanelTestReport.frx");
                _report.Load(templatePath);
                
                // 填充数据
                FillReportData(record);
                
                // 准备报表
                _report.Prepare();
                
                // 导出为PDF
                PDFSimpleExport pdfExport = new PDFSimpleExport();
                string outputPath = Path.Combine(_reportsPath, fileName.EndsWith(".pdf") ? fileName : fileName + ".pdf");
                _report.Export(pdfExport, outputPath);
                
                Logger.Info($"报表已导出为PDF: {outputPath}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"导出PDF失败: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// 生成基本报表内容（HTML格式）
        /// </summary>
        public string GenerateBasicReport(TestRecord record)
        {
            try
            {
                string html = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>太阳能电池测试报告</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        .header {{ text-align: center; margin-bottom: 30px; }}
        .data-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        .data-table th, .data-table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        .data-table th {{ background-color: #f2f2f2; font-weight: bold; }}
        .footer {{ margin-top: 30px; font-size: 12px; color: #666; }}
    </style>
</head>
<body>
    <div class='header'>
        <h2>太阳能电池测试报告</h2>
        <p>序列号: {record.TR_SerialNum ?? "N/A"}</p>
        <p>测试时间: {record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}</p>
    </div>
    
    <table class='data-table'>
        <tr><th>测试项目</th><th>测试值</th><th>单位</th></tr>
        <tr><td>短路电流 (Isc)</td><td>{record.FormatNumber(record.TR_Isc)}</td><td>A</td></tr>
        <tr><td>最大功率电流 (Imp)</td><td>{record.FormatNumber(record.TR_Ipm)}</td><td>A</td></tr>
        <tr><td>开路电压 (Voc)</td><td>{record.FormatNumber(record.TR_Voc)}</td><td>V</td></tr>
        <tr><td>最大功率电压 (Vmp)</td><td>{record.FormatNumber(record.TR_Vpm)}</td><td>V</td></tr>
        <tr><td>最大功率 (Pm)</td><td>{record.FormatNumber(record.TR_Pm)}</td><td>W</td></tr>
        <tr><td>打印次数</td><td>{record.TR_Print ?? 0}</td><td>次</td></tr>
    </table>
    
    <div class='footer'>
        <p>报表生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>
        <p>标准测试条件: 1000W/m², AM1.5, 25°C</p>
        <p>生成工具: ZebraPrinterMonitor v1.1.30 (FastReport.OpenSource)</p>
    </div>
</body>
</html>";

                return html;
            }
            catch (Exception ex)
            {
                Logger.Error($"生成基本报表失败: {ex.Message}", ex);
                return $"<html><body><h1>报表生成错误</h1><p>{ex.Message}</p></body></html>";
            }
        }

        private void FillReportData(TestRecord record)
        {
            DataSet dataSet = new DataSet();
            DataTable table = CreateTestRecordDataTable();
            
            // 添加数据行
            DataRow row = table.NewRow();
            row["SerialNumber"] = record.TR_SerialNum ?? "N/A";
            row["TestDateTime"] = record.TR_DateTime ?? DateTime.Now;
            row["Current"] = record.TR_Isc;
            row["CurrentImp"] = record.TR_Ipm;
            row["Voltage"] = record.TR_Voc;
            row["VoltageVpm"] = record.TR_Vpm;
            row["Power"] = record.TR_Pm;
            row["PrintCount"] = record.TR_Print ?? 0;
            row["CurrentTime"] = DateTime.Now;
            row["CurrentDate"] = DateTime.Now.Date;
            
            table.Rows.Add(row);
            dataSet.Tables.Add(table);
            
            // 重新注册数据
            _report.RegisterData(dataSet, "TestData");
        }

        /// <summary>
        /// 获取可用报表模板列表
        /// </summary>
        public string[] GetAvailableReports()
        {
            try
            {
                var files = Directory.GetFiles(_reportsPath, "*.frx");
                var reportNames = new string[files.Length];
                
                for (int i = 0; i < files.Length; i++)
                {
                    reportNames[i] = Path.GetFileNameWithoutExtension(files[i]);
                }
                
                return reportNames;
            }
            catch (Exception ex)
            {
                Logger.Error($"获取报表列表失败: {ex.Message}", ex);
                return new string[0];
            }
        }

        /// <summary>
        /// 显示报表模板说明信息
        /// </summary>
        public void ShowTemplateInfo(Form parentForm)
        {
            try
            {
                string info = @"FastReport.OpenSource 报表功能说明：

1. 基本功能：
   - 自动生成简单的报表模板
   - 支持测试数据预览（导出为PDF）
   - 导出PDF格式
   - 生成HTML报表内容

2. 模板位置：
   " + _reportsPath + @"

3. 使用说明：
   - 点击'报表预览'查看报表效果（PDF）
   - 使用'导出PDF'保存报表文件
   - 报表模板可手动编辑(.frx文件)

4. 开源版本限制：
   - 不包含可视化设计器
   - 无直接打印功能
   - 基本的PDF导出功能

如需完整设计功能请考虑商业版本。";

                MessageBox.Show(info, "FastReport 报表说明", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"显示模板信息失败: {ex.Message}", ex);
            }
        }

        public void Dispose()
        {
            _report?.Dispose();
        }
    }
} 