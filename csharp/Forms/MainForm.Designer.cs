using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Forms
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private TabControl tabControl1;
        private TabPage tabMonitor, tabConfig, tabLogs, tabTemplate;
        private GroupBox grpDatabaseConfig, grpMonitorControl, grpStatus, grpRecords;
        private GroupBox grpPrinterConfig, grpApplicationConfig, grpLanguageConfig;
        
        private Label lblDatabasePath, lblPrinter, lblPrintFormat, lblPollInterval, lblLanguage;
        private TextBox txtDatabasePath, txtLog;
        private Button btnBrowseDatabase, btnTestConnection, btnStartMonitoring, btnStopMonitoring;
        private Button btnManualPrint, btnRefresh, btnTestPrint, btnClearLog, btnSaveLog;
        private CheckBox chkAutoPrint, chkAutoStartMonitoring, chkMinimizeToTray;
        private CheckBox chkEnablePrintCount;  // 新增打印次数控制复选框
        private Button btnShowPrintMonitor;    // 新增小监控窗口按钮
        private ComboBox cmbPrinter, cmbPrintFormat, cmbLanguage;
        private NumericUpDown numPollInterval;
        private ListView lvRecords;
        private Label lblMonitoringStatus, lblTotalRecords, lblTotalPrints, lblLastRecord, lblPrinterStatus;
        
        // 打印模板页面控件
        private GroupBox grpTemplateEditor, grpTemplateList, grpTemplatePreview;
        private Label lblTemplateName, lblTemplateContent, lblTemplateFormat, lblAvailableFields;
        private TextBox txtTemplateName, txtTemplateContent;
        private ComboBox cmbTemplateFormat, cmbTemplateList;
        private Button btnSaveTemplate, btnDeleteTemplate, btnPreviewTemplate, btnNewTemplate;
        private ListBox lstAvailableFields;
        private RichTextBox rtbTemplatePreview;



        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new TabControl();
            this.tabMonitor = new TabPage();
            this.tabConfig = new TabPage();
            this.tabLogs = new TabPage();
            this.tabTemplate = new TabPage();
            
            this.SuspendLayout();
            this.tabControl1.SuspendLayout();
            
            // 主窗体设置 - 调整尺寸确保所有控件可见
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1200, 850);  // 增加高度到850确保按钮可见
            this.Controls.Add(this.tabControl1);
            this.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimumSize = new Size(1000, 650);  // 调整最小尺寸
            this.Name = "MainForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "太阳能电池测试打印监控系统 v1.1.27";
            
            // TabControl设置
            this.tabControl1.Controls.Add(this.tabMonitor);
            this.tabControl1.Controls.Add(this.tabConfig);
            this.tabControl1.Controls.Add(this.tabTemplate);
            this.tabControl1.Controls.Add(this.tabLogs);
            this.tabControl1.Dock = DockStyle.Fill;
            this.tabControl1.Location = new Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new Size(1200, 850);
            this.tabControl1.TabIndex = 0;
            
            InitializeMonitorTab();
            InitializeConfigTab();
            InitializeTemplateTab();
            InitializeLogsTab();
            
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "太阳能电池测试打印监控系统 v1.1.27";
            
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private void InitializeMonitorTab()
        {
            this.tabMonitor.Text = "数据监控";
            this.tabMonitor.UseVisualStyleBackColor = true;
            
            // 数据库配置组 - 调整大小和位置
            this.grpDatabaseConfig = new GroupBox { Text = LanguageManager.GetString("DatabaseConfig"), Location = new Point(10, 10), Size = new Size(1160, 80) };
            this.lblDatabasePath = new Label { Text = LanguageManager.GetString("DatabasePath"), Location = new Point(15, 25), AutoSize = true };
            this.txtDatabasePath = new TextBox { Location = new Point(15, 45), Size = new Size(850, 23) };
            this.btnBrowseDatabase = new Button { Text = LanguageManager.GetString("Browse"), Location = new Point(875, 43), Size = new Size(80, 27) };
            this.btnTestConnection = new Button { Text = LanguageManager.GetString("TestConnection"), Location = new Point(965, 43), Size = new Size(100, 27) };
            
            this.btnBrowseDatabase.Click += btnBrowseDatabase_Click;
            this.btnTestConnection.Click += btnTestConnection_Click;
            
            this.grpDatabaseConfig.Controls.AddRange(new Control[] { lblDatabasePath, txtDatabasePath, btnBrowseDatabase, btnTestConnection });
            
            // 监控控制组 - 调整大小
            this.grpMonitorControl = new GroupBox { Text = LanguageManager.GetString("MonitoringControl"), Location = new Point(10, 100), Size = new Size(1160, 65) };
            this.btnStartMonitoring = new Button { Text = LanguageManager.GetString("StartMonitoring"), Location = new Point(15, 25), Size = new Size(100, 30) };
            this.btnStopMonitoring = new Button { Text = LanguageManager.GetString("StopMonitoring"), Location = new Point(125, 25), Size = new Size(100, 30) };
            this.btnShowPrintMonitor = new Button { Text = LanguageManager.GetString("PrintMonitorTitle"), Location = new Point(235, 25), Size = new Size(100, 30) };
            this.chkAutoPrint = new CheckBox { Text = LanguageManager.GetString("AutoPrint"), Location = new Point(350, 30), Checked = true, AutoSize = true };
            this.chkEnablePrintCount = new CheckBox { Text = LanguageManager.GetString("EnablePrintCount"), Location = new Point(470, 30), Checked = false, AutoSize = true };
            
            this.btnStartMonitoring.Click += btnStartMonitoring_Click;
            this.btnStopMonitoring.Click += btnStopMonitoring_Click;
            this.chkEnablePrintCount.CheckedChanged += chkEnablePrintCount_CheckedChanged;
            this.btnShowPrintMonitor.Click += btnShowPrintMonitor_Click;
            
            this.grpMonitorControl.Controls.AddRange(new Control[] { btnStartMonitoring, btnStopMonitoring, btnShowPrintMonitor, chkAutoPrint, chkEnablePrintCount });
            
            // 状态信息组 - 调整大小和布局
            this.grpStatus = new GroupBox { Text = LanguageManager.GetString("StatusInfo"), Location = new Point(10, 175), Size = new Size(1160, 85) };
            this.lblMonitoringStatus = new Label { Text = LanguageManager.GetString("MonitoringStatus") + " " + LanguageManager.GetString("Stopped"), Location = new Point(15, 25), ForeColor = Color.Red, Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold), AutoSize = true };
            this.lblTotalRecords = new Label { Text = LanguageManager.GetString("TotalRecords") + " 0", Location = new Point(15, 50), AutoSize = true };
            this.lblTotalPrints = new Label { Text = LanguageManager.GetString("TotalPrintJobs") + " 0", Location = new Point(200, 50), AutoSize = true };
            this.lblLastRecord = new Label { Text = LanguageManager.GetString("LastRecord") + " N/A", Location = new Point(400, 50), AutoSize = true };
            
            this.grpStatus.Controls.AddRange(new Control[] { lblMonitoringStatus, lblTotalRecords, lblTotalPrints, lblLastRecord });
            
            // 记录列表组 - 优化尺寸显示20条记录并确保按钮可见
            this.grpRecords = new GroupBox { Text = LanguageManager.GetString("RecentRecords"), Location = new Point(10, 270), Size = new Size(1160, 520) };
            this.lvRecords = new ListView 
            { 
                Location = new Point(15, 25), 
                Size = new Size(1130, 450), 
                View = View.Details, 
                FullRowSelect = true, 
                GridLines = true,
                MultiSelect = false,
                HideSelection = false
            };
            
            // 重新定义列，添加Vpm和打印次数列
            this.lvRecords.Columns.Add(LanguageManager.GetString("SerialNumber"), 150);
            this.lvRecords.Columns.Add(LanguageManager.GetString("TestDateTime"), 160);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Current"), 90);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Voltage"), 90);
            this.lvRecords.Columns.Add(LanguageManager.GetString("VoltageVpm"), 90);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Power"), 90);
            this.lvRecords.Columns.Add(LanguageManager.GetString("PrintCount"), 80);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Operation"), 120);
            
            // 添加双击事件用于打印
            this.lvRecords.DoubleClick += LvRecords_DoubleClick;
            this.lvRecords.MouseClick += LvRecords_MouseClick;
            
            this.btnManualPrint = new Button { Text = LanguageManager.GetString("ManualPrint"), Location = new Point(15, 485), Size = new Size(100, 30) };
            this.btnRefresh = new Button { Text = LanguageManager.GetString("Refresh"), Location = new Point(125, 485), Size = new Size(80, 30) };
            
            this.btnManualPrint.Click += btnManualPrint_Click;
            this.btnRefresh.Click += btnRefresh_Click;
            
            this.grpRecords.Controls.AddRange(new Control[] { lvRecords, btnManualPrint, btnRefresh });
            
            this.tabMonitor.Controls.AddRange(new Control[] { grpDatabaseConfig, grpMonitorControl, grpStatus, grpRecords });
        }

        private void InitializeConfigTab()
        {
            this.tabConfig.Text = "系统配置";
            this.tabConfig.UseVisualStyleBackColor = true;
            
            // 打印机配置组 - 调整尺寸
            this.grpPrinterConfig = new GroupBox { Text = LanguageManager.GetString("PrinterConfig"), Location = new Point(10, 10), Size = new Size(1160, 120) };
            this.lblPrinter = new Label { Text = LanguageManager.GetString("SelectedPrinter"), Location = new Point(15, 30), AutoSize = true };
            this.cmbPrinter = new ComboBox { Location = new Point(15, 50), Size = new Size(400, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.lblPrintFormat = new Label { Text = LanguageManager.GetString("PrintFormat"), Location = new Point(450, 30), AutoSize = true };
            this.cmbPrintFormat = new ComboBox { Location = new Point(450, 50), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.btnTestPrint = new Button { Text = LanguageManager.GetString("TestPrint"), Location = new Point(680, 50), Size = new Size(100, 30) };
            this.lblPrinterStatus = new Label { Text = LanguageManager.GetString("PrinterStatus") + " 未知", Location = new Point(15, 85), AutoSize = true };
            
            this.cmbPrintFormat.Items.AddRange(new string[] { "Text", "ZPL", "Code128", "QRCode" });
            this.cmbPrintFormat.SelectedIndex = 0;
            
            this.cmbPrinter.SelectedIndexChanged += cmbPrinter_SelectedIndexChanged;
            this.cmbPrintFormat.SelectedIndexChanged += cmbPrintFormat_SelectedIndexChanged;
            this.btnTestPrint.Click += btnTestPrint_Click;
            
            this.grpPrinterConfig.Controls.AddRange(new Control[] { lblPrinter, cmbPrinter, lblPrintFormat, cmbPrintFormat, btnTestPrint, lblPrinterStatus });
            
            // 应用程序配置组 - 调整尺寸
            this.grpApplicationConfig = new GroupBox { Text = LanguageManager.GetString("ApplicationConfig"), Location = new Point(10, 140), Size = new Size(1160, 120) };
            this.lblPollInterval = new Label { Text = LanguageManager.GetString("PollInterval"), Location = new Point(15, 30), AutoSize = true };
            this.numPollInterval = new NumericUpDown { Location = new Point(15, 50), Size = new Size(120, 25), Minimum = 500, Maximum = 60000, Value = 1000, Increment = 500 };
            this.chkAutoStartMonitoring = new CheckBox { Text = LanguageManager.GetString("AutoStartMonitoring"), Location = new Point(200, 53), AutoSize = true };
            this.chkMinimizeToTray = new CheckBox { Text = LanguageManager.GetString("MinimizeToTray"), Location = new Point(200, 83), Checked = true, AutoSize = true };
            
            this.numPollInterval.ValueChanged += numPollInterval_ValueChanged;
            
            this.grpApplicationConfig.Controls.AddRange(new Control[] { lblPollInterval, numPollInterval, chkAutoStartMonitoring, chkMinimizeToTray });
            
            // 语言配置组
            this.grpLanguageConfig = new GroupBox { Text = LanguageManager.GetString("LanguageConfig"), Location = new Point(10, 270), Size = new Size(1160, 80) };
            this.lblLanguage = new Label { Text = LanguageManager.GetString("Language"), Location = new Point(15, 30), AutoSize = true };
            this.cmbLanguage = new ComboBox { Location = new Point(15, 50), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            
            this.cmbLanguage.Items.Add("简体中文");
            this.cmbLanguage.Items.Add("English");
            this.cmbLanguage.SelectedIndex = 0;
            this.cmbLanguage.SelectedIndexChanged += cmbLanguage_SelectedIndexChanged;
            
            this.grpLanguageConfig.Controls.AddRange(new Control[] { lblLanguage, cmbLanguage });
            
            this.tabConfig.Controls.AddRange(new Control[] { grpPrinterConfig, grpApplicationConfig, grpLanguageConfig });
        }

        private void InitializeTemplateTab()
        {
            this.tabTemplate.Text = "打印模板";
            this.tabTemplate.UseVisualStyleBackColor = true;
            
            // 模板列表组
            this.grpTemplateList = new GroupBox { Text = LanguageManager.GetString("TemplateList"), Location = new Point(10, 10), Size = new Size(300, 350) };
            this.cmbTemplateList = new ComboBox { Location = new Point(15, 30), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.btnNewTemplate = new Button { Text = LanguageManager.GetString("NewTemplate"), Location = new Point(15, 70), Size = new Size(80, 30) };
            this.btnDeleteTemplate = new Button { Text = LanguageManager.GetString("DeleteTemplate"), Location = new Point(110, 70), Size = new Size(80, 30) };
            
            this.cmbTemplateList.SelectedIndexChanged += cmbTemplateList_SelectedIndexChanged;
            this.btnNewTemplate.Click += btnNewTemplate_Click;
            this.btnDeleteTemplate.Click += btnDeleteTemplate_Click;
            
            this.grpTemplateList.Controls.AddRange(new Control[] { cmbTemplateList, btnNewTemplate, btnDeleteTemplate });
            
            // 模板编辑器组
            this.grpTemplateEditor = new GroupBox { Text = LanguageManager.GetString("TemplateEditor"), Location = new Point(320, 10), Size = new Size(550, 350) };
            this.lblTemplateName = new Label { Text = LanguageManager.GetString("TemplateName"), Location = new Point(15, 30), AutoSize = true };
            this.txtTemplateName = new TextBox { Location = new Point(15, 50), Size = new Size(300, 25) };
            this.lblTemplateFormat = new Label { Text = LanguageManager.GetString("TemplateFormat"), Location = new Point(350, 30), AutoSize = true };
            this.cmbTemplateFormat = new ComboBox { Location = new Point(350, 50), Size = new Size(150, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.lblTemplateContent = new Label { Text = LanguageManager.GetString("TemplateContent"), Location = new Point(15, 85), AutoSize = true };
            this.txtTemplateContent = new TextBox { Location = new Point(15, 110), Size = new Size(520, 180), Multiline = true, ScrollBars = ScrollBars.Both };
            this.btnSaveTemplate = new Button { Text = LanguageManager.GetString("SaveTemplate"), Location = new Point(15, 305), Size = new Size(80, 30) };
            this.btnPreviewTemplate = new Button { Text = LanguageManager.GetString("PreviewTemplate"), Location = new Point(110, 305), Size = new Size(80, 30) };
            
            this.cmbTemplateFormat.Items.AddRange(new string[] { "Text", "ZPL", "Code128", "QRCode" });
            this.cmbTemplateFormat.SelectedIndex = 0;
            
            this.btnSaveTemplate.Click += btnSaveTemplate_Click;
            this.btnPreviewTemplate.Click += btnPreviewTemplate_Click;
            
            this.grpTemplateEditor.Controls.AddRange(new Control[] { lblTemplateName, txtTemplateName, lblTemplateFormat, cmbTemplateFormat, lblTemplateContent, txtTemplateContent, btnSaveTemplate, btnPreviewTemplate });
            
            // 可用字段组
            this.grpTemplatePreview = new GroupBox { Text = LanguageManager.GetString("TemplatePreview"), Location = new Point(880, 10), Size = new Size(300, 350) };
            this.lblAvailableFields = new Label { Text = LanguageManager.GetString("AvailableFields"), Location = new Point(15, 30), AutoSize = true };
            this.lstAvailableFields = new ListBox { Location = new Point(15, 50), Size = new Size(270, 120) };
            this.rtbTemplatePreview = new RichTextBox { Location = new Point(15, 190), Size = new Size(270, 150), ReadOnly = true };
            
            // 填充可用字段
            var fields = PrintTemplateManager.GetAvailableFields();
            foreach (var field in fields)
            {
                this.lstAvailableFields.Items.Add(field);
            }
            
            this.lstAvailableFields.DoubleClick += lstAvailableFields_DoubleClick;
            
            this.grpTemplatePreview.Controls.AddRange(new Control[] { lblAvailableFields, lstAvailableFields, rtbTemplatePreview });
            
            this.tabTemplate.Controls.AddRange(new Control[] { grpTemplateList, grpTemplateEditor, grpTemplatePreview });
        }

        private void InitializeLogsTab()
        {
            this.tabLogs.Text = "运行日志";
            this.tabLogs.UseVisualStyleBackColor = true;
            
            this.txtLog = new TextBox 
            { 
                Location = new Point(10, 50), 
                Size = new Size(1160, 670), 
                Multiline = true, 
                ScrollBars = ScrollBars.Vertical, 
                ReadOnly = true, 
                Font = new Font("Consolas", 9F)
            };
            
            this.btnClearLog = new Button { Text = LanguageManager.GetString("ClearLog"), Location = new Point(10, 15), Size = new Size(100, 30) };
            this.btnSaveLog = new Button { Text = LanguageManager.GetString("SaveLog"), Location = new Point(120, 15), Size = new Size(100, 30) };
            
            this.btnClearLog.Click += btnClearLog_Click;
            this.btnSaveLog.Click += btnSaveLog_Click;
            
            this.tabLogs.Controls.AddRange(new Control[] { txtLog, btnClearLog, btnSaveLog });
        }

        // 事件处理方法声明
        private void btnBrowseDatabase_Click(object? sender, EventArgs e)
        {
            using var openFileDialog = new OpenFileDialog
            {
                Title = "选择Access数据库文件",
                Filter = "Access数据库文件 (*.mdb;*.accdb)|*.mdb;*.accdb|所有文件 (*.*)|*.*",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtDatabasePath.Text = openFileDialog.FileName;
                
                // 更新配置
                var config = ConfigurationManager.Config;
                config.Database.DatabasePath = openFileDialog.FileName;
                ConfigurationManager.SaveConfig();
                
                AddLogMessage($"数据库路径已更新: {openFileDialog.FileName}");
            }
        }

        private void btnTestConnection_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtDatabasePath.Text))
            {
                MessageBox.Show("请先选择数据库文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var config = ConfigurationManager.Config.Database;
                var success = _databaseMonitor.Connect(txtDatabasePath.Text, config.TableName, config.MonitorField);
                
                if (success)
                {
                    MessageBox.Show("数据库连接测试成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    AddLogMessage("数据库连接测试成功");
                    
                    // 加载最近记录
                    LoadRecentRecords();
                }
                else
                {
                    MessageBox.Show("数据库连接失败，请检查文件路径和格式", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"数据库连接测试失败: {ex.Message}", ex);
                MessageBox.Show($"数据库连接测试失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStartMonitoring_Click(object? sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtDatabasePath.Text))
                {
                    MessageBox.Show("请先选择数据库文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var config = ConfigurationManager.Config.Database;
                
                // 先连接数据库
                if (!_databaseMonitor.Connect(txtDatabasePath.Text, config.TableName, config.MonitorField))
                {
                    MessageBox.Show("数据库连接失败，无法开始监控", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 开始监控
                _databaseMonitor.StartMonitoring(config.PollInterval);
                AddLogMessage("数据监控已启动");
                
                MessageBox.Show("数据监控已启动", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"启动监控失败: {ex.Message}", ex);
                MessageBox.Show($"启动监控失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStopMonitoring_Click(object? sender, EventArgs e)
        {
            try
            {
                _databaseMonitor.StopMonitoring();
                AddLogMessage("数据监控已停止");
                MessageBox.Show("数据监控已停止", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"停止监控失败: {ex.Message}", ex);
                MessageBox.Show($"停止监控失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnManualPrint_Click(object? sender, EventArgs e)
        {
            if (lvRecords.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先选择要打印的记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var selectedItem = lvRecords.SelectedItems[0];
                var record = selectedItem.Tag as TestRecord;
                
                if (record == null)
                {
                    MessageBox.Show("无效的记录数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 检查是否为重复打印
                if (!ConfirmPrintIfAlreadyPrinted(record))
                {
                    return; // 用户取消打印
                }

                var config = ConfigurationManager.Config.Printer;
                // 使用默认模板
                var templateName = config.DefaultTemplate;
                var printResult = _printerService.PrintRecord(record, config.PrintFormat, templateName);

                if (printResult.Success)
                {
                    _totalPrintJobs++;
                    _databaseMonitor.UpdatePrintCount(record.TR_SerialNum ?? record.TR_ID ?? "");
                    
                    AddLogMessage($"手动打印完成: {record.TR_SerialNum}");
                    MessageBox.Show("打印任务已发送", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    UpdateStatusDisplay();
                    btnRefresh_Click(null, EventArgs.Empty); // 刷新列表显示最新状态
                }
                else
                {
                    AddLogMessage($"手动打印失败: {printResult.ErrorMessage}");
                    MessageBox.Show($"打印失败:\n{printResult.ErrorMessage}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                    // 如果是因为没有打印机导致的失败，显示安装提示
                    if (printResult.ErrorMessage?.Contains("打印机") == true && !_printerService.HasAnyPrinter())
                    {
                        var title = LanguageManager.GetString("NoPrinterTitle");
                        var message = _printerService.GetNoPrinterMessage();
                        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"手动打印失败: {ex.Message}", ex);
                MessageBox.Show($"打印过程中发生错误:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRefresh_Click(object? sender, EventArgs e)
        {
            LoadRecentRecords();
        }

        private void cmbPrinter_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbPrinter.SelectedItem != null)
            {
                var printerName = cmbPrinter.SelectedItem.ToString();
                _printerService.UpdatePrinterName(printerName!);
                
                // 更新配置
                var config = ConfigurationManager.Config;
                config.Printer.PrinterName = printerName!;
                ConfigurationManager.SaveConfig();
                
                AddLogMessage($"打印机已更改为: {printerName}");
                lblPrinterStatus.Text = $"当前打印机: {printerName}";
                lblPrinterStatus.ForeColor = Color.Green;
            }
        }

        private void cmbPrintFormat_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbPrintFormat.SelectedItem != null)
            {
                var format = cmbPrintFormat.SelectedItem.ToString();
                
                // 更新配置
                var config = ConfigurationManager.Config;
                config.Printer.PrintFormat = format!;
                ConfigurationManager.SaveConfig();
                
                AddLogMessage($"打印格式已更改为: {format}");
            }
        }

        private void btnTestPrint_Click(object? sender, EventArgs e)
        {
            try
            {
                var result = _printerService.TestPrint();
                
                if (result.Success)
                {
                    AddLogMessage($"测试打印成功: {result.PrinterUsed}");
                    MessageBox.Show($"测试打印已发送到: {result.PrinterUsed}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    AddLogMessage($"测试打印失败: {result.ErrorMessage}");
                    MessageBox.Show($"测试打印失败:\n{result.ErrorMessage}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                    // 如果是因为没有打印机导致的失败，显示安装提示
                    if (result.ErrorMessage?.Contains("打印机") == true && !_printerService.HasAnyPrinter())
                    {
                        var title = LanguageManager.GetString("NoPrinterTitle");
                        var message = _printerService.GetNoPrinterMessage();
                        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"测试打印失败: {ex.Message}", ex);
                MessageBox.Show($"测试打印过程中发生错误:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void numPollInterval_ValueChanged(object? sender, EventArgs e)
        {
            var newInterval = (int)numPollInterval.Value;
            
            // 更新配置
            var config = ConfigurationManager.Config;
            config.Database.PollInterval = newInterval;
            ConfigurationManager.SaveConfig();
            
            AddLogMessage($"轮询间隔已更改为: {newInterval}ms");
        }

        private void btnClearLog_Click(object? sender, EventArgs e)
        {
            txtLog.Clear();
            AddLogMessage("日志已清空");
        }

        private void btnSaveLog_Click(object? sender, EventArgs e)
        {
            try
            {
                using var saveFileDialog = new SaveFileDialog
                {
                    Title = "保存日志文件",
                    Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                    DefaultExt = "txt",
                    FileName = $"监控日志_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, txtLog.Text);
                    MessageBox.Show("日志保存成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    AddLogMessage($"日志已保存到: {saveFileDialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"保存日志失败: {ex.Message}", ex);
                MessageBox.Show($"保存日志失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkEnablePrintCount_CheckedChanged(object? sender, EventArgs e)
        {
            // 更新配置
            var config = ConfigurationManager.Config;
            config.Database.EnablePrintCount = chkEnablePrintCount.Checked;
            ConfigurationManager.SaveConfig();
            
            var status = chkEnablePrintCount.Checked ? "启用" : "禁用";
            AddLogMessage($"打印次数统计已{status}");
            
            if (chkEnablePrintCount.Checked)
            {
                MessageBox.Show("打印次数统计已启用。\n新的打印操作将更新数据库中的TR_Print字段。", 
                    "功能启用", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("打印次数统计已禁用。\n打印操作将不会更新数据库中的TR_Print字段，保持数据库兼容性。", 
                    "功能禁用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnShowPrintMonitor_Click(object? sender, EventArgs e)
        {
            ShowPrintMonitor();
        }

        private void LvRecords_DoubleClick(object? sender, EventArgs e)
        {
            if (lvRecords.SelectedItems.Count > 0)
            {
                var selectedItem = lvRecords.SelectedItems[0];
                var record = selectedItem.Tag as TestRecord;
                if (record != null)
                {
                    PrintSelectedRecord(record);
                }
            }
        }

        private void LvRecords_MouseClick(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && lvRecords.SelectedItems.Count > 0)
            {
                var contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("打印此记录", null, (s, args) => {
                    var selectedItem = lvRecords.SelectedItems[0];
                    var record = selectedItem.Tag as TestRecord;
                    if (record != null)
                    {
                        PrintSelectedRecord(record);
                    }
                });
                contextMenu.Items.Add("查看详细信息", null, (s, args) => {
                    var selectedItem = lvRecords.SelectedItems[0];
                    var record = selectedItem.Tag as TestRecord;
                    if (record != null)
                    {
                        ShowRecordDetails(record);
                    }
                });
                contextMenu.Show(lvRecords, e.Location);
            }
        }

        private void PrintSelectedRecord(TestRecord record)
        {
            try
            {
                // 检查是否为重复打印
                if (!ConfirmPrintIfAlreadyPrinted(record))
                {
                    return; // 用户取消打印
                }

                var config = ConfigurationManager.Config;
                // 使用默认模板或配置的模板
                var templateName = config.Printer.DefaultTemplate;
                var printResult = _printerService.PrintRecord(record, config.Printer.PrintFormat, templateName);
                
                if (printResult.Success)
                {
                    _databaseMonitor.UpdatePrintCount(record.TR_SerialNum ?? "");
                    AddLogMessage($"成功打印记录: {record.TR_SerialNum}");
                    btnRefresh_Click(null, EventArgs.Empty); // 刷新列表
                }
                else
                {
                    AddLogMessage($"打印失败: {printResult.ErrorMessage}");
                    
                    // 如果是因为没有打印机导致的失败，显示安装提示
                    if (printResult.ErrorMessage?.Contains("打印机") == true && !_printerService.HasAnyPrinter())
                    {
                        var title = LanguageManager.GetString("NoPrinterTitle");
                        var message = _printerService.GetNoPrinterMessage();
                        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"打印记录失败: {ex.Message}", ex);
                AddLogMessage($"打印记录失败: {ex.Message}");
            }
        }

        private bool ConfirmPrintIfAlreadyPrinted(TestRecord record)
        {
            // 检查打印次数
            if (record.TR_Print > 0)
            {
                var title = LanguageManager.GetString("ReprintWarningTitle");
                var message = string.Format(
                    LanguageManager.GetString("ReprintWarningMessage"), 
                    record.TR_Print);
                
                var result = MessageBox.Show(
                    message, 
                    title, 
                    MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Question);
                
                return result == DialogResult.Yes;
            }
            
            return true; // 如果从未打印过，直接允许打印
        }

        private void ShowRecordDetails(TestRecord record)
        {
            var details = $"序列号: {record.TR_SerialNum ?? "N/A"}\n" +
                         $"测试时间: {record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A"}\n" +
                         $"短路电流 (Isc): {record.FormatNumber(record.TR_Isc)} A\n" +
                         $"开路电压 (Voc): {record.FormatNumber(record.TR_Voc)} V\n" +
                         $"最大功率点电压 (Vpm): {record.FormatNumber(record.TR_Vpm)} V\n" +
                         $"最大功率点电流 (Ipm): {record.FormatNumber(record.TR_Ipm)} A\n" +
                         $"最大功率 (Pm): {record.FormatNumber(record.TR_Pm)} W\n" +
                         $"效率: {record.FormatNumber(record.TR_CellEfficiency)}%\n" +
                         $"填充因子 (FF): {record.FormatNumber(record.TR_FF)}\n" +
                         $"等级: {record.TR_Grade ?? "N/A"}\n" +
                         $"打印次数: {record.TR_Print ?? 0}";
            
            MessageBox.Show(details, "记录详细信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
} 