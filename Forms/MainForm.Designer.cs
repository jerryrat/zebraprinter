#nullable enable
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
        private Button btnManualPrint, btnRefresh, btnTestPrint, btnClearLog, btnSaveLog, btnPrintPreview;
        private CheckBox chkAutoPrint, chkAutoStartMonitoring, chkMinimizeToTray;
        private CheckBox chkEnablePrintCount;  // 新增打印次数控制复选框
        private ComboBox cmbPrinter, cmbPrintFormat, cmbLanguage;
        private NumericUpDown numPollInterval;
        private ListView lvRecords;
        private Label lblMonitoringStatus, lblTotalRecords, lblTotalPrints, lblLastRecord, lblPrinterStatus;
        
        // 打印模板页面控件
        private GroupBox grpTemplateEditor, grpTemplateList, grpTemplatePreview;
        private Label lblTemplateName, lblTemplateContent, lblTemplateFormat, lblAvailableFields;
        private TextBox txtTemplateName, txtTemplateContent;
        private ComboBox cmbTemplateFormat, cmbTemplateList;
        private Button btnSaveTemplate, btnDeleteTemplate, btnPreviewTemplate, btnNewTemplate, btnVisualDesigner;
        private ListBox lstAvailableFields;
        private RichTextBox rtbTemplatePreview;

        // 预印刷标签模式相关控件已删除


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
            
            // 主窗体设置 - 扩大尺寸避免遮挡
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1220, 800);  // 增加宽度以适应新布局
            this.Controls.Add(this.tabControl1);
            this.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimumSize = new Size(1220, 700);  // 增加最小宽度
            this.Name = "MainForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = $"{LanguageManager.GetString("MainTitle")} v1.1.44";
            
            // TabControl设置
            this.tabControl1.Controls.Add(this.tabMonitor);
            this.tabControl1.Controls.Add(this.tabConfig);
            this.tabControl1.Controls.Add(this.tabTemplate);
            this.tabControl1.Controls.Add(this.tabLogs);
            this.tabControl1.Dock = DockStyle.Fill;
            this.tabControl1.Location = new Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new Size(1200, 800);
            this.tabControl1.TabIndex = 0;
            
            InitializeMonitorTab();
            InitializeConfigTab();
            InitializeTemplateTab();
            InitializeLogsTab();
            
            this.StartPosition = FormStartPosition.CenterScreen;
            
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private void InitializeMonitorTab()
        {
            this.tabMonitor.Text = LanguageManager.GetString("TabDataMonitoring");
            this.tabMonitor.UseVisualStyleBackColor = true;
            
            // 数据库配置组 - 调整大小和位置
            this.grpDatabaseConfig = new GroupBox { Text = LanguageManager.GetString("DatabaseConfig"), Location = new Point(10, 10), Size = new Size(1160, 80) };
            this.lblDatabasePath = new Label { Text = LanguageManager.GetString("DatabasePath"), Location = new Point(15, 25), AutoSize = true };
            this.txtDatabasePath = new TextBox { Location = new Point(15, 46), Size = new Size(850, 25), ReadOnly = true };
            this.btnBrowseDatabase = new Button { Text = LanguageManager.GetString("Browse"), Location = new Point(875, 43), Size = new Size(80, 27) };
            this.btnTestConnection = new Button { Text = LanguageManager.GetString("TestConnection"), Location = new Point(965, 43), Size = new Size(100, 27) };
            
            this.btnBrowseDatabase.Click += btnBrowseDatabase_Click;
            this.btnTestConnection.Click += btnTestConnection_Click;
            
            this.grpDatabaseConfig.Controls.AddRange(new Control[] { lblDatabasePath, txtDatabasePath, btnBrowseDatabase, btnTestConnection });
            
            // 监控控制组 - 调整大小
            this.grpMonitorControl = new GroupBox { Text = LanguageManager.GetString("MonitorControl"), Location = new Point(10, 100), Size = new Size(1160, 65) };
            this.btnStartMonitoring = new Button { Text = LanguageManager.GetString("StartMonitoring"), Location = new Point(15, 25), Size = new Size(100, 30) };
            this.btnStopMonitoring = new Button { Text = LanguageManager.GetString("StopMonitoring"), Location = new Point(125, 25), Size = new Size(100, 30) };
            this.btnPrintPreview = new Button { Text = LanguageManager.GetString("PrintPreview"), Location = new Point(235, 25), Size = new Size(100, 30), BackColor = SystemColors.Control, ForeColor = SystemColors.ControlText };
            this.chkAutoPrint = new CheckBox { Text = LanguageManager.GetString("AutoPrint"), Location = new Point(350, 30), Checked = true, AutoSize = true };
            this.chkEnablePrintCount = new CheckBox { Text = LanguageManager.GetString("EnablePrintCount"), Location = new Point(470, 30), Checked = false, AutoSize = true };
            
            this.btnStartMonitoring.Click += btnStartMonitoring_Click;
            this.btnStopMonitoring.Click += btnStopMonitoring_Click;
            this.btnPrintPreview.Click += btnPrintPreview_Click;
            this.chkEnablePrintCount.CheckedChanged += chkEnablePrintCount_CheckedChanged;
            
            this.grpMonitorControl.Controls.AddRange(new Control[] { btnStartMonitoring, btnStopMonitoring, btnPrintPreview, chkAutoPrint, chkEnablePrintCount });
            
            // 状态信息组 - 调整大小和布局
            this.grpStatus = new GroupBox { Text = LanguageManager.GetString("StatusInfo"), Location = new Point(10, 175), Size = new Size(1160, 85) };
            this.lblMonitoringStatus = new Label { Text = LanguageManager.GetString("MonitoringStatusStopped"), Location = new Point(15, 25), ForeColor = Color.Red, Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold), AutoSize = true };
            this.lblTotalRecords = new Label { Text = LanguageManager.GetString("TotalRecords"), Location = new Point(15, 50), AutoSize = true };
            this.lblTotalPrints = new Label { Text = LanguageManager.GetString("TotalPrints"), Location = new Point(200, 50), AutoSize = true };
            this.lblLastRecord = new Label { Text = LanguageManager.GetString("LastRecord"), Location = new Point(400, 50), AutoSize = true };
            
            this.grpStatus.Controls.AddRange(new Control[] { lblMonitoringStatus, lblTotalRecords, lblTotalPrints, lblLastRecord });
            
            // 记录列表组 - 大幅扩大尺寸并改进ListView
            this.grpRecords = new GroupBox { Text = LanguageManager.GetString("RecentRecords"), Location = new Point(10, 270), Size = new Size(1160, 480) };
            this.lvRecords = new ListView 
            { 
                Location = new Point(15, 30), 
                Size = new Size(1130, 400), 
                View = View.Details, 
                FullRowSelect = true, 
                GridLines = true,
                MultiSelect = false,
                HideSelection = false
            };
            
            // 重新定义列，添加Vpm和打印次数列
            this.lvRecords.Columns.Add(LanguageManager.GetString("SerialNumber"), 150);
            this.lvRecords.Columns.Add(LanguageManager.GetString("TestDateTime"), 150);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Current"), 120);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Voltage"), 120);
            this.lvRecords.Columns.Add(LanguageManager.GetString("VoltageVpm"), 120);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Power"), 120);
            this.lvRecords.Columns.Add(LanguageManager.GetString("PrintCount"), 100);
            this.lvRecords.Columns.Add(LanguageManager.GetString("Operation"), 250);
            
            // 添加双击事件用于打印
            this.lvRecords.DoubleClick += LvRecords_DoubleClick;
            this.lvRecords.MouseClick += LvRecords_MouseClick;
            
            this.btnManualPrint = new Button { Text = LanguageManager.GetString("PrintSelected"), Location = new Point(15, 445), Size = new Size(100, 30) };
            this.btnRefresh = new Button { Text = LanguageManager.GetString("Refresh"), Location = new Point(125, 445), Size = new Size(80, 30) };
            
            this.btnManualPrint.Click += btnManualPrint_Click;
            this.btnRefresh.Click += btnRefresh_Click;
            
            this.grpRecords.Controls.AddRange(new Control[] { lvRecords, btnManualPrint, btnRefresh });
            
            this.tabMonitor.Controls.AddRange(new Control[] { grpDatabaseConfig, grpMonitorControl, grpStatus, grpRecords });
        }

        private void InitializeConfigTab()
        {
            this.tabConfig.Text = LanguageManager.GetString("TabSystemConfig");
            this.tabConfig.UseVisualStyleBackColor = true;
            
            // 打印机配置组 - 调整尺寸
            this.grpPrinterConfig = new GroupBox { Text = LanguageManager.GetString("PrinterConfig"), Location = new Point(10, 10), Size = new Size(1160, 120) };
            this.lblPrinter = new Label { Text = LanguageManager.GetString("SelectedPrinter"), Location = new Point(15, 30), AutoSize = true };
            this.cmbPrinter = new ComboBox { Location = new Point(15, 50), Size = new Size(400, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.lblPrintFormat = new Label { Text = LanguageManager.GetString("PrintFormat"), Location = new Point(450, 30), AutoSize = true };
            this.cmbPrintFormat = new ComboBox { Location = new Point(450, 50), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.btnTestPrint = new Button { Text = LanguageManager.GetString("TestPrint"), Location = new Point(680, 50), Size = new Size(100, 30) };
            this.lblPrinterStatus = new Label { Text = LanguageManager.GetString("PrinterStatus"), Location = new Point(15, 85), AutoSize = true };
            
            this.cmbPrintFormat.Items.AddRange(new string[] { "Text", "ZPL", "Code128", "QRCode" });
            // SelectedIndex 将在 LoadConfiguration 中根据配置设置
            
            this.cmbPrinter.SelectedIndexChanged += cmbPrinter_SelectedIndexChanged;

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
            this.tabTemplate.Text = LanguageManager.GetString("TabPrintTemplate");
            this.tabTemplate.UseVisualStyleBackColor = true;
            
            // 模板列表组
            this.grpTemplateList = new GroupBox { Text = LanguageManager.GetString("TemplateList"), Location = new Point(10, 10), Size = new Size(280, 400) };
            this.cmbTemplateList = new ComboBox { Location = new Point(15, 30), Size = new Size(250, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.btnNewTemplate = new Button { Text = LanguageManager.GetString("NewTemplate"), Location = new Point(15, 70), Size = new Size(80, 30) };
            this.btnDeleteTemplate = new Button { Text = LanguageManager.GetString("DeleteTemplate"), Location = new Point(110, 70), Size = new Size(80, 30) };
            this.btnVisualDesigner = new Button { Text = LanguageManager.GetString("VisualDesigner"), Location = new Point(15, 110), Size = new Size(180, 30) };
            
            this.cmbTemplateList.SelectedIndexChanged += cmbTemplateList_SelectedIndexChanged;
            this.btnNewTemplate.Click += btnNewTemplate_Click;
            this.btnDeleteTemplate.Click += btnDeleteTemplate_Click;
            this.btnVisualDesigner.Click += btnVisualDesigner_Click;
            
            this.grpTemplateList.Controls.AddRange(new Control[] { cmbTemplateList, btnNewTemplate, btnDeleteTemplate, btnVisualDesigner });
            
            // 模板编辑器组 - 增加高度为新控件留出空间
            this.grpTemplateEditor = new GroupBox { Text = LanguageManager.GetString("TemplateEditor"), Location = new Point(300, 10), Size = new Size(580, 520) };
            this.lblTemplateName = new Label { Text = LanguageManager.GetString("TemplateName"), Location = new Point(15, 30), AutoSize = true };
            this.txtTemplateName = new TextBox { Location = new Point(15, 50), Size = new Size(250, 25) };
            this.lblTemplateFormat = new Label { Text = LanguageManager.GetString("PrintFormat"), Location = new Point(350, 30), AutoSize = true };
            this.cmbTemplateFormat = new ComboBox { Location = new Point(350, 50), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            
            this.lblTemplateContent = new Label { Text = LanguageManager.GetString("TemplateContent"), Location = new Point(15, 80), AutoSize = true };
            
            // 调整模板内容编辑框的位置和大小
            this.txtTemplateContent = new TextBox { 
                Location = new Point(15, 100), 
                Size = new Size(550, 280), 
                Multiline = true, 
                ScrollBars = ScrollBars.Both,
                Font = new Font("Consolas", 9F)
            };

            

            // 调整按钮位置
            this.btnSaveTemplate = new Button { 
                Text = LanguageManager.GetString("SaveTemplate"), 
                Location = new Point(15, 390), 
                Size = new Size(100, 40), 
                Visible = true,
                BackColor = Color.FromArgb(0, 123, 255), // 蓝色背景
                ForeColor = Color.White, // 白色文字
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
            this.btnSaveTemplate.FlatAppearance.BorderSize = 0;
            
            // 添加工具提示
            var toolTip = new ToolTip();
            toolTip.SetToolTip(this.btnSaveTemplate, LanguageManager.GetString("SaveTemplateToolTip"));
            
            this.btnPreviewTemplate = new Button { 
                Text = LanguageManager.GetString("PreviewTemplate"), 
                Location = new Point(125, 390), 
                Size = new Size(100, 40), 
                Visible = true,
                BackColor = Color.FromArgb(40, 167, 69), // 绿色背景
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
            this.btnPreviewTemplate.FlatAppearance.BorderSize = 0;
            
            var btnClearTemplate = new Button { 
                Text = LanguageManager.GetString("ClearContent"), 
                Location = new Point(235, 390), 
                Size = new Size(100, 40), 
                Visible = true,
                BackColor = Color.FromArgb(255, 193, 7), // 黄色背景
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
            btnClearTemplate.FlatAppearance.BorderSize = 0;
            
            var btnImportTemplate = new Button { 
                Text = LanguageManager.GetString("ImportTemplate"), 
                Location = new Point(345, 390), 
                Size = new Size(100, 40), 
                Visible = true,
                BackColor = Color.FromArgb(108, 117, 125), // 灰色背景
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
            btnImportTemplate.FlatAppearance.BorderSize = 0;
            
            // 添加分隔线说明
            var lblButtonsInfo = new Label { 
                Text = LanguageManager.GetString("ButtonsInfo"), 
                Location = new Point(15, 440), 
                Size = new Size(400, 20), 
                ForeColor = Color.Gray, 
                Font = new Font("Microsoft YaHei", 8F)
            };
            
            this.cmbTemplateFormat.Items.AddRange(new string[] { "Text", "ZPL", "Code128", "QRCode" });
            this.cmbTemplateFormat.SelectedIndex = 0;
            
            this.btnSaveTemplate.Click += btnSaveTemplate_Click;
            this.btnPreviewTemplate.Click += btnPreviewTemplate_Click;
            btnClearTemplate.Click += btnClearTemplate_Click;
            btnImportTemplate.Click += btnImportTemplate_Click;
            
            // 确保所有控件都被添加到组中
            this.grpTemplateEditor.Controls.AddRange(new Control[] { 
                this.lblTemplateName, this.txtTemplateName, 
                this.lblTemplateFormat, this.cmbTemplateFormat, 
                this.lblTemplateContent, this.txtTemplateContent, 
                this.btnSaveTemplate, this.btnPreviewTemplate, 
                btnClearTemplate, btnImportTemplate,
                lblButtonsInfo
            });
            
            // 可用字段组 - 调整位置和大小
            this.grpTemplatePreview = new GroupBox { Text = LanguageManager.GetString("TemplatePreview"), Location = new Point(890, 10), Size = new Size(300, 520) };
            this.lblAvailableFields = new Label { Text = LanguageManager.GetString("AvailableFields"), Location = new Point(15, 30), AutoSize = true };
            this.lstAvailableFields = new ListBox { Location = new Point(15, 50), Size = new Size(270, 130) };
            var lblPreview = new Label { Text = LanguageManager.GetString("PreviewLabel"), Location = new Point(15, 210), AutoSize = true };
            this.rtbTemplatePreview = new RichTextBox { Location = new Point(15, 230), Size = new Size(270, 275), ReadOnly = true, WordWrap = true };
            
            // 填充可用字段
            var fields = PrintTemplateManager.GetAvailableFields();
            foreach (var field in fields)
            {
                this.lstAvailableFields.Items.Add(field);
            }
            
            this.lstAvailableFields.DoubleClick += lstAvailableFields_DoubleClick;
            
            this.grpTemplatePreview.Controls.AddRange(new Control[] { this.lblAvailableFields, this.lstAvailableFields, lblPreview, this.rtbTemplatePreview });
            
            this.tabTemplate.Controls.AddRange(new Control[] { this.grpTemplateList, this.grpTemplateEditor, this.grpTemplatePreview });
        }

        private void InitializeLogsTab()
        {
            this.tabLogs.Text = LanguageManager.GetString("TabRuntimeLogs");
            this.tabLogs.UseVisualStyleBackColor = true;
            
            // 日志控件
            this.txtLog = new TextBox { 
                Location = new Point(10, 55), 
                Size = new Size(1165, 690), 
                Multiline = true, 
                ReadOnly = true, 
                WordWrap = true,
                ScrollBars = ScrollBars.Vertical
            };
            
            this.btnClearLog = new Button { Text = LanguageManager.GetString("ClearLogs"), Location = new Point(10, 15), Size = new Size(100, 30) };
            this.btnSaveLog = new Button { Text = LanguageManager.GetString("SaveLogs"), Location = new Point(120, 15), Size = new Size(100, 30) };
            
            this.btnClearLog.Click += btnClearLog_Click;
            this.btnSaveLog.Click += btnSaveLog_Click;
            
            this.tabLogs.Controls.AddRange(new Control[] { this.txtLog, this.btnClearLog, this.btnSaveLog });
        }

        // 事件处理方法声明
        private void btnBrowseDatabase_Click(object? sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = LanguageManager.GetString("DatabaseFileFilter");
                openFileDialog.Title = LanguageManager.GetString("SelectDatabaseDialogTitle");
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.txtDatabasePath.Text = openFileDialog.FileName;
                    
                    // 保存到配置
                    var config = ConfigurationManager.Config;
                    config.Database.DatabasePath = openFileDialog.FileName;
                    ConfigurationManager.SaveConfig();
                }
            }
        }

        private void btnTestConnection_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtDatabasePath.Text))
            {
                MessageBox.Show(LanguageManager.GetString("DatabasePath"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 测试数据库连接
                if (File.Exists(this.txtDatabasePath.Text))
                {
                    MessageBox.Show(LanguageManager.GetString("Success"), LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(LanguageManager.GetString("DatabaseConnectionFailed"), LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"数据库连接测试失败: {ex.Message}", ex);
                MessageBox.Show($"{LanguageManager.GetString("Error")}: {ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStartMonitoring_Click(object? sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txtDatabasePath.Text))
                {
                    MessageBox.Show(LanguageManager.GetString("DatabasePath"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var config = ConfigurationManager.Config.Database;
                
                // 先连接数据库
                if (!_databaseMonitor.Connect(this.txtDatabasePath.Text, config.TableName, config.MonitorField))
                {
                    MessageBox.Show(LanguageManager.GetString("DatabaseConnectionFailed"), LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 开始监控
                _databaseMonitor.StartMonitoring(config.PollInterval);
                AddLogMessage("数据监控已启动");
                
                MessageBox.Show(LanguageManager.GetString("MonitoringStarted"), LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"启动监控失败: {ex.Message}", ex);
                MessageBox.Show($"{LanguageManager.GetString("Error")}: {ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStopMonitoring_Click(object? sender, EventArgs e)
        {
            try
            {
                _databaseMonitor.StopMonitoring();
                AddLogMessage("数据监控已停止");
                MessageBox.Show(LanguageManager.GetString("MonitoringStopped"), LanguageManager.GetString("Information"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"停止监控失败: {ex.Message}", ex);
                MessageBox.Show($"{LanguageManager.GetString("Error")}: {ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnManualPrint_Click(object? sender, EventArgs e)
        {
            if (this.lvRecords.SelectedItems.Count == 0)
            {
                MessageBox.Show(LanguageManager.GetString("SelectDatabase"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var selectedItem = this.lvRecords.SelectedItems[0];
                var record = selectedItem.Tag as TestRecord;
                
                if (record == null)
                {
                    MessageBox.Show(LanguageManager.GetString("InvalidRecordData"), LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show(LanguageManager.GetString("PrintTaskSent"), LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    UpdateStatusDisplay();
                    btnRefresh_Click(null, EventArgs.Empty); // 刷新列表显示最新状态
                }
                else
                {
                    AddLogMessage($"手动打印失败: {printResult.ErrorMessage}");
                    MessageBox.Show($"{LanguageManager.GetString("PrintFailed")}:\n{printResult.ErrorMessage}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
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
                MessageBox.Show($"{LanguageManager.GetString("PrintError")}:\n{ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRefresh_Click(object? sender, EventArgs e)
        {
            LoadRecentRecords();
        }

        private void cmbPrinter_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (this.cmbPrinter.SelectedItem != null)
            {
                var printerName = this.cmbPrinter.SelectedItem.ToString();
                _printerService.UpdatePrinterName(printerName!);
                
                // 更新配置
                var config = ConfigurationManager.Config;
                config.Printer.PrinterName = printerName!;
                ConfigurationManager.SaveConfig();
                
                AddLogMessage($"打印机已更改为: {printerName}");
                this.lblPrinterStatus.Text = $"{LanguageManager.GetString("CurrentPrinter")}: {printerName}";
                this.lblPrinterStatus.ForeColor = Color.Green;
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
                    MessageBox.Show($"{LanguageManager.GetString("TestPrintSent")}: {result.PrinterUsed}", LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    AddLogMessage($"测试打印失败: {result.ErrorMessage}");
                    MessageBox.Show($"{LanguageManager.GetString("TestPrintFailed")}:\n{result.ErrorMessage}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
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
                MessageBox.Show($"{LanguageManager.GetString("TestPrintError")}:\n{ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void numPollInterval_ValueChanged(object? sender, EventArgs e)
        {
            var newInterval = (int)this.numPollInterval.Value;
            
            // 更新配置
            var config = ConfigurationManager.Config;
            config.Database.PollInterval = newInterval;
            ConfigurationManager.SaveConfig();
            
            AddLogMessage($"{LanguageManager.GetString("PollIntervalChanged")}: {newInterval}ms");
        }

        private void btnClearLog_Click(object? sender, EventArgs e)
        {
            this.txtLog.Clear();
            AddLogMessage("日志已清空");
        }

        private void btnSaveLog_Click(object? sender, EventArgs e)
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = LanguageManager.GetString("LogFileFilter");
                saveFileDialog.DefaultExt = "txt";
                saveFileDialog.Title = LanguageManager.GetString("SaveLogDialogTitle");
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        File.WriteAllText(saveFileDialog.FileName, this.txtLog.Text);
                        MessageBox.Show(LanguageManager.GetString("Success"), LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                        AddLogMessage($"{LanguageManager.GetString("LogSaved")}: {saveFileDialog.FileName}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"保存日志失败: {ex.Message}", ex);
                        MessageBox.Show($"{LanguageManager.GetString("Error")}: {ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void chkEnablePrintCount_CheckedChanged(object? sender, EventArgs e)
        {
            // 更新配置
            var config = ConfigurationManager.Config;
            config.Database.EnablePrintCount = this.chkEnablePrintCount.Checked;
            ConfigurationManager.SaveConfig();
            
            var status = this.chkEnablePrintCount.Checked ? LanguageManager.GetString("Enabled") : LanguageManager.GetString("Disabled");
            AddLogMessage($"{LanguageManager.GetString("PrintCountStatistic")} {status}");
            
            if (this.chkEnablePrintCount.Checked)
            {
                MessageBox.Show($"{LanguageManager.GetString("PrintCountEnabledMessage")}\n{LanguageManager.GetString("PrintOperationWillUpdate")}", 
                    LanguageManager.GetString("FunctionEnabled"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"{LanguageManager.GetString("PrintCountDisabledMessage")}\n{LanguageManager.GetString("PrintOperationWillNotUpdate")}", 
                    LanguageManager.GetString("FunctionDisabled"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LvRecords_DoubleClick(object? sender, EventArgs e)
        {
            if (this.lvRecords.SelectedItems.Count > 0)
            {
                var selectedItem = this.lvRecords.SelectedItems[0];
                var record = selectedItem.Tag as TestRecord;
                if (record != null)
                {
                    PrintSelectedRecord(record);
                }
            }
        }

        private void LvRecords_MouseClick(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && this.lvRecords.SelectedItems.Count > 0)
            {
                var contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add(LanguageManager.GetString("PrintThisRecord"), null, (s, args) => {
                    var selectedItem = this.lvRecords.SelectedItems[0];
                    var record = selectedItem.Tag as TestRecord;
                    if (record != null)
                    {
                        PrintSelectedRecord(record);
                    }
                });
                contextMenu.Items.Add(LanguageManager.GetString("ViewDetails"), null, (s, args) => {
                    var selectedItem = this.lvRecords.SelectedItems[0];
                    var record = selectedItem.Tag as TestRecord;
                    if (record != null)
                    {
                        ShowRecordDetails(record);
                    }
                });
                contextMenu.Show(this.lvRecords, e.Location);
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

                var config = ConfigurationManager.Config.Printer;
                // 使用默认模板或配置的模板
                var templateName = config.DefaultTemplate;
                var printResult = _printerService.PrintRecord(record, config.PrintFormat, templateName);
                
                if (printResult.Success)
                {
                    _databaseMonitor.UpdatePrintCount(record.TR_SerialNum ?? "");
                    AddLogMessage($"{LanguageManager.GetString("PrintRecordSuccess")}: {record.TR_SerialNum}");
                    btnRefresh_Click(null, EventArgs.Empty); // 刷新列表
                }
                else
                {
                    AddLogMessage($"{LanguageManager.GetString("PrintRecordFailed")}: {printResult.ErrorMessage}");
                    
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
                AddLogMessage($"{LanguageManager.GetString("PrintRecordFailed")}: {ex.Message}");
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
            var details = $"{LanguageManager.GetString("SerialNumber")}: {record.TR_SerialNum ?? LanguageManager.GetString("NA")}\n" +
                         $"{LanguageManager.GetString("TestDateTime")}: {record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? LanguageManager.GetString("NA")}\n" +
                         $"{LanguageManager.GetString("ShortCircuitCurrent")} ({LanguageManager.GetString("Isc")}): {record.FormatNumber(record.TR_Isc)} {LanguageManager.GetString("A")}\n" +
                         $"{LanguageManager.GetString("OpenCircuitVoltage")} ({LanguageManager.GetString("Voc")}): {record.FormatNumber(record.TR_Voc)} {LanguageManager.GetString("V")}\n" +
                         $"{LanguageManager.GetString("MaximumPowerPointVoltage")} ({LanguageManager.GetString("Vpm")}): {record.FormatNumber(record.TR_Vpm)} {LanguageManager.GetString("V")}\n" +
                         $"{LanguageManager.GetString("MaximumPowerPointCurrent")} ({LanguageManager.GetString("Ipm")}): {record.FormatNumber(record.TR_Ipm)} {LanguageManager.GetString("A")}\n" +
                         $"{LanguageManager.GetString("MaximumPower")} ({LanguageManager.GetString("Pm")}): {record.FormatNumber(record.TR_Pm)} {LanguageManager.GetString("W")}\n" +
                         $"{LanguageManager.GetString("Efficiency")}: {record.FormatNumber(record.TR_CellEfficiency)}%\n" +
                         $"{LanguageManager.GetString("FillFactor")} ({LanguageManager.GetString("FF")}): {record.FormatNumber(record.TR_FF)}\n" +
                         $"{LanguageManager.GetString("Grade")}: {record.TR_Grade ?? LanguageManager.GetString("NA")}\n" +
                         $"{LanguageManager.GetString("PrintCount")}: {record.TR_Print ?? 0}";
            
            MessageBox.Show(details, LanguageManager.GetString("RecordDetails"), MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
} 