#nullable enable
using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;
using System.Linq; // Added for .Take()

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
        private Button btnBrowseDatabase, btnTestConnection, btnDiagnoseConnection, btnStartMonitoring, btnStopMonitoring;
        private Button btnManualPrint, btnRefresh, btnTestPrint, btnClearLog, btnSaveLog, btnPrintPreview;
        private CheckBox chkAutoPrint, chkAutoStartMonitoring, chkMinimizeToTray;
        private CheckBox chkEnablePrintCount;  // 新增打印次数控制复选框
        private ComboBox cmbPrinter, cmbPrintFormat, cmbLanguage;
        private NumericUpDown numPollInterval;
        private ListView lvRecords;
        private Label lblMonitoringStatus, lblTotalRecords, lblTotalPrints, lblLastRecord, lblPrinterStatus, lblCurrentPrint;
        
        // 打印模板页面控件
        private GroupBox grpTemplateEditor, grpTemplateList, grpTemplatePreview;
        private Label lblTemplateName, lblTemplateContent, lblTemplateFormat, lblAvailableFields, lblFontSize, lblFontName;
        private TextBox txtTemplateName, txtTemplateContent;
        private ComboBox cmbTemplateFormat, cmbTemplateList, cmbFontName;
        private NumericUpDown numFontSize;
        private Button btnSaveTemplate, btnDeleteTemplate, btnPreviewTemplate, btnNewTemplate, btnVisualDesigner, btnHeaderFooterSettings;
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
            this.btnBrowseDatabase = new Button { Text = LanguageManager.GetString("Browse"), Location = new Point(870, 43), Size = new Size(60, 27) };
            this.btnTestConnection = new Button { Text = LanguageManager.GetString("TestConnection"), Location = new Point(940, 43), Size = new Size(60, 27) };
            this.btnDiagnoseConnection = new Button { 
                Text = "诊断", 
                Location = new Point(1010, 43), 
                Size = new Size(50, 27),
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            
            this.btnBrowseDatabase.Click += btnBrowseDatabase_Click;
            this.btnTestConnection.Click += btnTestConnection_Click;
            this.btnDiagnoseConnection.Click += btnDiagnoseConnection_Click;
            
            this.grpDatabaseConfig.Controls.AddRange(new Control[] { lblDatabasePath, txtDatabasePath, btnBrowseDatabase, btnTestConnection, btnDiagnoseConnection });
            
            // 监控控制组 - 调整大小
            this.grpMonitorControl = new GroupBox { Text = LanguageManager.GetString("MonitorControl"), Location = new Point(10, 100), Size = new Size(1160, 65) };
            this.btnStartMonitoring = new Button { 
                Text = LanguageManager.GetString("StartMonitoring"), 
                Location = new Point(15, 25), 
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White
            };
            this.btnStopMonitoring = new Button { 
                Text = LanguageManager.GetString("StopMonitoring"), 
                Location = new Point(125, 25), 
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText,
                Enabled = false
            };
            this.btnPrintPreview = new Button { Text = LanguageManager.GetString("PrintPreview"), Location = new Point(235, 25), Size = new Size(100, 30), BackColor = SystemColors.Control, ForeColor = SystemColors.ControlText };
            this.chkAutoPrint = new CheckBox { Text = LanguageManager.GetString("AutoPrint"), Location = new Point(350, 30), Checked = true, AutoSize = true };
            this.chkEnablePrintCount = new CheckBox { Text = LanguageManager.GetString("EnablePrintCount"), Location = new Point(470, 30), Checked = false, AutoSize = true };
            
            this.btnStartMonitoring.Click += btnStartMonitoring_Click;
            this.btnStopMonitoring.Click += btnStopMonitoring_Click;
            this.btnPrintPreview.Click += btnPrintPreview_Click;
            this.chkEnablePrintCount.CheckedChanged += chkEnablePrintCount_CheckedChanged;
            
            this.grpMonitorControl.Controls.AddRange(new Control[] { btnStartMonitoring, btnStopMonitoring, btnPrintPreview, chkAutoPrint, chkEnablePrintCount });
            
            // 状态信息组 - 调整大小和布局
            this.grpStatus = new GroupBox { Text = LanguageManager.GetString("StatusInfo"), Location = new Point(10, 175), Size = new Size(1160, 110) };
            this.lblMonitoringStatus = new Label { Text = LanguageManager.GetString("MonitoringStatusStopped"), Location = new Point(15, 25), ForeColor = Color.Red, Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold), AutoSize = true };
            this.lblTotalRecords = new Label { Text = LanguageManager.GetString("TotalRecords"), Location = new Point(15, 50), AutoSize = true };
            this.lblTotalPrints = new Label { Text = LanguageManager.GetString("TotalPrints"), Location = new Point(200, 50), AutoSize = true };
            this.lblLastRecord = new Label { Text = LanguageManager.GetString("LastRecord"), Location = new Point(400, 50), AutoSize = true };
            this.lblCurrentPrint = new Label { Text = LanguageManager.GetString("CurrentPrintInfo"), Location = new Point(15, 75), ForeColor = Color.Blue, Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular), AutoSize = true };
            
            this.grpStatus.Controls.AddRange(new Control[] { lblMonitoringStatus, lblTotalRecords, lblTotalPrints, lblLastRecord, lblCurrentPrint });
            
            // 记录列表组 - 大幅扩大尺寸并改进ListView
            this.grpRecords = new GroupBox { Text = LanguageManager.GetString("RecentRecords"), Location = new Point(10, 295), Size = new Size(1160, 455) };
            this.lvRecords = new ListView 
            { 
                Location = new Point(15, 30), 
                Size = new Size(1130, 375), 
                View = View.Details, 
                FullRowSelect = true, 
                GridLines = true,
                MultiSelect = false,
                HideSelection = false
            };
            
            // 重新定义列，包含所有需要的数据字段
            this.lvRecords.Columns.Add(LanguageManager.GetString("SerialNumber"), 150);    // 序列号
            this.lvRecords.Columns.Add(LanguageManager.GetString("TestDateTime"), 150);    // 测试时间
            this.lvRecords.Columns.Add(LanguageManager.GetString("Current"), 120);         // ISC
            this.lvRecords.Columns.Add(LanguageManager.GetString("Voltage"), 120);         // VOC
            this.lvRecords.Columns.Add(LanguageManager.GetString("Power"), 120);           // Pm
            this.lvRecords.Columns.Add(LanguageManager.GetString("CurrentIpm"), 120);      // Ipm
            this.lvRecords.Columns.Add(LanguageManager.GetString("VoltageVpm"), 120);      // Vpm
            this.lvRecords.Columns.Add(LanguageManager.GetString("PrintCount"), 100);      // 打印次数
            this.lvRecords.Columns.Add(LanguageManager.GetString("Operation"), 200);       // 操作
            this.lvRecords.Columns.Add(LanguageManager.GetString("RecordID"), 120);        // 记录ID
            
            // 添加双击事件用于打印
            this.lvRecords.DoubleClick += LvRecords_DoubleClick;
            this.lvRecords.MouseClick += LvRecords_MouseClick;
            // 新功能3：添加单击事件用于刷新打印预览
            this.lvRecords.SelectedIndexChanged += LvRecords_SelectedIndexChanged;
            
            this.btnManualPrint = new Button { Text = LanguageManager.GetString("PrintSelected"), Location = new Point(15, 420), Size = new Size(100, 30) };
            this.btnRefresh = new Button { Text = LanguageManager.GetString("Refresh"), Location = new Point(125, 420), Size = new Size(80, 30) };
            
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
            this.lblTemplateFormat = new Label { Text = LanguageManager.GetString("PrintFormat"), Location = new Point(300, 30), AutoSize = true };
            this.cmbTemplateFormat = new ComboBox { Location = new Point(300, 50), Size = new Size(120, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.lblFontSize = new Label { Text = "字体大小:", Location = new Point(450, 30), AutoSize = true };
            this.numFontSize = new NumericUpDown { Location = new Point(450, 50), Size = new Size(80, 25), Minimum = 6, Maximum = 72, Value = 10 };
            
            this.lblFontName = new Label { Text = "字体名称:", Location = new Point(15, 80), AutoSize = true };
            this.cmbFontName = new ComboBox { Location = new Point(15, 100), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            
            // 填充系统字体
            var systemFonts = PrinterService.GetSystemFonts();
            this.cmbFontName.Items.AddRange(systemFonts.ToArray());
            this.cmbFontName.SelectedItem = "Arial";
            
            // 页眉页脚设置按钮
            this.btnHeaderFooterSettings = new Button 
            { 
                Text = "页眉页脚设置", 
                Location = new Point(250, 100), 
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            this.btnHeaderFooterSettings.FlatAppearance.BorderSize = 0;
            
            this.lblTemplateContent = new Label { Text = LanguageManager.GetString("TemplateContent"), Location = new Point(15, 140), AutoSize = true };
            
            // 调整模板内容编辑框的位置和大小
            this.txtTemplateContent = new TextBox { 
                Location = new Point(15, 160), 
                Size = new Size(550, 220), 
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
            this.btnHeaderFooterSettings.Click += btnHeaderFooterSettings_Click;
            btnClearTemplate.Click += btnClearTemplate_Click;
            btnImportTemplate.Click += btnImportTemplate_Click;
            
            // 确保所有控件都被添加到组中
            this.grpTemplateEditor.Controls.AddRange(new Control[] { 
                this.lblTemplateName, this.txtTemplateName, 
                this.lblTemplateFormat, this.cmbTemplateFormat, 
                this.lblFontSize, this.numFontSize,
                this.lblFontName, this.cmbFontName, this.btnHeaderFooterSettings,
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
            // 🔧 修复模板预览文字遮挡问题：优化RichTextBox设置
            this.rtbTemplatePreview = new RichTextBox 
            { 
                Location = new Point(15, 230), 
                Size = new Size(270, 275), 
                ReadOnly = true, 
                WordWrap = true,                        // 启用自动换行
                ScrollBars = RichTextBoxScrollBars.Both,// 添加滚动条，避免文字被遮挡
                DetectUrls = false,                     // 禁用URL检测
                Multiline = true,                       // 确保多行显示
                Font = new Font("Consolas", 9F)         // 使用等宽字体便于预览
            };
            
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

        private async void btnTestConnection_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtDatabasePath.Text))
            {
                // 🔧 修复打印预览窗口和弹窗冲突问题
                HandlePreviewFormBeforeDialog();
                MessageBox.Show("请先选择数据库文件！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                HandlePreviewFormAfterDialog();
                return;
            }

            try
            {
                AddLogMessage("🔍 正在异步测试数据库连接...");
                
                // 使用与诊断一致的异步连接测试
                if (await _databaseMonitor.ConnectAsync(this.txtDatabasePath.Text, "TestRecord"))
                {
                    AddLogMessage("✅ 数据库连接测试成功！");
                    // 🔧 修复打印预览窗口和弹窗冲突问题
                    HandlePreviewFormBeforeDialog();
                    MessageBox.Show("✅ 数据库连接成功！", "连接测试", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    HandlePreviewFormAfterDialog();
                }
                else
                {
                    AddLogMessage("❌ 数据库连接测试失败！");
                    // 🔧 修复打印预览窗口和弹窗冲突问题
                    HandlePreviewFormBeforeDialog();
                    MessageBox.Show("❌ 数据库连接失败！\n请点击'诊断'按钮查看详细错误信息。", "连接测试", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    HandlePreviewFormAfterDialog();
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"数据库连接测试失败: {ex.Message}", ex);
                AddLogMessage($"❌ 数据库连接测试异常: {ex.Message}");
                // 🔧 修复打印预览窗口和弹窗冲突问题
                HandlePreviewFormBeforeDialog();
                MessageBox.Show($"❌ 连接测试异常: {ex.Message}\n请点击'诊断'按钮查看详细解决方案。", "连接测试", MessageBoxButtons.OK, MessageBoxIcon.Error);
                HandlePreviewFormAfterDialog();
            }
        }

        private void btnDiagnoseConnection_Click(object? sender, EventArgs e)
        {
            try
            {
                AddLogMessage("🔍 开始数据库连接诊断...");
                
                string databasePath = this.txtDatabasePath.Text;
                if (string.IsNullOrEmpty(databasePath))
                {
                    databasePath = "（未选择）";
                }
                
                var (success, message) = DatabaseConnectionHelper.DiagnoseConnection(this.txtDatabasePath.Text, "TestRecord");
                
                if (success)
                {
                    AddLogMessage("✅ 诊断完成：连接正常！");
                    // 🔧 修复打印预览窗口和弹窗冲突问题
                    HandlePreviewFormBeforeDialog();
                    MessageBox.Show(message, "🔍 连接诊断 - 成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    HandlePreviewFormAfterDialog();
                }
                else
                {
                    AddLogMessage("❌ 诊断完成：发现问题！");
                    // 🔧 修复打印预览窗口和弹窗冲突问题
                    HandlePreviewFormBeforeDialog();
                    MessageBox.Show(message, "🔍 连接诊断 - 发现问题", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    HandlePreviewFormAfterDialog();
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"连接诊断失败: {ex.Message}", ex);
                AddLogMessage($"❌ 诊断过程出错: {ex.Message}");
                // 🔧 修复打印预览窗口和弹窗冲突问题
                HandlePreviewFormBeforeDialog();
                MessageBox.Show($"❌ 诊断过程出错: {ex.Message}", "诊断错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                HandlePreviewFormAfterDialog();
            }
        }

        private async void btnStartMonitoring_Click(object? sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txtDatabasePath.Text))
                {
                    AddLogMessage("❌ 请先选择数据库文件路径！");
                    return;
                }

                var config = ConfigurationManager.Config.Database;
                
                AddLogMessage($"🔗 正在异步连接数据库: {this.txtDatabasePath.Text}");
                
                // 使用异步连接数据库
                if (!await _databaseMonitor.ConnectAsync(this.txtDatabasePath.Text, config.TableName))
                {
                    AddLogMessage("❌ 数据库连接失败！请检查数据库文件路径和格式");
                    return;
                }

                AddLogMessage("✅ 数据库连接成功");
                
                // 获取表字段信息
                var columns = _databaseMonitor.GetTableColumns(config.TableName);
                AddLogMessage($"📊 表结构: {columns.Count} 个字段");
                
                // 开始监控
                _databaseMonitor.StartMonitoring(config.PollInterval);
                AddLogMessage("🚀 数据库监控已启动");
                AddLogMessage($"⚡ 监控间隔: {config.PollInterval}ms");
                
                // 更新按钮状态
                UpdateMonitoringButtonStates(true);
                
                // 立即加载一次数据
                LoadRecentRecords();
                
                AddLogMessage("✅ 监控启动完成，等待数据变化...");
            }
            catch (Exception ex)
            {
                Logger.Error($"启动监控失败: {ex.Message}", ex);
                AddLogMessage($"❌ 启动监控失败: {ex.Message}");
            }
        }

        private void btnStopMonitoring_Click(object? sender, EventArgs e)
        {
            try
            {
                _databaseMonitor.StopMonitoring();
                AddLogMessage("数据监控已停止");
                
                // 更新按钮状态 - 未监控
                UpdateMonitoringButtonStates(false);
                
                MessageBox.Show(LanguageManager.GetString("MonitoringStopped"), LanguageManager.GetString("Information"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"停止监控失败: {ex.Message}", ex);
                MessageBox.Show($"{LanguageManager.GetString("Error")}: {ex.Message}", LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 更新监控按钮的状态和颜色
        /// </summary>
        /// <param name="isMonitoring">是否正在监控</param>
        private void UpdateMonitoringButtonStates(bool isMonitoring)
        {
            if (isMonitoring)
            {
                // 监控中状态
                // 开始监控按钮：绿色且不可用
                btnStartMonitoring.BackColor = Color.Green;
                btnStartMonitoring.ForeColor = Color.White;
                btnStartMonitoring.FlatStyle = FlatStyle.Flat;
                btnStartMonitoring.Enabled = false;
                
                // 停止监控按钮：红色且可用
                btnStopMonitoring.BackColor = Color.Red;
                btnStopMonitoring.ForeColor = Color.White;
                btnStopMonitoring.FlatStyle = FlatStyle.Flat;
                btnStopMonitoring.Enabled = true;
            }
            else
            {
                // 未监控状态
                // 开始监控按钮：正常色且可用
                btnStartMonitoring.BackColor = Color.FromArgb(0, 123, 255); // 蓝色
                btnStartMonitoring.ForeColor = Color.White;
                btnStartMonitoring.FlatStyle = FlatStyle.Flat;
                btnStartMonitoring.Enabled = true;
                
                // 停止监控按钮：正常色且不可用
                btnStopMonitoring.BackColor = SystemColors.Control; // 默认背景色
                btnStopMonitoring.ForeColor = SystemColors.ControlText; // 默认文字色
                btnStopMonitoring.FlatStyle = FlatStyle.Flat;
                btnStopMonitoring.Enabled = false;
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
                    _databaseMonitor.UpdatePrintCount(record);
                    
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
            try
            {
                AddLogMessage("🔄 手动刷新：强制检查数据库变化...");
                
                // 触发数据库监控的强制检查
                if (_databaseMonitor.IsMonitoring)
                {
                    _databaseMonitor.ForceRefresh();
                    AddLogMessage("✅ 强制检查已触发");
                }
                else
                {
                    AddLogMessage("⚠️ 监控未启动，仅刷新记录列表");
                }
                
                // 刷新记录列表
                LoadRecentRecords();
                
                // 更新状态显示
                UpdateStatusDisplay();
                
                AddLogMessage("🔄 记录列表已刷新");
            }
            catch (Exception ex)
            {
                Logger.Error($"手动刷新失败: {ex.Message}", ex);
                AddLogMessage($"❌ 刷新失败: {ex.Message}");
            }
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
            
            // 更新ListView列显示
            UpdatePrintCountColumnVisibility();
            
            // 刷新数据显示
            btnRefresh_Click(null, EventArgs.Empty);
            
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

                // 更新当前打印信息
                UpdateCurrentPrintInfo(record, "用户双击选择");

                var config = ConfigurationManager.Config.Printer;
                // 使用默认模板或配置的模板
                var templateName = config.DefaultTemplate;
                var printResult = _printerService.PrintRecord(record, config.PrintFormat, templateName);
                
                if (printResult.Success)
                {
                    _databaseMonitor.UpdatePrintCount(record);
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

                // 打印完成后清除当前打印信息
                UpdateCurrentPrintInfo();
            }
            catch (Exception ex)
            {
                Logger.Error($"打印记录失败: {ex.Message}", ex);
                AddLogMessage($"{LanguageManager.GetString("PrintRecordFailed")}: {ex.Message}");
                // 出错时也清除当前打印信息
                UpdateCurrentPrintInfo();
            }
        }

        private bool ConfirmPrintIfAlreadyPrinted(TestRecord record)
        {
            var config = ConfigurationManager.Config;
            
            // 只有启用打印次数统计时才检查重打
            if (!config.Database.EnablePrintCount)
            {
                return true; // 未启用打印次数统计，直接允许打印
            }
            
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

        // 新功能3：数据列表选择改变时自动刷新打印预览窗口
        private void LvRecords_SelectedIndexChanged(object? sender, EventArgs e)
        {
            try
            {
                // 如果打印预览窗口已开启且有选中的记录，自动刷新预览数据
                if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && _printPreviewForm.Visible && 
                    lvRecords.SelectedItems.Count > 0)
                {
                    var selectedItem = lvRecords.SelectedItems[0];
                    var record = selectedItem.Tag as TestRecord;
                    
                    if (record != null)
                    {
                        // 更新打印预览窗口的数据
                        _printPreviewForm.LoadRecord(record);
                        _printPreviewForm.SetAutoPrintMode(chkAutoPrint.Checked);
                        Logger.Info($"📺 打印预览窗口已更新为选中记录: {record.TR_SerialNum}");
                        AddLogMessage($"📺 预览已更新: {record.TR_SerialNum}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印预览失败: {ex.Message}", ex);
                AddLogMessage($"❌ 更新预览失败: {ex.Message}");
            }
        }
    }
} 