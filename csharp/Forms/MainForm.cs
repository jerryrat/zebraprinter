using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;
using System.IO; // Added for Path and File

namespace ZebraPrinterMonitor.Forms
{
    public partial class MainForm : Form
    {
        private readonly DatabaseMonitor _databaseMonitor;
        private readonly PrinterService _printerService;
        private NotifyIcon _notifyIcon;
        private PrintMonitorForm? _printMonitorForm;  // 小监控窗口
        private int _totalRecordsProcessed = 0;
        private int _totalPrintJobs = 0;

        public MainForm()
        {
            InitializeComponent();
            
            _databaseMonitor = new DatabaseMonitor();
            _printerService = new PrinterService();
            
            // 初始化小监控窗口
            _printMonitorForm = new PrintMonitorForm(_printerService);
            _printMonitorForm.ShowMainRequested += OnShowMainRequested;
            _printMonitorForm.PrintRequested += OnPrintRequested;
            
            SetupNotifyIcon();
            SetupEventHandlers();
            InitializeUI();
            
            Logger.Info("主窗体初始化完成");
        }

        private void SetupNotifyIcon()
        {
            try
            {
                // 尝试从嵌入资源加载Zebra图标
                Icon zebraIcon = null;
                
                try
                {
                    // 优先从嵌入资源加载图标
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceNames = assembly.GetManifestResourceNames();
                    
                    // 尝试不同的资源名称
                    var possibleNames = new[] { "Zebra.ico", "zebra_icon.ico" };
                    
                    foreach (var resourceName in resourceNames)
                    {
                        foreach (var possibleName in possibleNames)
                        {
                            if (resourceName.EndsWith(possibleName))
                            {
                                using (var stream = assembly.GetManifestResourceStream(resourceName))
                                {
                                    if (stream != null)
                                    {
                                        zebraIcon = new Icon(stream);
                                        Logger.Info($"成功从嵌入资源加载图标: {resourceName}");
                                        break;
                                    }
                                }
                            }
                        }
                        if (zebraIcon != null) break;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"从嵌入资源加载图标失败: {ex.Message}");
                }

                // 如果嵌入资源失败，尝试从文件加载
                if (zebraIcon == null)
                {
                    var iconPaths = new[]
                    {
                        "Zebra.ico",
                        "zebra_icon.ico",
                        Path.Combine(Application.StartupPath, "Zebra.ico"),
                        Path.Combine(Application.StartupPath, "zebra_icon.ico"),
                        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Zebra.ico"),
                        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "zebra_icon.ico")
                    };

                    foreach (var iconPath in iconPaths)
                    {
                        if (File.Exists(iconPath))
                        {
                            try
                            {
                                zebraIcon = new Icon(iconPath);
                                Logger.Info($"成功从文件加载图标: {iconPath}");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Logger.Warning($"加载图标文件失败 {iconPath}: {ex.Message}");
                            }
                        }
                    }
                }

                // 设置系统托盘图标
                _notifyIcon = new NotifyIcon
                {
                    Icon = zebraIcon ?? SystemIcons.Application, // 如果找不到Zebra图标则使用默认图标
                    Text = "太阳能电池测试打印监控系统 v1.1.21",
                    Visible = false
                };

                // 设置窗体图标也使用相同的图标
                if (zebraIcon != null)
                {
                    this.Icon = zebraIcon;
                }
                else
                {
                    // 如果没有找到自定义图标，使用系统默认图标
                    this.Icon = SystemIcons.Application;
                    Logger.Warning("未找到自定义图标，使用系统默认图标");
                }

                // 设置托盘菜单
                var contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("显示主界面", null, (s, e) => ShowMainWindow());
                contextMenu.Items.Add("-"); // 分隔线
                contextMenu.Items.Add("退出程序", null, (s, e) => ExitApplication());
                
                _notifyIcon.ContextMenuStrip = contextMenu;
                _notifyIcon.DoubleClick += (s, e) => ShowMainWindow();
                
                Logger.Info("系统托盘图标设置完成");
            }
            catch (Exception ex)
            {
                Logger.Error($"设置系统托盘图标失败: {ex.Message}", ex);
                
                // 如果出错，创建基本的NotifyIcon
                _notifyIcon = new NotifyIcon
                {
                    Icon = SystemIcons.Application,
                    Text = "太阳能电池测试打印监控系统",
                    Visible = false
                };
                
                // 确保窗体也有图标
                this.Icon = SystemIcons.Application;
            }
        }

        private void SetupEventHandlers()
        {
            // 数据库监控事件
            _databaseMonitor.NewRecordFound += OnNewRecordFound;
            _databaseMonitor.MonitoringError += OnMonitoringError;
            _databaseMonitor.StatusChanged += OnStatusChanged;

            // 窗体事件
            this.Load += OnFormLoad;
            this.Resize += OnFormResize;
            this.FormClosing += OnFormClosing;
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "太阳能电池测试打印监控系统 v1.1.27";
            this.Size = new Size(1200, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 650);

            // 加载配置并更新UI
            LoadConfiguration();
            UpdatePrinterList();
            UpdateStatusDisplay();

            Logger.Info("UI初始化完成");
        }

        private void LoadConfiguration()
        {
            var config = ConfigurationManager.Config;

            // 设置数据库路径
            if (!string.IsNullOrEmpty(config.Database.DatabasePath))
            {
                txtDatabasePath.Text = config.Database.DatabasePath;
            }

            // 设置打印机
            if (!string.IsNullOrEmpty(config.Printer.PrinterName))
            {
                // 在UpdatePrinterList后会自动选择
            }

            // 设置其他选项
            chkAutoStartMonitoring.Checked = config.Application.AutoStartMonitoring;
            chkMinimizeToTray.Checked = config.Application.MinimizeToTray;
            chkEnablePrintCount.Checked = config.Database.EnablePrintCount;  // 加载打印次数控制配置
            numPollInterval.Value = config.Database.PollInterval;
        }

        private void UpdatePrinterList()
        {
            try
            {
                var printers = _printerService.GetAvailablePrinters();
                cmbPrinter.Items.Clear();
                cmbPrinter.Items.AddRange(printers.ToArray());

                var config = ConfigurationManager.Config;
                if (!string.IsNullOrEmpty(config.Printer.PrinterName) && 
                    printers.Contains(config.Printer.PrinterName))
                {
                    cmbPrinter.SelectedItem = config.Printer.PrinterName;
                }
                else if (cmbPrinter.Items.Count > 0)
                {
                    cmbPrinter.SelectedIndex = 0;
                }

                lblPrinterStatus.Text = $"找到 {printers.Count} 个打印机";
                lblPrinterStatus.ForeColor = printers.Count > 0 ? Color.Green : Color.Red;
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印机列表失败: {ex.Message}", ex);
                lblPrinterStatus.Text = "获取打印机列表失败";
                lblPrinterStatus.ForeColor = Color.Red;
            }
        }

        private void UpdateStatusDisplay()
        {
            lblMonitoringStatus.Text = _databaseMonitor.IsMonitoring ? "监控中..." : "已停止";
            lblMonitoringStatus.ForeColor = _databaseMonitor.IsMonitoring ? Color.Green : Color.Red;
            
            btnStartMonitoring.Enabled = !_databaseMonitor.IsMonitoring;
            btnStopMonitoring.Enabled = _databaseMonitor.IsMonitoring;
            
            lblTotalRecords.Text = $"处理记录: {_totalRecordsProcessed}";
            lblTotalPrints.Text = $"打印任务: {_totalPrintJobs}";
            lblLastRecord.Text = $"最后记录: {_databaseMonitor.LastRecordId}";
        }

        private void OnNewRecordFound(object? sender, TestRecord record)
        {
            // 使用Invoke确保在UI线程上执行
            this.Invoke(new Action(() =>
            {
                try
                {
                    _totalRecordsProcessed++;
                    
                    // 更新小监控窗口
                    _printMonitorForm?.UpdateRecord(record);
                    
                    // 添加到记录列表
                    var item = new ListViewItem(record.TR_SerialNum ?? "N/A");
                    item.SubItems.Add(record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
                    item.SubItems.Add(record.FormatNumber(record.TR_Isc));
                    item.SubItems.Add(record.FormatNumber(record.TR_Voc));
                    item.SubItems.Add(record.FormatNumber(record.TR_Vpm));  // 新增Vpm列
                    item.SubItems.Add(record.FormatNumber(record.TR_Pm));
                    item.SubItems.Add((record.TR_Print ?? 0).ToString());    // 新增打印次数列
                    item.SubItems.Add("双击打印");                            // 新增操作列
                    item.Tag = record;
                    
                    // 根据打印次数设置颜色
                    if (record.TR_Print > 0)
                    {
                        item.ForeColor = Color.Gray;  // 已打印的记录显示为灰色
                    }
                    else
                    {
                        item.ForeColor = Color.Black; // 未打印的记录显示为黑色
                    }
                    
                    lvRecords.Items.Insert(0, item);
                    
                    // 限制显示的记录数量
                    while (lvRecords.Items.Count > 100)
                    {
                        lvRecords.Items.RemoveAt(lvRecords.Items.Count - 1);
                    }

                    // 自动打印
                    if (chkAutoPrint.Checked)
                    {
                        AutoPrintRecord(record);
                    }

                    UpdateStatusDisplay();
                    
                    // 显示通知
                    ShowNotification($"新记录: {record.TR_SerialNum}", "发现新的测试记录");
                }
                catch (Exception ex)
                {
                    Logger.Error($"处理新记录失败: {ex.Message}", ex);
                }
            }));
        }

        private void AutoPrintRecord(TestRecord record)
        {
            try
            {
                var config = ConfigurationManager.Config;
                // 使用默认模板或配置的模板
                var templateName = config.Printer.DefaultTemplate;
                var printResult = _printerService.PrintRecord(record, config.Printer.PrintFormat, templateName);

                if (printResult.Success)
                {
                    _totalPrintJobs++;
                    
                    // 更新数据库打印计数
                    _databaseMonitor.UpdatePrintCount(record.TR_SerialNum ?? record.TR_ID ?? "");
                    
                    Logger.Info($"自动打印完成: {record.TR_SerialNum}");
                    AddLogMessage($"自动打印: {record.TR_SerialNum} -> {printResult.PrinterUsed}");
                }
                else
                {
                    Logger.Error($"自动打印失败: {printResult.ErrorMessage}");
                    AddLogMessage($"打印失败: {record.TR_SerialNum} - {printResult.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"自动打印异常: {ex.Message}", ex);
                AddLogMessage($"打印异常: {record.TR_SerialNum} - {ex.Message}");
            }
        }

        private void OnMonitoringError(object? sender, string error)
        {
            this.Invoke(new Action(() =>
            {
                AddLogMessage($"监控错误: {error}");
                UpdateStatusDisplay();
                ShowNotification("监控错误", error);
            }));
        }

        private void OnStatusChanged(object? sender, string status)
        {
            this.Invoke(new Action(() =>
            {
                AddLogMessage($"状态变更: {status}");
                UpdateStatusDisplay();
            }));
        }

        private void AddLogMessage(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var logEntry = $"[{timestamp}] {message}";
            
            txtLog.AppendText(logEntry + Environment.NewLine);
            txtLog.ScrollToCaret();
            
            // 限制日志行数
            var lines = txtLog.Lines;
            if (lines.Length > 1000)
            {
                txtLog.Lines = lines.Skip(500).ToArray();
            }
        }

        private void ShowNotification(string title, string text)
        {
            if (_notifyIcon.Visible)
            {
                _notifyIcon.ShowBalloonTip(3000, title, text, ToolTipIcon.Info);
            }
        }

        private void ShowMainWindow()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
            _notifyIcon.Visible = false;
        }

        private void OnFormLoad(object? sender, EventArgs e)
        {
            // 窗体完全加载后再加载数据
            LoadRecentRecords();
            
            // 加载打印模板列表
            LoadTemplateList();
            
            // 设置语言
            var config = ConfigurationManager.Config;
            LanguageManager.CurrentLanguage = config.UI.Language;
            if (LanguageManager.CurrentLanguage == "zh-CN")
            {
                cmbLanguage.SelectedIndex = 0;
            }
            else
            {
                cmbLanguage.SelectedIndex = 1;
            }
            
            // 更新界面语言
            UpdateUILanguage();
            
            // 检查打印机安装状态
            CheckPrinterInstallation();
        }

        private void CheckPrinterInstallation()
        {
            if (!_printerService.HasAnyPrinter())
            {
                var title = LanguageManager.GetString("NoPrinterTitle");
                var message = _printerService.GetNoPrinterMessage();
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AddLogMessage(LanguageManager.GetString("NoPrinterFound"));
            }
        }

        private void OnFormResize(object? sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized && chkMinimizeToTray.Checked)
            {
                this.Hide();
                _notifyIcon.Visible = true;
                ShowNotification("程序已最小化到系统托盘", "双击托盘图标可恢复窗口");
            }
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing && chkMinimizeToTray.Checked)
            {
                e.Cancel = true;
                this.Hide();
                _notifyIcon.Visible = true;
                ShowNotification("程序已最小化到系统托盘", "程序仍在后台运行");
            }
            else
            {
                ExitApplication();
            }
        }

        private void ExitApplication()
        {
            try
            {
                _databaseMonitor.StopMonitoring();
                _notifyIcon.Visible = false;
                Application.Exit();
            }
            catch (Exception ex)
            {
                Logger.Error($"退出程序时发生错误: {ex.Message}", ex);
            }
        }

        private void LoadRecentRecords()
        {
            try
            {
                var config = ConfigurationManager.Config.Database;
                
                // 如果数据库路径为空，则不加载数据
                if (string.IsNullOrEmpty(config.DatabasePath))
                {
                    AddLogMessage("数据库路径未设置，跳过数据加载");
                    return;
                }

                // 如果数据库文件不存在，则不加载数据
                if (!System.IO.File.Exists(config.DatabasePath))
                {
                    AddLogMessage($"数据库文件不存在: {config.DatabasePath}");
                    return;
                }

                // 尝试连接数据库并获取数据
                if (!_databaseMonitor.Connect(config.DatabasePath, config.TableName, config.MonitorField))
                {
                    AddLogMessage("数据库连接失败，无法加载数据");
                    return;
                }

                var records = _databaseMonitor.GetRecentRecords(50);
                
                lvRecords.Items.Clear();
                
                foreach (var record in records)
                {
                    var item = new ListViewItem(record.TR_SerialNum ?? "N/A");
                    item.SubItems.Add(record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
                    item.SubItems.Add(record.FormatNumber(record.TR_Isc));
                    item.SubItems.Add(record.FormatNumber(record.TR_Voc));
                    item.SubItems.Add(record.FormatNumber(record.TR_Vpm));  // 新增Vpm列
                    item.SubItems.Add(record.FormatNumber(record.TR_Pm));
                    item.SubItems.Add((record.TR_Print ?? 0).ToString());    // 新增打印次数列
                    item.SubItems.Add("双击打印");                            // 新增操作列
                    item.Tag = record;
                    
                    // 根据打印次数设置颜色
                    if (record.TR_Print > 0)
                    {
                        item.ForeColor = Color.Gray;  // 已打印的记录显示为灰色
                    }
                    else
                    {
                        item.ForeColor = Color.Black; // 未打印的记录显示为黑色
                    }
                    
                    lvRecords.Items.Add(item);
                }
                
                AddLogMessage($"已加载 {records.Count} 条最近记录（按测试日期倒序）");
                Logger.Info($"已加载 {records.Count} 条最近记录");
            }
            catch (Exception ex)
            {
                Logger.Error($"加载最近记录失败: {ex.Message}", ex);
                AddLogMessage($"加载最近记录失败: {ex.Message}");
            }
        }

        // 语言选择事件处理
        private void cmbLanguage_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbLanguage.SelectedIndex == 0)
            {
                LanguageManager.CurrentLanguage = "zh-CN";
            }
            else if (cmbLanguage.SelectedIndex == 1)
            {
                LanguageManager.CurrentLanguage = "en-US";
            }

            // 保存语言设置
            var config = ConfigurationManager.Config;
            config.UI.Language = LanguageManager.CurrentLanguage;
            ConfigurationManager.SaveConfig();

            // 更新界面文本
            UpdateUILanguage();
            
            // 更新小监控窗口语言
            _printMonitorForm?.RefreshLanguage();
            
            AddLogMessage($"语言已切换到: {LanguageManager.GetLanguageName(LanguageManager.CurrentLanguage)}");
        }

        // 打印模板事件处理程序
        private void cmbTemplateList_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbTemplateList.SelectedItem != null)
            {
                var templateName = cmbTemplateList.SelectedItem.ToString();
                var template = PrintTemplateManager.GetTemplate(templateName!);
                if (template != null)
                {
                    txtTemplateName.Text = template.Name;
                    txtTemplateContent.Text = template.Content;
                    cmbTemplateFormat.SelectedItem = template.Format.ToString();
                }
            }
        }

        private void btnNewTemplate_Click(object? sender, EventArgs e)
        {
            txtTemplateName.Text = "新模板";
            txtTemplateContent.Text = "";
            cmbTemplateFormat.SelectedIndex = 0;
        }

        private void btnDeleteTemplate_Click(object? sender, EventArgs e)
        {
            if (cmbTemplateList.SelectedItem != null)
            {
                var templateName = cmbTemplateList.SelectedItem.ToString();
                if (MessageBox.Show($"确定要删除模板 '{templateName}' 吗？", "确认删除", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    PrintTemplateManager.DeleteTemplate(templateName!);
                    LoadTemplateList();
                    AddLogMessage($"已删除模板: {templateName}");
                }
            }
        }

        private void btnSaveTemplate_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTemplateName.Text))
            {
                MessageBox.Show("请输入模板名称", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var template = new PrintTemplate
            {
                Name = txtTemplateName.Text,
                Content = txtTemplateContent.Text,
                Format = Enum.Parse<PrintFormat>(cmbTemplateFormat.SelectedItem?.ToString() ?? "Text")
            };

            PrintTemplateManager.SaveTemplate(template);
            LoadTemplateList();
            AddLogMessage($"已保存模板: {template.Name}");
        }

        private void btnPreviewTemplate_Click(object? sender, EventArgs e)
        {
            if (lvRecords.SelectedItems.Count > 0)
            {
                var record = (TestRecord)lvRecords.SelectedItems[0].Tag!;
                var template = new PrintTemplate
                {
                    Name = txtTemplateName.Text,
                    Content = txtTemplateContent.Text,
                    Format = Enum.Parse<PrintFormat>(cmbTemplateFormat.SelectedItem?.ToString() ?? "Text")
                };

                var preview = PrintTemplateManager.ProcessTemplate(template, record);
                rtbTemplatePreview.Text = preview;
            }
            else
            {
                MessageBox.Show("请先选择一条测试记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void lstAvailableFields_DoubleClick(object? sender, EventArgs e)
        {
            if (lstAvailableFields.SelectedItem != null)
            {
                var field = lstAvailableFields.SelectedItem.ToString();
                txtTemplateContent.Text += field;
                txtTemplateContent.Focus();
                txtTemplateContent.SelectionStart = txtTemplateContent.Text.Length;
            }
        }

        private void btnAdvancedEditor_Click(object? sender, EventArgs e)
        {
            try
            {
                // 获取当前选中的模板
                PrintTemplate? currentTemplate = null;
                if (cmbTemplateList.SelectedItem != null)
                {
                    string templateName = cmbTemplateList.SelectedItem.ToString() ?? "";
                    currentTemplate = PrintTemplateManager.GetTemplate(templateName);
                }

                // 打开高级模板编辑器
                using (var editorForm = new TemplateEditorForm(currentTemplate))
                {
                    if (editorForm.ShowDialog(this) == DialogResult.OK)
                    {
                        // 刷新模板列表
                        LoadTemplateList();
                        Logger.Info("模板编辑器已关闭，模板列表已刷新");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"打开高级模板编辑器失败: {ex.Message}", ex);
                MessageBox.Show($"打开高级模板编辑器失败: {ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadTemplateList()
        {
            cmbTemplateList.Items.Clear();
            var templates = PrintTemplateManager.GetTemplates();
            foreach (var template in templates)
            {
                cmbTemplateList.Items.Add(template.Name);
            }

            if (cmbTemplateList.Items.Count > 0)
            {
                cmbTemplateList.SelectedIndex = 0;
            }
        }

        private void UpdateUILanguage()
        {
            // 更新主窗体标题
            this.Text = $"{LanguageManager.GetString("MainTitle")} v1.1.27";
            
            // 更新选项卡标题
            if (tabControl1.TabPages.Count >= 4)
            {
                tabControl1.TabPages[0].Text = LanguageManager.GetString("TabDataMonitoring");
                tabControl1.TabPages[1].Text = LanguageManager.GetString("TabSystemConfig");
                tabControl1.TabPages[2].Text = LanguageManager.GetString("TabPrintTemplate");
                tabControl1.TabPages[3].Text = LanguageManager.GetString("TabRuntimeLogs");
            }

            // 更新ListView列标题
            if (lvRecords.Columns.Count >= 8)
            {
                lvRecords.Columns[0].Text = LanguageManager.GetString("SerialNumber");
                lvRecords.Columns[1].Text = LanguageManager.GetString("TestDateTime");
                lvRecords.Columns[2].Text = LanguageManager.GetString("Current");
                lvRecords.Columns[3].Text = LanguageManager.GetString("Voltage");
                lvRecords.Columns[4].Text = LanguageManager.GetString("VoltageVpm");
                lvRecords.Columns[5].Text = LanguageManager.GetString("Power");
                lvRecords.Columns[6].Text = LanguageManager.GetString("PrintCount");
                lvRecords.Columns[7].Text = LanguageManager.GetString("Operation");
            }

            // 更新按钮和标签文本
            lblLanguage.Text = LanguageManager.GetString("Language");
            lblDatabasePath.Text = LanguageManager.GetString("DatabasePath");
            lblPrinter.Text = LanguageManager.GetString("SelectedPrinter");
            lblPrintFormat.Text = LanguageManager.GetString("PrintFormat");
            lblPollInterval.Text = LanguageManager.GetString("PollInterval");
            
            chkAutoStartMonitoring.Text = LanguageManager.GetString("AutoStartMonitoring");
            chkMinimizeToTray.Text = LanguageManager.GetString("MinimizeToTray");
            
            btnStartMonitoring.Text = LanguageManager.GetString("StartMonitoring");
            btnStopMonitoring.Text = LanguageManager.GetString("StopMonitoring");
            btnTestPrint.Text = LanguageManager.GetString("Print");
            btnClearLog.Text = LanguageManager.GetString("ClearLogs");
        }

        // 小监控窗口事件处理
        private void OnShowMainRequested(object? sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
            this.Activate();
        }

        private void OnPrintRequested(object? sender, TestRecord record)
        {
            try
            {
                var config = ConfigurationManager.Config;
                var templateName = config.Printer.DefaultTemplate;
                var printResult = _printerService.PrintRecord(record, config.Printer.PrintFormat, templateName);

                if (printResult.Success)
                {
                    _totalPrintJobs++;
                    
                    // 更新数据库打印计数
                    _databaseMonitor.UpdatePrintCount(record.TR_SerialNum ?? record.TR_ID ?? "");
                    
                    Logger.Info($"小监控窗口打印完成: {record.TR_SerialNum}");
                    AddLogMessage($"小监控窗口打印: {record.TR_SerialNum} -> {printResult.PrinterUsed}");
                    
                    // 通知小监控窗口打印成功
                    _printMonitorForm?.ShowNotification(LanguageManager.GetString("PrintSuccess"), MessageType.Success);
                    
                    // 刷新主界面列表
                    btnRefresh_Click(null, EventArgs.Empty);
                }
                else
                {
                    Logger.Error($"小监控窗口打印失败: {printResult.ErrorMessage}");
                    AddLogMessage($"打印失败: {record.TR_SerialNum} - {printResult.ErrorMessage}");
                    
                    // 通知小监控窗口打印失败
                    _printMonitorForm?.ShowNotification($"{LanguageManager.GetString("PrintError")}: {printResult.ErrorMessage}", MessageType.Error);
                }
                
                UpdateStatusDisplay();
            }
            catch (Exception ex)
            {
                Logger.Error($"小监控窗口打印异常: {ex.Message}", ex);
                AddLogMessage($"打印异常: {record.TR_SerialNum} - {ex.Message}");
                _printMonitorForm?.ShowNotification($"{LanguageManager.GetString("PrintError")}: {ex.Message}", MessageType.Error);
            }
        }

        // 显示/隐藏小监控窗口
        public void ShowPrintMonitor()
        {
            if (_printMonitorForm != null)
            {
                _printMonitorForm.Show();
                _printMonitorForm.BringToFront();
            }
        }

        public void HidePrintMonitor()
        {
            _printMonitorForm?.Hide();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _databaseMonitor?.Dispose();
                _notifyIcon?.Dispose();
                components?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 