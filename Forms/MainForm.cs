using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;
using System.IO; // Added for Path and File
using System.Text; // Added for Encoding

namespace ZebraPrinterMonitor.Forms
{
    public partial class MainForm : Form
    {
        private readonly DatabaseMonitor _databaseMonitor;
        private readonly PrinterService _printerService;
        private NotifyIcon _notifyIcon;
        private PrintPreviewForm? _printPreviewForm;
        private int _totalRecordsProcessed = 0;
        private int _totalPrintJobs = 0;

        public MainForm()
        {
            InitializeComponent();
            
            _databaseMonitor = new DatabaseMonitor();
            _printerService = new PrinterService();
            
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
                    Text = "太阳能电池测试打印监控系统 v1.1.42",
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
            
            // 控件事件
            chkAutoPrint.CheckedChanged += OnAutoPrintChanged;
            cmbPrintFormat.SelectedIndexChanged += OnPrintFormatChanged;
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "太阳能电池测试打印监控系统 v1.1.42";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 600);

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

            // 设置打印格式（临时禁用事件处理器避免重复保存）
            cmbPrintFormat.SelectedIndexChanged -= OnPrintFormatChanged;
            try
            {
                if (!string.IsNullOrEmpty(config.Printer.PrintFormat))
                {
                    var formatIndex = cmbPrintFormat.Items.IndexOf(config.Printer.PrintFormat);
                    if (formatIndex >= 0)
                    {
                        cmbPrintFormat.SelectedIndex = formatIndex;
                        Logger.Info($"已加载打印格式配置: {config.Printer.PrintFormat}");
                    }
                    else
                    {
                        // 如果配置中的格式不在列表中，设置为默认值并更新配置
                        cmbPrintFormat.SelectedIndex = 0;
                        config.Printer.PrintFormat = "Text";
                        ConfigurationManager.SaveConfig();
                        Logger.Warning($"配置中的打印格式 '{config.Printer.PrintFormat}' 无效，已重置为 'Text'");
                    }
                }
                else
                {
                    // 如果配置中没有打印格式，设置默认值并保存
                    cmbPrintFormat.SelectedIndex = 0;
                    config.Printer.PrintFormat = "Text";
                    ConfigurationManager.SaveConfig();
                    Logger.Info("初始化默认打印格式: Text");
                }
                
                // 同步默认模板的格式与配置格式
                SyncDefaultTemplateFormat();
            }
            finally
            {
                // 重新启用事件处理器
                cmbPrintFormat.SelectedIndexChanged += OnPrintFormatChanged;
            }

            // 设置其他选项
            chkAutoStartMonitoring.Checked = config.Application.AutoStartMonitoring;
            chkMinimizeToTray.Checked = config.Application.MinimizeToTray;
            chkEnablePrintCount.Checked = config.Database.EnablePrintCount;  // 加载打印次数控制配置
            numPollInterval.Value = config.Database.PollInterval;
        }

        private void SyncDefaultTemplateFormat()
        {
            try
            {
                var config = ConfigurationManager.Config;
                var defaultTemplate = PrintTemplateManager.GetDefaultTemplate();
                
                if (defaultTemplate != null && !string.IsNullOrEmpty(config.Printer.PrintFormat))
                {
                    if (Enum.TryParse<PrintFormat>(config.Printer.PrintFormat, out var configFormat))
                    {
                        // 如果默认模板的格式与配置不一致，同步更新默认模板
                        if (defaultTemplate.Format != configFormat)
                        {
                            defaultTemplate.Format = configFormat;
                            PrintTemplateManager.SaveTemplate(defaultTemplate);
                            Logger.Info($"默认模板格式已同步为配置格式: {config.Printer.PrintFormat}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"同步默认模板格式失败: {ex.Message}", ex);
            }
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
                    
                    // 确保保存按钮可用
                    btnSaveTemplate.Enabled = true;
                    btnSaveTemplate.Visible = true;
                    
                    // 保存为默认模板
                    var config = ConfigurationManager.Config;
                    config.Printer.DefaultTemplate = templateName!;
                    ConfigurationManager.SaveConfig();
                    
                    AddLogMessage($"默认打印模板已更改为: {templateName}");
                    
                    // 如果预览窗口已打开，刷新预览内容
                    if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && _printPreviewForm.Visible)
                    {
                        _printPreviewForm.RefreshPreview();
                    }
                }
            }
        }

        private void btnNewTemplate_Click(object? sender, EventArgs e)
        {
            // 清空模板列表选择
            cmbTemplateList.SelectedIndex = -1;
            
            // 设置新模板的默认值
            txtTemplateName.Text = "新模板";
            txtTemplateContent.Text = "";
            cmbTemplateFormat.SelectedIndex = 0;
            
            // 确保保存按钮可用
            btnSaveTemplate.Enabled = true;
            btnSaveTemplate.Visible = true;
            
            // 聚焦到模板名称输入框
            txtTemplateName.Focus();
            txtTemplateName.SelectAll();
            
            AddLogMessage("创建新模板");
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

        private void btnVisualDesigner_Click(object? sender, EventArgs e)
        {
            try
            {
                var designerForm = new TemplateDesignerForm();
                if (designerForm.ShowDialog() == DialogResult.OK)
                {
                    // 刷新模板列表
                    LoadTemplateList();
                    AddLogMessage("模板设计器已保存新模板");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"打开模板设计器失败: {ex.Message}", ex);
                MessageBox.Show($"打开模板设计器失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSaveTemplate_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTemplateName.Text))
            {
                MessageBox.Show("请输入模板名称", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTemplateName.Focus();
                return;
            }

            try
            {
                // 检查是否为新模板或修改现有模板
                bool isNewTemplate = cmbTemplateList.SelectedIndex == -1 || 
                                    cmbTemplateList.SelectedItem?.ToString() != txtTemplateName.Text;
                
                var template = new PrintTemplate
                {
                    Name = txtTemplateName.Text,
                    Content = txtTemplateContent.Text,
                    Format = Enum.Parse<PrintFormat>(cmbTemplateFormat.Text),
                    IsDefault = false
                };

                // 如果是新模板或用户确认覆盖现有模板
                if (isNewTemplate || MessageBox.Show(
                    $"模板 '{txtTemplateName.Text}' 已存在，是否覆盖？",
                    "确认覆盖",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // 显示保存状态
                    var originalText = btnSaveTemplate.Text;
                    btnSaveTemplate.Text = "保存中...";
                    btnSaveTemplate.Enabled = false;
                    
                    try
                    {
                        PrintTemplateManager.SaveTemplate(template);
                        
                        // 刷新模板列表
                        LoadTemplateList();
                        
                        // 选择刚保存的模板
                        cmbTemplateList.SelectedItem = template.Name;
                        
                        MessageBox.Show("模板保存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        AddLogMessage($"模板保存成功: {template.Name}");
                    }
                    finally
                    {
                        btnSaveTemplate.Text = originalText;
                        btnSaveTemplate.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"保存模板失败: {ex.Message}", ex);
                MessageBox.Show($"保存模板失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                
                // 获取当前光标位置
                int cursorPosition = txtTemplateContent.SelectionStart;
                
                // 在光标位置插入字段
                txtTemplateContent.Text = txtTemplateContent.Text.Insert(cursorPosition, field);
                
                // 设置光标位置到插入字段的末尾
                txtTemplateContent.SelectionStart = cursorPosition + field.Length;
                txtTemplateContent.Focus();
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
                // 尝试选择配置中的默认模板
                var config = ConfigurationManager.Config;
                var defaultTemplate = config.Printer.DefaultTemplate;
                
                var defaultIndex = -1;
                for (int i = 0; i < cmbTemplateList.Items.Count; i++)
                {
                    if (cmbTemplateList.Items[i].ToString() == defaultTemplate)
                    {
                        defaultIndex = i;
                        break;
                    }
                }
                
                // 如果找到默认模板则选择它，否则选择第一个
                cmbTemplateList.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
            }
        }

        private void UpdateUILanguage()
        {
            // 更新主窗体标题
            this.Text = $"{LanguageManager.GetString("MainTitle")} v1.1.42";
            
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

        private void btnPrintPreview_Click(object? sender, EventArgs e)
        {
            try
            {
                // 获取当前选中的记录或最新记录
                TestRecord? recordToPreview = null;
                
                if (lvRecords.SelectedItems.Count > 0)
                {
                    // 使用选中的记录
                    var selectedItem = lvRecords.SelectedItems[0];
                    var serialNumber = selectedItem.SubItems[0].Text;
                    
                    // 从数据库获取完整记录
                    var records = _databaseMonitor.GetRecentRecords(100);
                    recordToPreview = records.FirstOrDefault(r => r.TR_SerialNum == serialNumber);
                }
                else if (lvRecords.Items.Count > 0)
                {
                    // 使用最新记录
                    var latestItem = lvRecords.Items[0];
                    var serialNumber = latestItem.SubItems[0].Text;
                    
                    // 从数据库获取完整记录
                    var records = _databaseMonitor.GetRecentRecords(100);
                    recordToPreview = records.FirstOrDefault(r => r.TR_SerialNum == serialNumber);
                }
                
                if (recordToPreview == null)
                {
                    MessageBox.Show("没有可预览的数据，请先进行测试或选择一条记录。", 
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // 创建或显示预览窗口
                if (_printPreviewForm == null || _printPreviewForm.IsDisposed)
                {
                    _printPreviewForm = new PrintPreviewForm();
                    _printPreviewForm.PrintRequested += OnPrintPreviewRequested;
                }
                
                // 加载记录和设置自动打印状态
                _printPreviewForm.LoadRecord(recordToPreview, chkAutoPrint.Checked);
                
                // 显示窗口
                if (_printPreviewForm.Visible)
                {
                    _printPreviewForm.BringToFront();
                    _printPreviewForm.RefreshPreview();
                }
                else
                {
                    _printPreviewForm.Show(this);
                }
                
                Logger.Info($"打开打印预览窗口，序列号: {recordToPreview.TR_SerialNum}");
            }
            catch (Exception ex)
            {
                Logger.Error($"打开打印预览失败: {ex.Message}", ex);
                MessageBox.Show($"打开打印预览失败:\n{ex.Message}", 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnPrintPreviewRequested(object? sender, TestRecord record)
        {
            try
            {
                // 执行打印
                var config = ConfigurationManager.Config;
                var templateName = config.Printer.DefaultTemplate;
                
                var printResult = _printerService.PrintRecord(record, config.Printer.PrintFormat, templateName);
                
                if (printResult.Success)
                {
                    _totalPrintJobs++;
                    UpdateStatusDisplay();
                    AddLogMessage($"通过预览窗口打印成功: {record.TR_SerialNum}, 打印机: {printResult.PrinterUsed}");
                    
                    // 刷新记录列表
                    LoadRecentRecords();
                }
                else
                {
                    var errorMsg = $"通过预览窗口打印失败: {printResult.ErrorMessage}";
                    AddLogMessage(errorMsg);
                    MessageBox.Show(errorMsg, "打印错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"预览窗口打印处理失败: {ex.Message}", ex);
                AddLogMessage($"预览窗口打印处理失败: {ex.Message}");
            }
        }

        private void OnAutoPrintChanged(object? sender, EventArgs e)
        {
            try
            {
                // 如果预览窗口已打开，同步更新按钮状态
                if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && _printPreviewForm.Visible)
                {
                    _printPreviewForm.SetAutoPrintMode(chkAutoPrint.Checked);
                    Logger.Info($"同步更新预览窗口自动打印状态: {chkAutoPrint.Checked}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"更新预览窗口自动打印状态失败: {ex.Message}", ex);
            }
        }

        private void OnPrintFormatChanged(object? sender, EventArgs e)
        {
            try
            {
                if (cmbPrintFormat.SelectedItem != null)
                {
                    var format = cmbPrintFormat.SelectedItem.ToString();
                    
                    // 更新配置
                    var config = ConfigurationManager.Config;
                    config.Printer.PrintFormat = format!;
                    ConfigurationManager.SaveConfig();
                    
                    // 同时更新默认模板的格式，确保打印格式选择生效
                    var defaultTemplate = PrintTemplateManager.GetDefaultTemplate();
                    if (defaultTemplate != null)
                    {
                        if (Enum.TryParse<PrintFormat>(format, out var printFormat))
                        {
                            defaultTemplate.Format = printFormat;
                            PrintTemplateManager.SaveTemplate(defaultTemplate);
                            Logger.Info($"默认模板格式已同步更新为: {format}");
                        }
                    }
                    
                    AddLogMessage($"打印格式已更改为: {format}");
                    
                    // 如果预览窗口已打开，刷新预览内容
                    if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && _printPreviewForm.Visible)
                    {
                        _printPreviewForm.RefreshPreview();
                        Logger.Info($"预览窗口已刷新，使用新的打印格式: {format}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印格式失败: {ex.Message}", ex);
            }
        }

        private void btnClearTemplate_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show("确定要清空模板内容吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                txtTemplateContent.Clear();
                rtbTemplatePreview.Clear();
                AddLogMessage("模板内容已清空");
            }
        }

        private void btnImportTemplate_Click(object? sender, EventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                    Title = "导入模板文件"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var content = File.ReadAllText(openFileDialog.FileName, Encoding.UTF8);
                    txtTemplateContent.Text = content;
                    AddLogMessage($"已导入模板文件: {Path.GetFileName(openFileDialog.FileName)}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"导入模板失败: {ex.Message}", ex);
                MessageBox.Show($"导入模板失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkPrePrintedLabel_CheckedChanged(object? sender, EventArgs e)
        {
            var checkbox = sender as CheckBox;
            if (checkbox == null) return;

            // 在模板编辑器中查找相关控件
            var templateEditor = grpTemplateEditor;
            var txtContent = templateEditor.Controls.OfType<TextBox>().FirstOrDefault(c => c.Name == "txtTemplateContent" || c.Multiline);
            var pnlDesign = templateEditor.Controls.OfType<Panel>().FirstOrDefault();
            var grpFieldPos = templateEditor.Controls.OfType<GroupBox>().FirstOrDefault(g => g.Text == "字段位置设置");

            if (txtContent != null && pnlDesign != null && grpFieldPos != null)
            {
                if (checkbox.Checked)
                {
                    // 切换到预印刷标签模式
                    txtContent.Visible = false;
                    pnlDesign.Visible = true;
                    grpFieldPos.Visible = true;
                    
                    // 清空设计面板并添加网格
                    pnlDesign.Controls.Clear();
                    DrawGridOnPanel(pnlDesign);
                }
                else
                {
                    // 切换回普通文本模式
                    txtContent.Visible = true;
                    pnlDesign.Visible = false;
                    grpFieldPos.Visible = false;
                }
            }
        }

        private void DrawGridOnPanel(Panel panel)
        {
            panel.Paint += (sender, e) =>
            {
                var g = e.Graphics;
                var gridSize = 20;
                var pen = new Pen(Color.LightGray, 1);
                
                // 绘制垂直线
                for (int x = 0; x < panel.Width; x += gridSize)
                {
                    g.DrawLine(pen, x, 0, x, panel.Height);
                }
                
                // 绘制水平线
                for (int y = 0; y < panel.Height; y += gridSize)
                {
                    g.DrawLine(pen, 0, y, panel.Width, y);
                }
                
                pen.Dispose();
            };
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 添加快捷键支持
            if (keyData == (Keys.Control | Keys.S))
            {
                // Ctrl+S 保存当前模板
                if (tabControl1.SelectedTab == tabTemplate)
                {
                    btnSaveTemplate_Click(null, null);
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _databaseMonitor?.Dispose();
                _notifyIcon?.Dispose();
                _printPreviewForm?.Dispose();
                components?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 