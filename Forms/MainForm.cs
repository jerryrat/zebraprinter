using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;
using System.IO; // Added for Path and File
using System.Text; // Added for Encoding
using System.Collections.Generic; // Added for List

namespace ZebraPrinterMonitor.Forms
{
    public partial class MainForm : Form
    {


        // 实例变量
        private DatabaseMonitor _databaseMonitor;
        private PrinterService _printerService;
        private NotifyIcon _notifyIcon;
        private System.Windows.Forms.Timer _statusUpdateTimer;
        private PrintPreviewForm? _printPreviewForm;
        private int _totalRecordsProcessed = 0;
        private int _totalPrintJobs = 0;
        
        // 页眉页脚相关字段
        private bool _showHeader = false;
        private bool _showFooter = false;
        private string _headerText = "";
        private string _footerText = "";
        private string _headerImagePath = "";
        private string _footerImagePath = "";

        public MainForm()
        {
            InitializeComponent();
            InitializeDatabaseMonitor(); // 🔧 使用统一监控系统初始化
            InitializePrinterService();
            SetupNotifyIcon();
            InitializeTimer();
            
            // 设置默认状态 - 不开启监控
            UpdateMonitoringButtonStates(false);
            
            // 窗体事件
            this.Load += OnFormLoad;
            this.Resize += OnFormResize;
            this.FormClosing += OnFormClosing;
            
            // 控件事件
            chkAutoPrint.CheckedChanged += OnAutoPrintChanged;
            cmbPrintFormat.SelectedIndexChanged += OnPrintFormatChanged;
        }
        
        /// <summary>
        /// 初始化打印服务
        /// </summary>
        private void InitializePrinterService()
        {
            try
            {
                _printerService = new PrinterService();
                Logger.Info("✅ 打印服务初始化完成");
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 打印服务初始化失败: {ex.Message}", ex);
            }
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
                    Text = "太阳能电池测试打印监控系统 v1.2.2 - 连接诊断增强版",
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
                contextMenu.Items.Add(LanguageManager.GetString("ShowMainWindow"), null, (s, e) => ShowMainWindow());
                contextMenu.Items.Add("-"); // 分隔线
                contextMenu.Items.Add(LanguageManager.GetString("ExitProgram"), null, (s, e) => ExitApplication());
                
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
            this.Text = "太阳能电池测试打印监控系统 v1.3.8 - 排序逻辑和时间显示修复版";
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
                Logger.Info("开始更新打印机列表...");
                AddLogMessage("正在获取打印机列表...");
                
                var printers = _printerService.GetAvailablePrinters();
                Logger.Info($"获取到打印机列表，数量: {printers.Count}");
                
                if (printers.Count > 0)
                {
                    Logger.Info($"打印机列表: {string.Join(", ", printers)}");
                    AddLogMessage($"找到 {printers.Count} 台打印机: {string.Join(", ", printers)}");
                }
                else
                {
                    Logger.Warning("没有找到任何打印机");
                    AddLogMessage("警告: 没有找到任何打印机");
                }
                
                cmbPrinter.Items.Clear();
                cmbPrinter.Items.AddRange(printers.ToArray());
                
                Logger.Info($"下拉框已更新，项目数: {cmbPrinter.Items.Count}");
                AddLogMessage($"打印机下拉列表已更新，包含 {cmbPrinter.Items.Count} 个项目");
                
                // 恢复保存的打印机选择
                var config = ConfigurationManager.Config;
                if (!string.IsNullOrEmpty(config.Printer.PrinterName))
                {
                    Logger.Info($"尝试选择配置中的打印机: {config.Printer.PrinterName}");
                    var printerIndex = cmbPrinter.FindString(config.Printer.PrinterName);
                    if (printerIndex >= 0)
                    {
                        cmbPrinter.SelectedIndex = printerIndex;
                        lblPrinterStatus.Text = LanguageManager.GetString("PrinterStatusOK");
                        lblPrinterStatus.ForeColor = Color.Green;
                        Logger.Info($"成功选择打印机: {config.Printer.PrinterName}，索引: {printerIndex}");
                        AddLogMessage($"已选择打印机: {config.Printer.PrinterName}");
                    }
                    else
                    {
                        lblPrinterStatus.Text = LanguageManager.GetString("PrinterStatusError");
                        lblPrinterStatus.ForeColor = Color.Red;
                        Logger.Warning($"配置中的打印机未找到: {config.Printer.PrinterName}");
                        AddLogMessage($"错误: 配置中的打印机未找到: {config.Printer.PrinterName}");
                    }
                }
                else
                {
                    lblPrinterStatus.Text = LanguageManager.GetString("PrinterStatus");
                    lblPrinterStatus.ForeColor = Color.Gray;
                    Logger.Info("配置中没有设置默认打印机");
                    AddLogMessage("提示: 请选择一台打印机");
                }
                
                Logger.Info($"打印机列表更新完成，共发现 {printers.Count} 台打印机");
            }
            catch (Exception ex)
            {
                Logger.Error($"更新打印机列表失败: {ex.Message}", ex);
                AddLogMessage($"错误: 更新打印机列表失败: {ex.Message}");
                lblPrinterStatus.Text = LanguageManager.GetString("GetPrinterListFailed");
                lblPrinterStatus.ForeColor = Color.Red;
            }
        }

        private void UpdateStatusDisplay()
        {
            if (_databaseMonitor.IsMonitoring)
            {
                lblMonitoringStatus.Text = LanguageManager.GetString("MonitoringStatusRunning");
                lblMonitoringStatus.ForeColor = Color.Green;
            }
            else
            {
                lblMonitoringStatus.Text = LanguageManager.GetString("MonitoringStatusStopped");
                lblMonitoringStatus.ForeColor = Color.Red;
            }
            
            lblTotalRecords.Text = $"{LanguageManager.GetString("ProcessedRecords")}: {_totalRecordsProcessed}";
            lblTotalPrints.Text = $"{LanguageManager.GetString("PrintJobs")}: {_totalPrintJobs}";
            
            // 显示最后记录的序列号
            try
            {
                var lastRecord = _databaseMonitor.GetLastRecord();
                if (lastRecord != null)
                {
                    lblLastRecord.Text = $"{LanguageManager.GetString("LastRecord")}: {lastRecord.TR_SerialNum ?? "N/A"}";
                }
                else
                {
                    lblLastRecord.Text = $"{LanguageManager.GetString("LastRecord")}: N/A";
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"获取最后记录失败: {ex.Message}", ex);
                lblLastRecord.Text = $"{LanguageManager.GetString("LastRecord")}: N/A";
            }
        }

        private void OnNewRecordFound(object? sender, TestRecord record)
        {
            Logger.Info($"🔔 新记录事件触发: TR_ID={record.TR_ID}, SerialNum={record.TR_SerialNum}");
            
            // 使用Invoke确保在UI线程上执行
            this.Invoke(new Action(() =>
            {
                try
                {
                    _totalRecordsProcessed++;
                    
                    AddLogMessage($"🎯 新记录 #{_totalRecordsProcessed}: {record.TR_SerialNum} (ID: {record.TR_ID})");
                    AddLogMessage("🔄 检测到数据变动，执行完整刷新...");
                    
                    // 🔧 新增功能：执行完整的刷新流程（类似 btnRefresh_Click）
                    try
                    {
                        // 1. 触发数据库监控的强制检查（如果需要）
                        if (_databaseMonitor.IsMonitoring)
                        {
                            AddLogMessage("🔍 触发数据库强制检查...");
                        }
                        
                        // 2. 强制刷新数据库连接以获取最新数据
                        _databaseMonitor.ForceRefreshConnection();
                        AddLogMessage("🔄 强制刷新数据库连接以获取最新数据");
                        
                        // 3. 刷新记录列表（完整加载）
                        LoadRecentRecords();
                        AddLogMessage("📋 记录列表已完整刷新");
                        
                        // 4. 🔧 新增：确保第一行高亮显示并滚动到可见位置
                        if (lvRecords.Items.Count > 0)
                        {
                            // 清除所有选择和高亮
                            lvRecords.SelectedItems.Clear();
                            foreach (ListViewItem item in lvRecords.Items)
                            {
                                item.BackColor = Color.White; // 重置背景色
                            }
                            
                            // 选中并高亮第一行
                            var firstItem = lvRecords.Items[0];
                            firstItem.Selected = true;
                            firstItem.Focused = true;
                            firstItem.BackColor = Color.LightYellow; // 淡黄色高亮显示新记录
                            firstItem.EnsureVisible(); // 确保滚动到可见位置
                            
                            AddLogMessage("🌟 第一行记录已高亮显示并滚动到可见位置");
                        }
                        
                        // 5. 更新状态显示
                        UpdateStatusDisplay();
                        AddLogMessage("📊 状态显示已更新");
                        
                        AddLogMessage("✅ 完整刷新流程执行完成");
                    }
                    catch (Exception refreshEx)
                    {
                        Logger.Error($"完整刷新流程失败: {refreshEx.Message}", refreshEx);
                        AddLogMessage($"❌ 完整刷新失败: {refreshEx.Message}");
                    }

                    // 🔧 修复重复打印问题：由于v1.3.9.0使用统一监控系统，禁用旧系统的自动打印
                    // 自动打印现在由OnDataUpdated方法统一处理，避免重复打印
                    /* 
                    // 自动打印新记录 - 已移至统一监控系统
                    try
                    {
                        AddLogMessage($"🖨️ 开始自动打印: {record.TR_SerialNum}");
                        AutoPrintRecord(record);
                        AddLogMessage($"✅ 自动打印完成: {record.TR_SerialNum}");
                    }
                    catch (Exception printEx)
                    {
                        Logger.Error($"自动打印失败: {printEx.Message}", printEx);
                        AddLogMessage($"❌ 自动打印失败: {printEx.Message}");
                    }
                    */
                    
                    AddLogMessage("🖨️ 自动打印由统一监控系统处理，避免重复打印");
                    
                    // 显示通知
                    ShowNotification($"新记录检测 #{_totalRecordsProcessed}", $"序列号: {record.TR_SerialNum} 已自动处理并高亮显示");
                    
                    Logger.Info($"✅ 新记录处理完成，总处理数: {_totalRecordsProcessed}");
                }
                catch (Exception ex)
                {
                    Logger.Error($"处理新记录失败: {ex.Message}", ex);
                    AddLogMessage($"❌ 处理新记录时出错: {ex.Message}");
                }
            }));
        }

        private void AutoPrintRecord(TestRecord record)
        {
            try
            {
                Logger.Info($"🖨️ 开始自动打印记录: {record.TR_SerialNum}");
                AddLogMessage($"🖨️ 准备打印: {record.TR_SerialNum}");
                
                // 更新当前打印信息
                UpdateCurrentPrintInfo(record, "监控检测到新记录-自动打印");

                var config = ConfigurationManager.Config;
                
                // 检查打印机配置
                if (string.IsNullOrEmpty(config.Printer.PrinterName))
                {
                    AddLogMessage("⚠️ 未选择打印机，尝试使用默认打印机");
                }
                
                // 使用默认模板或配置的模板
                var templateName = config.Printer.DefaultTemplate;
                AddLogMessage($"📄 使用打印模板: {templateName}");
                AddLogMessage($"🔧 打印格式: {config.Printer.PrintFormat}");
                
                var printResult = _printerService.PrintRecord(record, config.Printer.PrintFormat, templateName);

                if (printResult.Success)
                {
                    _totalPrintJobs++;
                    
                    // 更新数据库打印计数（使用TestRecord对象，优先TR_ID匹配）
                    try
                    {
                        _databaseMonitor.UpdatePrintCount(record);
                        AddLogMessage($"📊 已更新打印计数: {record.TR_SerialNum}");
                    }
                    catch (Exception countEx)
                    {
                        Logger.Warning($"更新打印计数失败: {countEx.Message}");
                        AddLogMessage($"⚠️ 更新打印计数失败: {countEx.Message}");
                    }
                    
                    Logger.Info($"✅ 自动打印完成: {record.TR_SerialNum}");
                    AddLogMessage($"✅ 打印成功: {record.TR_SerialNum} -> {printResult.PrinterUsed}");
                    AddLogMessage($"📈 总打印任务数: {_totalPrintJobs}");
                }
                else
                {
                    Logger.Error($"❌ 自动打印失败: {printResult.ErrorMessage}");
                    AddLogMessage($"❌ 打印失败: {record.TR_SerialNum}");
                    AddLogMessage($"📝 失败原因: {printResult.ErrorMessage}");
                    
                    // 如果是因为没有打印机导致的失败，显示详细提示
                    if (printResult.ErrorMessage?.Contains("打印机") == true)
                    {
                        AddLogMessage("💡 提示: 请检查打印机是否正确安装和配置");
                        if (!_printerService.HasAnyPrinter())
                        {
                            AddLogMessage("⚠️ 系统中未检测到可用的打印机");
                        }
                    }
                }

                // 打印完成后清除当前打印信息
                UpdateCurrentPrintInfo();
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 自动打印异常: {ex.Message}", ex);
                AddLogMessage($"❌ 打印异常: {record.TR_SerialNum}");
                AddLogMessage($"📝 异常详情: {ex.Message}");
                
                // 出错时也清除当前打印信息
                UpdateCurrentPrintInfo();
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
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
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
            // 加载配置并初始化UI
            LoadConfiguration();
            UpdatePrinterList();
            UpdateStatusDisplay();
            
            // 初始化打印次数列显示状态
            UpdatePrintCountColumnVisibility();
            
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
            
            // 检查自动启动监控配置
            CheckAutoStartMonitoring();
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

        private void CheckAutoStartMonitoring()
        {
            try
            {
                var config = ConfigurationManager.Config;
                
                // 如果启用了自动开始监控
                if (config.Application.AutoStartMonitoring)
                {
                    // 检查是否配置了数据库路径
                    if (string.IsNullOrEmpty(config.Database.DatabasePath))
                    {
                        AddLogMessage("⚠️ 数据库路径未配置，请在【配置】选项卡中设置数据库路径");
                        return;
                    }
                    
                    // 检查数据库文件是否存在
                    if (!System.IO.File.Exists(config.Database.DatabasePath))
                    {
                        AddLogMessage($"⚠️ 数据库文件不存在: {config.Database.DatabasePath}");
                        return;
                    }
                    
                    // 延迟1秒后启动监控，确保UI完全初始化
                    var autoStartTimer = new System.Windows.Forms.Timer();
                    autoStartTimer.Interval = 1000; // 1秒延迟
                    autoStartTimer.Tick += (sender, e) =>
                    {
                        autoStartTimer.Stop();
                        autoStartTimer.Dispose();
                        StartMonitoringDirectly();
                    };
                    autoStartTimer.Start();
                    
                    AddLogMessage("🚀 自动启动监控已启用，1秒后开始监控");
                }
                else
                {
                    AddLogMessage("自动启动监控已禁用");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"检查自动启动监控配置失败: {ex.Message}", ex);
                AddLogMessage($"❌ 检查自动启动监控配置失败: {ex.Message}");
            }
        }

        // 直接启动监控，无弹窗 - 使用异步连接
        private async void StartMonitoringDirectly()
        {
            try
            {
                var config = ConfigurationManager.Config.Database;
                
                AddLogMessage($"🔗 正在异步连接数据库: {config.DatabasePath}");
                
                // 使用异步连接方法（借鉴AccessDatabaseMonitor）
                if (!await _databaseMonitor.ConnectAsync(config.DatabasePath, config.TableName))
                {
                    AddLogMessage("❌ 数据库连接失败");
                    return;
                }
                
                AddLogMessage("✅ 数据库连接成功");
                
                // 获取表字段信息（静默）
                var columns = _databaseMonitor.GetTableColumns(config.TableName);
                AddLogMessage($"📊 检测到 {columns.Count} 个字段: {string.Join(", ", columns.Take(10))}{(columns.Count > 10 ? "..." : "")}");
                
                // 开始监控
                _databaseMonitor.StartMonitoring(config.PollInterval);
                AddLogMessage("🚀 数据库监控已启动");
                
                // 更新UI状态
                UpdateStatusDisplay();
                UpdateMonitoringButtonStates(true);
                
                // 立即加载一次数据
                LoadRecentRecords();
                
                Logger.Info("Direct monitoring started successfully");
            }
            catch (Exception ex)
            {
                Logger.Error($"直接启动监控失败: {ex.Message}", ex);
                AddLogMessage($"❌ 启动监控失败: {ex.Message}");
            }
        }

        private string RunMonitoringDiagnostic(DatabaseConfig config)
        {
            var diagnostics = new List<string>();
            
            try
            {
                // 检查数据库路径
                if (string.IsNullOrEmpty(config.DatabasePath))
                {
                    diagnostics.Add("❌ 数据库路径未配置");
                }
                else if (!File.Exists(config.DatabasePath))
                {
                    diagnostics.Add($"❌ 数据库文件不存在: {config.DatabasePath}");
                }
                else
                {
                    diagnostics.Add($"✅ 数据库文件存在: {Path.GetFileName(config.DatabasePath)}");
                }
                
                // 检查表名配置
                if (string.IsNullOrEmpty(config.TableName))
                {
                    diagnostics.Add("❌ 表名未配置");
                }
                else
                {
                    diagnostics.Add($"✅ 监控表: {config.TableName}");
                }
                
                // 检查监控字段配置
                if (string.IsNullOrEmpty(config.MonitorField))
                {
                    diagnostics.Add("❌ 监控字段未配置");
                }
                else
                {
                    diagnostics.Add($"✅ 监控字段: {config.MonitorField}");
                }
                
                // 检查轮询间隔
                if (config.PollInterval < 100)
                {
                    diagnostics.Add($"⚠️  轮询间隔过短: {config.PollInterval}ms (建议 >= 1000ms)");
                }
                else
                {
                    diagnostics.Add($"✅ 轮询间隔: {config.PollInterval}ms");
                }
                
                return string.Join(" | ", diagnostics);
            }
            catch (Exception ex)
            {
                Logger.Error($"监控诊断失败: {ex.Message}", ex);
                return $"诊断失败: {ex.Message}";
            }
        }

        private void OnFormResize(object? sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized && chkMinimizeToTray.Checked)
            {
                this.Hide();
                _notifyIcon.Visible = true;
                ShowNotification(LanguageManager.GetString("TrayNotificationTitle"), LanguageManager.GetString("TrayNotificationMessage"));
            }
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing && chkMinimizeToTray.Checked)
            {
                e.Cancel = true;
                this.Hide();
                _notifyIcon.Visible = true;
                ShowNotification(LanguageManager.GetString("TrayNotificationTitle"), LanguageManager.GetString("TrayNotificationMessage"));
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

        /// <summary>
        /// 🔧 统一监控系统：简化的记录加载方法
        /// 不再直接查询数据库，依赖统一监控系统的数据更新
        /// </summary>
        private async void LoadRecentRecords()
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

                // 🔧 统一监控系统：不再直接查询数据库
                // 如果监控正在运行，等待统一监控系统的数据更新
                if (_databaseMonitor.IsMonitoring)
                {
                    AddLogMessage("📋 统一监控运行中，等待下次数据更新...");
                    Logger.Info("LoadRecentRecords: 统一监控运行中，依赖DataUpdated事件");
                    return;
                }
                
                // 🔧 仅在监控未启动时，手动获取初始数据
                AddLogMessage("📋 监控未启动，手动获取初始数据...");
                Logger.Info("LoadRecentRecords: 监控未启动，手动获取初始50条记录");
                
                var records = _databaseMonitor.GetRecentRecords(50);
                UpdateRecordsList(records);
                
                AddLogMessage($"📊 手动加载完成，共 {records.Count} 条记录");
                
                // 更新状态显示
                UpdateStatusDisplay();
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 加载记录失败: {ex.Message}", ex);
                AddLogMessage($"❌ 加载记录失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 强制刷新最近记录 - 解决数据库同步问题（异步版本）
        /// </summary>
        private async Task ForceRefreshRecentRecords()
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

                // 🔧 修复：不要重启监控，只刷新数据显示
                // 原来的代码会重启监控，导致已知记录基线被重置，破坏监控连续性
                
                // 直接获取最新记录用于显示刷新
                var records = _databaseMonitor.GetRecentRecords(50);
                
                AddLogMessage($"🔍 强制刷新获取到 {records.Count} 条记录");
                
                lvRecords.Items.Clear();
                
                foreach (var record in records)
                {
                    var item = new ListViewItem(record.TR_SerialNum ?? "N/A");                          // 序列号
                    item.SubItems.Add(record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");   // 测试时间
                    item.SubItems.Add(record.FormatNumber(record.TR_Isc));                              // ISC
                    item.SubItems.Add(record.FormatNumber(record.TR_Voc));                              // VOC
                    item.SubItems.Add(record.FormatNumber(record.TR_Pm));                               // Pm
                    item.SubItems.Add(record.FormatNumber(record.TR_Ipm));                              // Ipm
                    item.SubItems.Add(record.FormatNumber(record.TR_Vpm));                              // Vpm
                    item.SubItems.Add((record.TR_Print ?? 0).ToString());                               // 打印次数
                    item.SubItems.Add("双击打印");                                                       // 操作
                    item.SubItems.Add(record.TR_ID ?? "N/A");                                           // 记录ID
                    item.Tag = record;
                    
                    // 根据打印次数设置颜色（仅在启用打印次数统计时）
                    var printConfig = ConfigurationManager.Config;
                    if (printConfig.Database.EnablePrintCount && record.TR_Print > 0)
                    {
                        item.ForeColor = Color.Gray;  // 已打印的记录显示为灰色
                    }
                    else
                    {
                        item.ForeColor = Color.Black; // 未打印的记录显示为黑色
                    }
                    
                    lvRecords.Items.Add(item);
                }
                
                AddLogMessage($"✅ 强制刷新完成：已加载 {records.Count} 条最近记录");
                Logger.Info($"强制刷新完成：已加载 {records.Count} 条最近记录");
            }
            catch (Exception ex)
            {
                Logger.Error($"强制刷新最近记录失败: {ex.Message}", ex);
                AddLogMessage($"❌ 强制刷新失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 强制刷新最近记录 - 同步版本，用于UI线程调用
        /// </summary>
        private void SyncForceRefreshRecentRecords()
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

                // 🔧 修复：不要重启监控，只刷新数据显示
                // 原来的代码会重启监控，导致已知记录基线被重置，破坏监控连续性
                
                // 直接获取最新记录用于显示刷新
                var records = _databaseMonitor.GetRecentRecords(50);
                
                AddLogMessage($"🔍 强制刷新获取到 {records.Count} 条记录");
                
                lvRecords.Items.Clear();
                
                foreach (var record in records)
                {
                    var item = new ListViewItem(record.TR_SerialNum ?? "N/A");                          // 序列号
                    item.SubItems.Add(record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");   // 测试时间
                    item.SubItems.Add(record.FormatNumber(record.TR_Isc));                              // ISC
                    item.SubItems.Add(record.FormatNumber(record.TR_Voc));                              // VOC
                    item.SubItems.Add(record.FormatNumber(record.TR_Pm));                               // Pm
                    item.SubItems.Add(record.FormatNumber(record.TR_Ipm));                              // Ipm
                    item.SubItems.Add(record.FormatNumber(record.TR_Vpm));                              // Vpm
                    item.SubItems.Add((record.TR_Print ?? 0).ToString());                               // 打印次数
                    item.SubItems.Add("双击打印");                                                       // 操作
                    item.SubItems.Add(record.TR_ID ?? "N/A");                                           // 记录ID
                    item.Tag = record;
                    
                    // 根据打印次数设置颜色（仅在启用打印次数统计时）
                    var printConfig = ConfigurationManager.Config;
                    if (printConfig.Database.EnablePrintCount && record.TR_Print > 0)
                    {
                        item.ForeColor = Color.Gray;  // 已打印的记录显示为灰色
                    }
                    else
                    {
                        item.ForeColor = Color.Black; // 未打印的记录显示为黑色
                    }
                    
                    lvRecords.Items.Add(item);
                }
                
                AddLogMessage($"✅ 强制刷新完成：已加载 {records.Count} 条最近记录");
                Logger.Info($"强制刷新完成：已加载 {records.Count} 条最近记录");
            }
            catch (Exception ex)
            {
                Logger.Error($"强制刷新最近记录失败: {ex.Message}", ex);
                AddLogMessage($"❌ 强制刷新失败: {ex.Message}");
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
                    numFontSize.Value = template.FontSize;
                    cmbFontName.SelectedItem = template.FontName;
                    
                    // 加载页眉页脚设置
                    _showHeader = template.ShowHeader;
                    _headerText = template.HeaderText;
                    _headerImagePath = template.HeaderImagePath;
                    _showFooter = template.ShowFooter;
                    _footerText = template.FooterText;
                    _footerImagePath = template.FooterImagePath;
                    
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
                    IsDefault = false,
                    FontSize = (int)numFontSize.Value,
                    FontName = cmbFontName.SelectedItem?.ToString() ?? "Arial",
                    ShowHeader = _showHeader,
                    HeaderText = _headerText,
                    HeaderImagePath = _headerImagePath,
                    ShowFooter = _showFooter,
                    FooterText = _footerText,
                    FooterImagePath = _footerImagePath
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
                    Format = Enum.Parse<PrintFormat>(cmbTemplateFormat.SelectedItem?.ToString() ?? "Text"),
                    FontSize = (int)numFontSize.Value,
                    FontName = cmbFontName.SelectedItem?.ToString() ?? "Arial",
                    ShowHeader = _showHeader,
                    HeaderText = _headerText,
                    HeaderImagePath = _headerImagePath,
                    ShowFooter = _showFooter,
                    FooterText = _footerText,
                    FooterImagePath = _footerImagePath
                };

                var preview = PrintTemplateManager.ProcessTemplate(template, record);
                rtbTemplatePreview.Text = preview;
                
                // 设置预览的字体大小和名称
                rtbTemplatePreview.Font = new Font(template.FontName, template.FontSize);
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

        private void btnHeaderFooterSettings_Click(object? sender, EventArgs e)
        {
            // 创建包含当前页眉页脚设置的模板对象
            var currentTemplate = new PrintTemplate
            {
                Name = txtTemplateName.Text,
                Content = txtTemplateContent.Text,
                Format = Enum.Parse<PrintFormat>(cmbTemplateFormat.SelectedItem?.ToString() ?? "Text"),
                FontSize = (int)numFontSize.Value,
                FontName = cmbFontName.SelectedItem?.ToString() ?? "Arial",
                ShowHeader = _showHeader,
                HeaderText = _headerText,
                HeaderImagePath = _headerImagePath,
                ShowFooter = _showFooter,
                FooterText = _footerText,
                FooterImagePath = _footerImagePath
            };

            using var headerFooterForm = new HeaderFooterSettingsForm(currentTemplate);
            if (headerFooterForm.ShowDialog() == DialogResult.OK)
            {
                // 更新主窗体中的页眉页脚设置
                _showHeader = currentTemplate.ShowHeader;
                _headerText = currentTemplate.HeaderText;
                _headerImagePath = currentTemplate.HeaderImagePath;
                _showFooter = currentTemplate.ShowFooter;
                _footerText = currentTemplate.FooterText;
                _footerImagePath = currentTemplate.FooterImagePath;

                MessageBox.Show("页眉页脚设置已更新，请保存模板以应用更改。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                AddLogMessage($"页眉页脚设置已更新");
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
            this.Text = $"{LanguageManager.GetString("MainTitle")} v1.2.7 - 数据库查询简化修复版";
            
            // 更新选项卡标题
            if (tabControl1.TabPages.Count >= 4)
            {
                tabControl1.TabPages[0].Text = LanguageManager.GetString("TabDataMonitoring");
                tabControl1.TabPages[1].Text = LanguageManager.GetString("TabSystemConfig");
                tabControl1.TabPages[2].Text = LanguageManager.GetString("TabPrintTemplate");
                tabControl1.TabPages[3].Text = LanguageManager.GetString("TabRuntimeLogs");
            }

            // 更新ListView列标题
            if (lvRecords.Columns.Count >= 10)
            {
                lvRecords.Columns[0].Text = LanguageManager.GetString("SerialNumber");      // 序列号
                lvRecords.Columns[1].Text = LanguageManager.GetString("TestDateTime");      // 测试时间
                lvRecords.Columns[2].Text = LanguageManager.GetString("Current");           // ISC
                lvRecords.Columns[3].Text = LanguageManager.GetString("Voltage");           // VOC
                lvRecords.Columns[4].Text = LanguageManager.GetString("Power");             // Pm
                lvRecords.Columns[5].Text = LanguageManager.GetString("CurrentIpm");        // Ipm
                lvRecords.Columns[6].Text = LanguageManager.GetString("VoltageVpm");        // Vpm
                lvRecords.Columns[7].Text = LanguageManager.GetString("PrintCount");        // 打印次数
                lvRecords.Columns[8].Text = LanguageManager.GetString("Operation");         // 操作
                lvRecords.Columns[9].Text = LanguageManager.GetString("RecordID");          // 记录ID
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
                    // 直接使用选中记录的Tag中存储的TestRecord对象
                    var selectedItem = lvRecords.SelectedItems[0];
                    recordToPreview = selectedItem.Tag as TestRecord;
                    
                    if (recordToPreview != null)
                    {
                        Logger.Info($"使用选中的记录进行预览: TR_ID={recordToPreview.TR_ID}, 序列号={recordToPreview.TR_SerialNum}");
                    }
                    else
                    {
                        // 如果Tag为空，使用TR_ID（主键）查找记录
                        var recordId = selectedItem.SubItems.Count > 8 ? selectedItem.SubItems[8].Text : null;
                        
                        if (!string.IsNullOrEmpty(recordId) && recordId != "N/A")
                        {
                            var records = _databaseMonitor.GetRecentRecords(100);
                            recordToPreview = records.FirstOrDefault(r => r.TR_ID == recordId);
                            Logger.Warning($"ListView项Tag为空，通过主键TR_ID查找记录: {recordId}");
                        }
                        else
                        {
                            Logger.Error("无法获取TR_ID，无法定位记录");
                        }
                    }
                }
                else if (lvRecords.Items.Count > 0)
                {
                    // 使用最新记录
                    var latestItem = lvRecords.Items[0];
                    recordToPreview = latestItem.Tag as TestRecord;
                    
                    if (recordToPreview != null)
                    {
                        Logger.Info($"使用最新记录进行预览: TR_ID={recordToPreview.TR_ID}, 序列号={recordToPreview.TR_SerialNum}");
                    }
                    else
                    {
                        // 如果Tag为空，使用TR_ID（主键）查找最新记录
                        var recordId = latestItem.SubItems.Count > 8 ? latestItem.SubItems[8].Text : null;
                        
                        if (!string.IsNullOrEmpty(recordId) && recordId != "N/A")
                        {
                            var records = _databaseMonitor.GetRecentRecords(100);
                            recordToPreview = records.FirstOrDefault(r => r.TR_ID == recordId);
                            Logger.Warning($"ListView项Tag为空，通过主键TR_ID查找最新记录: {recordId}");
                        }
                        else
                        {
                            Logger.Error("无法获取最新记录的TR_ID，无法定位记录");
                        }
                    }
                }
                
                if (recordToPreview == null)
                {
                    MessageBox.Show("没有可预览的数据，请先进行测试或选择一条记录。", 
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // 创建或更新预览窗口
                if (_printPreviewForm == null || _printPreviewForm.IsDisposed)
                {
                    _printPreviewForm = new PrintPreviewForm(recordToPreview, _printerService);
                    _printPreviewForm.PrintRequested += OnPrintPreviewRequested;
                    _printPreviewForm.Show(this);
                    Logger.Info("创建新的打印预览窗口");
                }
                else
                {
                    // 窗口已存在，更新记录数据
                    _printPreviewForm.LoadRecord(recordToPreview);
                    _printPreviewForm.SetAutoPrintMode(chkAutoPrint.Checked);
                    
                    if (!_printPreviewForm.Visible)
                    {
                        _printPreviewForm.Show(this);
                    }
                    else
                    {
                        _printPreviewForm.BringToFront();
                    }
                    Logger.Info("更新现有打印预览窗口的数据");
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

        private void OnPrintPreviewRequested(object? sender, EventArgs e)
        {
            // 处理打印预览请求
            if (lvRecords.SelectedItems.Count > 0)
            {
                var selectedItem = lvRecords.SelectedItems[0];
                var record = selectedItem.Tag as TestRecord;
                
                if (record != null)
                {
                    // 执行打印
                    btnManualPrint_Click(null, EventArgs.Empty);
                }
            }
        }
        
        /// <summary>
        /// 🔧 修复打印预览窗口和弹窗冲突问题：弹出模态对话框前的处理
        /// </summary>
        private void HandlePreviewFormBeforeDialog()
        {
            try
            {
                if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && _printPreviewForm.Visible)
                {
                    // 临时将打印预览窗口设置为不可见，避免焦点冲突
                    _printPreviewForm.Visible = false;
                    Logger.Info("临时隐藏打印预览窗口以避免模态对话框冲突");
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"处理打印预览窗口焦点时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 🔧 修复打印预览窗口和弹窗冲突问题：模态对话框关闭后的处理
        /// </summary>
        private void HandlePreviewFormAfterDialog()
        {
            try
            {
                if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && !_printPreviewForm.Visible)
                {
                    // 恢复打印预览窗口的显示
                    _printPreviewForm.Show(this);
                    Logger.Info("恢复打印预览窗口显示");
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"恢复打印预览窗口显示时出错: {ex.Message}");
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

        private void InitializeServices()
        {
            _databaseMonitor = new DatabaseMonitor();
            _printerService = new PrinterService();
            Logger.Info("服务初始化完成");
        }

        /// <summary>
        /// 初始化数据库监控服务
        /// </summary>
        private void InitializeDatabaseMonitor()
        {
            try
            {
                _databaseMonitor = new DatabaseMonitor();
                
                // 🔧 统一监控系统：订阅统一数据更新事件
                _databaseMonitor.DataUpdated += OnDataUpdated;
                
                // 保持兼容性事件订阅
                _databaseMonitor.NewRecordFound += OnNewRecordFound;
                _databaseMonitor.StatusChanged += OnStatusChanged;
                _databaseMonitor.MonitoringError += OnMonitoringError;
                
                Logger.Info("✅ 数据库监控服务初始化完成 - 统一监控系统");
                AddLogMessage("✅ 数据库监控服务初始化完成 - 基于GetLastRecord的统一监控");
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 数据库监控服务初始化失败: {ex.Message}", ex);
                AddLogMessage($"❌ 数据库监控服务初始化失败: {ex.Message}");
            }
        }

        // 旧的InitializeNotifyIcon方法已被SetupNotifyIcon替代
        // 该方法使用更完整的图标加载逻辑和错误处理

        private void InitializeTimer()
        {
            _statusUpdateTimer = new System.Windows.Forms.Timer
            {
                Interval = 1000, // 1秒更新一次状态
                Enabled = true
            };
            _statusUpdateTimer.Tick += (s, e) => UpdateStatusDisplay();
        }

        private void OnAutoPrintChanged(object? sender, EventArgs e)
        {
            try
            {
                var config = ConfigurationManager.Config;
                config.Printer.AutoPrint = chkAutoPrint.Checked;
                ConfigurationManager.SaveConfig();
                
                var status = chkAutoPrint.Checked ? LanguageManager.GetString("Enabled") : LanguageManager.GetString("Disabled");
                AddLogMessage($"{LanguageManager.GetString("AutoPrint")} {status}");
                
                // 如果预览窗口已打开，同步更新按钮状态
                if (_printPreviewForm != null && !_printPreviewForm.IsDisposed && _printPreviewForm.Visible)
                {
                    _printPreviewForm.SetAutoPrintMode(chkAutoPrint.Checked);
                    Logger.Info($"同步更新预览窗口自动打印状态: {chkAutoPrint.Checked}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"更新自动打印状态失败: {ex.Message}", ex);
            }
        }



        // 预印刷标签相关方法已删除

        private void lvRecords_DoubleClick(object? sender, EventArgs e)
        {
            btnManualPrint_Click(sender, e);
        }

        // 预印刷标签相关字段控制方法已删除

        private T? FindControlByName<T>(Control parent, string name) where T : Control
        {
            foreach (Control control in parent.Controls)
            {
                if (control is T typedControl && (control.Name == name || control.GetType().Name.Contains(name.Replace("cmb", "ComboBox").Replace("num", "NumericUpDown").Replace("chk", "CheckBox").Replace("pnl", "Panel"))))
                {
                    return typedControl;
                }
                
                var found = FindControlByName<T>(control, name);
                if (found != null) return found;
            }
            return null;
        }

        private string GetFieldSampleValue(string fieldName)
        {
            return fieldName switch
            {
                "{SerialNumber}" => "ABC123456",
                "{TestDateTime}" => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                "{ShortCircuitCurrent}" => "10.25",
                "{OpenCircuitVoltage}" => "24.5",
                "{MaxPowerVoltage}" => "20.8",
                "{MaxPower}" => "250.5",
                "{MaxPowerCurrent}" => "12.04",
                "{PrintCount}" => "1",
                "{CurrentTime}" => DateTime.Now.ToString("HH:mm:ss"),
                "{CurrentDate}" => DateTime.Now.ToString("yyyy-MM-dd"),
                _ => "示例值"
            };
        }

        private string GetFieldDisplayName(string fieldName)
        {
            var descriptions = PrintTemplateManager.GetFieldDescriptions();
            return descriptions.TryGetValue(fieldName, out var description) ? description : fieldName;
        }

        private ContentAlignment GetContentAlignment(string alignment)
        {
            return alignment switch
            {
                "Right" => ContentAlignment.MiddleRight,
                "Center" => ContentAlignment.MiddleCenter,
                _ => ContentAlignment.MiddleLeft
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

        private void UpdatePrintCountColumnVisibility()
        {
            var printCountColumnIndex = 7; // 打印次数列的索引
            var config = ConfigurationManager.Config;
            
            if (config.Database.EnablePrintCount)
            {
                // 显示打印次数列
                if (this.lvRecords.Columns.Count > printCountColumnIndex)
                {
                    this.lvRecords.Columns[printCountColumnIndex].Width = 100;
                }
            }
            else
            {
                // 隐藏打印次数列
                if (this.lvRecords.Columns.Count > printCountColumnIndex)
                {
                    this.lvRecords.Columns[printCountColumnIndex].Width = 0;
                }
            }
        }

        private void UpdateCurrentPrintInfo(TestRecord? record = null, string source = "")
        {
            try
            {
                if (record != null)
                {
                    var printInfo = $"当前打印: 序列号 {record.TR_SerialNum ?? "N/A"} (来源: {source})";
                    if (lblCurrentPrint != null)
                    {
                        lblCurrentPrint.Text = printInfo;
                        lblCurrentPrint.ForeColor = Color.Green;
                    }
                    Logger.Info($"打印监控: {printInfo}");
                }
                else
                {
                    var printInfo = LanguageManager.GetString("CurrentPrintInfo");
                    if (lblCurrentPrint != null)
                    {
                        lblCurrentPrint.Text = printInfo;
                        lblCurrentPrint.ForeColor = Color.Blue;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"更新当前打印信息失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 🔧 统一监控系统：处理统一数据更新事件
        /// 基于GetLastRecord监控，一次性接收最后记录和50条记录列表
        /// </summary>
        private void OnDataUpdated(object? sender, DataUpdateEventArgs e)
        {
            // 确保在UI线程上执行
            this.Invoke(new Action(() =>
            {
                try
                {
                    Logger.Info($"📋 统一数据更新事件: {e.UpdateType} - {e.ChangeDetails}");
                    AddLogMessage($"📋 统一数据更新: {e.UpdateType} - {e.LastRecord.TR_SerialNum}");
                    AddLogMessage($"📊 接收到 {e.RecentRecords.Count} 条最新记录");
                    
                    // 🔧 核心：基于统一监控数据，直接更新UI列表
                    UpdateRecordsList(e.RecentRecords, e.LastRecord);
                    
                    // 更新状态显示
                    UpdateStatusDisplay();
                    
                    // 🔧 新增：如果是记录更新（非初始化），执行自动打印
                    if (e.UpdateType == "记录更新")
                    {
                        // 检查是否启用了自动打印功能
                        if (chkAutoPrint.Checked)
                        {
                            AddLogMessage($"🖨️ 开始自动打印: {e.LastRecord.TR_SerialNum}");
                            try
                            {
                                AutoPrintRecord(e.LastRecord);
                                AddLogMessage($"✅ 自动打印完成: {e.LastRecord.TR_SerialNum}");
                            }
                            catch (Exception printEx)
                            {
                                Logger.Error($"自动打印失败: {printEx.Message}", printEx);
                                AddLogMessage($"❌ 自动打印失败: {printEx.Message}");
                            }
                        }
                        else
                        {
                            AddLogMessage($"⏸️ 自动打印已禁用，跳过打印: {e.LastRecord.TR_SerialNum}");
                        }
                        
                        // 显示通知
                        ShowNotification($"新记录检测", $"序列号: {e.LastRecord.TR_SerialNum} 已自动处理并高亮显示");
                    }
                    
                    Logger.Info($"✅ 统一数据更新处理完成");
                }
                catch (Exception ex)
                {
                    Logger.Error($"❌ 统一数据更新处理失败: {ex.Message}", ex);
                    AddLogMessage($"❌ 数据更新处理失败: {ex.Message}");
                }
            }));
        }
        
        /// <summary>
        /// 🔧 核心：基于统一监控数据更新记录列表
        /// </summary>
        private void UpdateRecordsList(List<TestRecord> records, TestRecord? highlightRecord = null)
        {
            try
            {
                lvRecords.Items.Clear();
                
                foreach (var record in records)
                {
                    var item = new ListViewItem(record.TR_SerialNum ?? "N/A");                          // 序列号
                    item.SubItems.Add(record.TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");   // 测试时间
                    item.SubItems.Add(record.FormatNumber(record.TR_Isc));                              // ISC
                    item.SubItems.Add(record.FormatNumber(record.TR_Voc));                              // VOC
                    item.SubItems.Add(record.FormatNumber(record.TR_Pm));                               // Pm
                    item.SubItems.Add(record.FormatNumber(record.TR_Ipm));                              // Ipm
                    item.SubItems.Add(record.FormatNumber(record.TR_Vpm));                              // Vpm
                    item.SubItems.Add((record.TR_Print ?? 0).ToString());                               // 打印次数
                    item.SubItems.Add("双击打印");                                                       // 操作
                    item.SubItems.Add(record.TR_ID ?? "N/A");                                           // 记录ID
                    item.Tag = record;
                    
                    // 根据打印次数设置颜色（仅在启用打印次数统计时）
                    var printConfig = ConfigurationManager.Config;
                    if (printConfig.Database.EnablePrintCount && record.TR_Print > 0)
                    {
                        item.BackColor = Color.LightGray; // 已打印记录显示灰色
                    }
                    
                    lvRecords.Items.Add(item);
                }
                
                // 🔧 高亮最新记录（如果指定）
                if (highlightRecord != null && lvRecords.Items.Count > 0)
                {
                    // 清除所有选择和高亮
                    lvRecords.SelectedItems.Clear();
                    foreach (ListViewItem item in lvRecords.Items)
                    {
                        if (item.BackColor != Color.LightGray) // 保持已打印记录的灰色
                        {
                            item.BackColor = Color.White;
                        }
                    }
                    
                    // 查找并高亮匹配的记录
                    foreach (ListViewItem item in lvRecords.Items)
                    {
                        if (item.Tag is TestRecord record && 
                            record.TR_SerialNum == highlightRecord.TR_SerialNum)
                        {
                            item.Selected = true;
                            item.Focused = true;
                            item.BackColor = Color.LightYellow; // 淡黄色高亮显示新记录
                            item.EnsureVisible(); // 确保滚动到可见位置
                            
                            AddLogMessage("🌟 新记录已高亮显示并滚动到可见位置");
                            break;
                        }
                    }
                }
                
                AddLogMessage($"📊 记录列表已更新，共 {records.Count} 条记录");
                
                // 显示最后记录的序列号
                if (records.Count > 0)
                {
                    var lastRecord = records[0]; // 第一条是最新的
                    lblLastRecord.Text = $"{LanguageManager.GetString("LastRecord")}: {lastRecord.TR_SerialNum ?? "N/A"}";
                }
                else
                {
                    lblLastRecord.Text = $"{LanguageManager.GetString("LastRecord")}: N/A";
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"❌ 更新记录列表失败: {ex.Message}", ex);
                AddLogMessage($"❌ 更新记录列表失败: {ex.Message}");
            }
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