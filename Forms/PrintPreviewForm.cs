using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;
using System.Reflection;

namespace ZebraPrinterMonitor.Forms
{
    public partial class PrintPreviewForm : Form
    {
        private TestRecord _record; // 移除readonly，允许动态更新记录
        private readonly PrinterService _printerService;
        private System.Windows.Forms.Timer? _blinkTimer;
        private System.Windows.Forms.Timer? _titleBlinkTimer;
        private int _blinkCount;
        private Color _originalColor;
        private string _originalTitle;
        private bool _isPrinting = false;
        private bool _titleBlinkState = false;
        
        public PrintPreviewForm(TestRecord record, PrinterService printerService)
        {
            _record = record;
            _printerService = printerService;
            
            InitializeComponent();
            
            // 🔧 设置窗口置顶显示
            this.TopMost = true;
            this.ShowInTaskbar = true; // 在任务栏显示，方便用户管理
            
            InitializeBlinkTimer();
            InitializeTitleBlinkTimer();
            LoadPreviewData();
            
            // 保存原始标题
            _originalTitle = this.Text;
            
            // 添加窗体关闭事件处理
            this.FormClosing += OnFormClosing;
            
            // 初始化按钮状态
            UpdateMainWindowButton();
        }
        
        private void InitializeBlinkTimer()
        {
            _blinkTimer = new System.Windows.Forms.Timer();
            _blinkTimer.Interval = 200; // 200ms间隔闪烁
            _blinkTimer.Tick += BlinkTimer_Tick;
            _originalColor = Color.Green; // 默认绿色
        }
        
        /// <summary>
        /// 初始化标题闪烁定时器
        /// </summary>
        private void InitializeTitleBlinkTimer()
        {
            _titleBlinkTimer = new System.Windows.Forms.Timer();
            _titleBlinkTimer.Interval = 800; // 800ms间隔闪烁，相对较慢
            _titleBlinkTimer.Tick += TitleBlinkTimer_Tick;
        }
        
        private void BlinkTimer_Tick(object? sender, EventArgs e)
        {
            _blinkCount++;
            
            if (_blinkCount <= 6) // 闪烁3次（6个切换）
            {
                // 在深绿色和原始绿色之间切换
                lblSerialNumber.ForeColor = (_blinkCount % 2 == 1) ? Color.DarkGreen : _originalColor;
            }
            else
            {
                // 停止闪烁，恢复到正常绿色
                _blinkTimer?.Stop();
                lblSerialNumber.ForeColor = _originalColor;
                _blinkCount = 0;
            }
        }
        
        /// <summary>
        /// 标题闪烁定时器事件
        /// </summary>
        private void TitleBlinkTimer_Tick(object? sender, EventArgs e)
        {
            if (_isPrinting)
            {
                _titleBlinkState = !_titleBlinkState;
                if (_titleBlinkState)
                {
                    this.Text = _originalTitle + " - 🖨️ 打印中....";
                }
                else
                {
                    this.Text = _originalTitle + " - 打印中....";
                }
            }
        }
        
        private void StartBlinkEffect()
        {
            if (_blinkTimer != null)
            {
                _blinkTimer.Stop();
                _blinkCount = 0;
                _blinkTimer.Start();
            }
        }
        
        /// <summary>
        /// 开始打印状态显示和标题闪烁
        /// </summary>
        private void StartPrintingStatus()
        {
            _isPrinting = true;
            _titleBlinkTimer?.Start();
            Logger.Info("开始打印状态显示 - 标题闪烁已启动");
        }
        
        /// <summary>
        /// 停止打印状态显示和标题闪烁
        /// </summary>
        private void StopPrintingStatus()
        {
            _isPrinting = false;
            _titleBlinkTimer?.Stop();
            this.Text = _originalTitle; // 恢复原始标题
            Logger.Info("停止打印状态显示 - 标题已恢复");
        }
        
        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            // 清理定时器资源
            if (_blinkTimer != null)
            {
                _blinkTimer.Stop();
                _blinkTimer.Dispose();
                _blinkTimer = null;
            }
            
            if (_titleBlinkTimer != null)
            {
                _titleBlinkTimer.Stop();
                _titleBlinkTimer.Dispose();
                _titleBlinkTimer = null;
            }
        }
        
        /// <summary>
        /// 检测主窗口是否隐藏到系统托盘
        /// </summary>
        /// <returns>true表示主窗口隐藏，false表示主窗口显示</returns>
        private bool IsMainWindowHidden()
        {
            var mainForm = Application.OpenForms["MainForm"];
            if (mainForm != null)
            {
                // 检查窗口是否可见且未最小化
                return !mainForm.Visible || mainForm.WindowState == FormWindowState.Minimized;
            }
            return true; // 如果找不到主窗口，假设已隐藏
        }
        
        /// <summary>
        /// 更新主窗口控制按钮的文字和状态
        /// </summary>
        private void UpdateMainWindowButton()
        {
            try
            {
                if (IsMainWindowHidden())
                {
                    btnShowMain.Text = LanguageManager.GetString("ShowMainWindow");
                    btnShowMain.BackColor = Color.FromArgb(40, 167, 69); // 绿色 - 显示
                }
                else
                {
                    btnShowMain.Text = LanguageManager.GetString("HideMainWindow");
                    btnShowMain.BackColor = Color.FromArgb(255, 193, 7); // 黄色 - 隐藏
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"更新主窗口按钮状态失败: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// 获取主窗口的NotifyIcon实例（通过反射）
        /// </summary>
        private NotifyIcon? GetMainFormNotifyIcon()
        {
            try
            {
                var mainForm = Application.OpenForms["MainForm"];
                if (mainForm != null)
                {
                    // 通过反射获取_notifyIcon字段
                    var notifyIconField = mainForm.GetType().GetField("_notifyIcon", 
                        BindingFlags.NonPublic | BindingFlags.Instance);
                    return notifyIconField?.GetValue(mainForm) as NotifyIcon;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"获取主窗口NotifyIcon失败: {ex.Message}", ex);
            }
            return null;
        }
        
        private void LoadPreviewData()
        {
            try
            {
                if (_record != null)
                {
                    Logger.Info($"加载预览数据: 序列号={_record.TR_SerialNum}, 时间={_record.TR_DateTime}");
                    
                    // 启动闪烁效果
                    StartBlinkEffect();
                    
                    lblSerialNumber.Text = _record.TR_SerialNum ?? LanguageManager.GetString("NA");
                    
                    // 加载并显示打印内容
                    var config = ConfigurationManager.Config;
                    var templateName = config.Printer.DefaultTemplate;
                    
                    Logger.Info($"生成打印内容: 模板={templateName}, 格式={config.Printer.PrintFormat}");
                    var content = _printerService.GeneratePrintContent(_record, templateName);
                    
                    if (!string.IsNullOrEmpty(content))
                    {
                        rtbPreviewContent.Text = content;
                        Logger.Info($"预览内容已生成，长度: {content.Length} 字符");
                    }
                    else
                    {
                        rtbPreviewContent.Text = LanguageManager.GetString("NoPreviewData");
                        Logger.Warning("生成的打印内容为空");
                    }
                    
                    // 检查是否启用了自动打印
                    if (config.Printer.AutoPrint)
                    {
                        btnConfirmPrint.Text = LanguageManager.GetString("AutoPrintEnabled");
                        btnConfirmPrint.Enabled = false;
                    }
                    else
                    {
                        btnConfirmPrint.Text = LanguageManager.GetString("ConfirmPrint");
                        btnConfirmPrint.Enabled = true;
                    }
                    
                    // 显示记录的关键信息用于调试
                    Logger.Info($"预览数据详情: Isc={_record.TR_Isc}, Voc={_record.TR_Voc}, Pm={_record.TR_Pm}");
                }
                else
                {
                    lblSerialNumber.Text = LanguageManager.GetString("NA");
                    rtbPreviewContent.Text = LanguageManager.GetString("NoPreviewData");
                    Logger.Warning("预览记录为null，无法加载数据");
                }
                
                // 更新主窗口按钮状态
                UpdateMainWindowButton();
            }
            catch (Exception ex)
            {
                Logger.Error($"加载预览数据失败: {ex.Message}", ex);
                // 🔧 修复弹窗冲突问题：确保不受TopMost影响
                var originalTopMost = this.TopMost;
                this.TopMost = false;
                MessageBox.Show($"{LanguageManager.GetString("LoadPreviewError")}: {ex.Message}", 
                    LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.TopMost = originalTopMost;
            }
        }
        
        private void btnConfirmPrint_Click(object sender, EventArgs e)
        {
            try
            {
                if (_record != null)
                {
                    // 🔧 开始打印状态显示
                    StartPrintingStatus();
                    
                    var config = ConfigurationManager.Config;
                    var templateName = config.Printer.DefaultTemplate;
                    var printResult = _printerService.PrintRecord(_record, config.Printer.PrintFormat, templateName);
                    
                    // 🔧 停止打印状态显示
                    StopPrintingStatus();
                    
                    if (printResult.Success)
                    {
                        // 🔧 修复弹窗冲突问题：确保不受TopMost影响
                        var originalTopMost = this.TopMost;
                        this.TopMost = false;
                        MessageBox.Show(LanguageManager.GetString("PrintCompleted"), 
                            LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.TopMost = originalTopMost;
                        this.Close();
                    }
                    else
                    {
                        // 🔧 修复弹窗冲突问题：确保不受TopMost影响
                        var originalTopMost = this.TopMost;
                        this.TopMost = false;
                        MessageBox.Show($"{LanguageManager.GetString("PrintFailed")}: {printResult.ErrorMessage}", 
                            LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.TopMost = originalTopMost;
                    }
                }
            }
            catch (Exception ex)
            {
                // 🔧 停止打印状态显示
                StopPrintingStatus();
                
                // 🔧 修复弹窗冲突问题：确保不受TopMost影响
                var originalTopMost = this.TopMost;
                this.TopMost = false;
                MessageBox.Show($"{LanguageManager.GetString("PrintError")}: {ex.Message}", 
                    LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.TopMost = originalTopMost;
            }
        }
        
        private void btnShowMain_Click(object sender, EventArgs e)
        {
            try
            {
                var mainForm = Application.OpenForms["MainForm"];
                if (mainForm != null)
                {
                    if (IsMainWindowHidden())
                    {
                        // 主窗口当前隐藏，显示它
                        mainForm.Show();
                        mainForm.WindowState = FormWindowState.Normal;
                        mainForm.BringToFront();
                        
                        // 🔧 隐藏系统托盘图标
                        var notifyIcon = GetMainFormNotifyIcon();
                        if (notifyIcon != null)
                        {
                            notifyIcon.Visible = false;
                        }
                        
                        Logger.Info("主窗口已从系统托盘显示");
                    }
                    else
                    {
                        // 主窗口当前显示，隐藏到系统托盘
                        mainForm.Hide();
                        
                        // 🔧 显示系统托盘图标
                        var notifyIcon = GetMainFormNotifyIcon();
                        if (notifyIcon != null)
                        {
                            notifyIcon.Visible = true;
                            
                            // 显示托盘通知
                            try
                            {
                                notifyIcon.ShowBalloonTip(3000, 
                                    LanguageManager.GetString("TrayNotificationTitle"), 
                                    LanguageManager.GetString("TrayNotificationMessage"), 
                                    ToolTipIcon.Info);
                            }
                            catch
                            {
                                // 忽略通知显示错误
                            }
                        }
                        
                        Logger.Info("主窗口已隐藏到系统托盘");
                    }
                    
                    // 更新按钮状态
                    UpdateMainWindowButton();
                }
                else
                {
                    Logger.Warning("未找到主窗口实例");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"切换主窗口显示状态失败: {ex.Message}", ex);
                var originalTopMost = this.TopMost;
                this.TopMost = false;
                MessageBox.Show($"切换主窗口状态失败: {ex.Message}", 
                    LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.TopMost = originalTopMost;
            }
        }
        
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void RefreshPreview()
        {
            LoadPreviewData();
        }

        public event EventHandler? PrintRequested;

        public void LoadRecord(TestRecord record)
        {
            // 更新当前记录
            _record = record;
            
            // 启动闪烁效果提示新数据加载
            StartBlinkEffect();
            
            // 重新加载预览内容
            RefreshPreview();
            
            Logger.Info($"打印预览窗口已更新记录: {record.TR_SerialNum}");
        }

        public void SetAutoPrintMode(bool enabled)
        {
            if (enabled)
            {
                btnConfirmPrint.Text = LanguageManager.GetString("AutoPrintEnabled");
                btnConfirmPrint.Enabled = false;
            }
            else
            {
                btnConfirmPrint.Text = LanguageManager.GetString("ConfirmPrint");
                btnConfirmPrint.Enabled = true;
            }
        }
        
        /// <summary>
        /// 公共方法：刷新主窗口按钮状态（供外部调用）
        /// </summary>
        public void RefreshMainWindowButton()
        {
            UpdateMainWindowButton();
        }

    }
} 