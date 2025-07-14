using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Forms
{
    public partial class PrintMonitorForm : Form
    {
        private readonly PrinterService _printerService;
        private TestRecord? _currentRecord;
        
        // UI Controls
        private Label lblSerialNumber;
        private Label lblStatus;
        private RichTextBox rtbPreview;
        private Button btnConfirmPrint;
        private Button btnShowMain;
        private Button btnClose;
        private Panel pnlHeader;
        private Panel pnlContent;
        private Panel pnlButtons;

        public PrintMonitorForm(PrinterService printerService)
        {
            _printerService = printerService;
            InitializeComponent();
            InitializeUI();
            UpdateLanguage();
        }

        public event EventHandler? ShowMainRequested;
        public event EventHandler<TestRecord>? PrintRequested;

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // 设置窗体基本属性 - 手机屏幕大小 (360 x 640)
            this.Size = new Size(380, 680);
            this.MinimumSize = new Size(350, 600);
            this.MaximumSize = new Size(400, 800);
            this.StartPosition = FormStartPosition.Manual;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.ShowInTaskbar = true;
            this.TopMost = false;
            this.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.BackColor = Color.White;
            
            // 创建头部面板
            this.pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 120,
                BackColor = Color.FromArgb(240, 248, 255),
                Padding = new Padding(10)
            };
            
            // 序列号标签 - 大号加粗
            this.lblSerialNumber = new Label
            {
                Text = "暂无记录",
                Font = new Font("Microsoft YaHei UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 122, 183),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.Transparent,
                AutoEllipsis = true
            };
            
            // 状态标签
            this.lblStatus = new Label
            {
                Text = "等待数据...",
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Bottom,
                Height = 30,
                BackColor = Color.Transparent
            };
            
            this.pnlHeader.Controls.Add(this.lblSerialNumber);
            this.pnlHeader.Controls.Add(this.lblStatus);
            
            // 创建内容面板
            this.pnlContent = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.White
            };
            
            // 预览文本框
            this.rtbPreview = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle,
                Text = "打印内容预览区域\n等待新的测试记录..."
            };
            
            this.pnlContent.Controls.Add(this.rtbPreview);
            
            // 创建按钮面板
            this.pnlButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 100,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(10)
            };
            
            // 确认打印按钮
            this.btnConfirmPrint = new Button
            {
                Text = "确认打印",
                Size = new Size(340, 35),
                Location = new Point(10, 10),
                BackColor = Color.FromArgb(92, 184, 92),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                Enabled = false
            };
            this.btnConfirmPrint.FlatAppearance.BorderSize = 0;
            this.btnConfirmPrint.Click += BtnConfirmPrint_Click;
            
            // 显示主界面按钮
            this.btnShowMain = new Button
            {
                Text = "显示主界面",
                Size = new Size(110, 30),
                Location = new Point(10, 55),
                BackColor = Color.FromArgb(91, 192, 222),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9F)
            };
            this.btnShowMain.FlatAppearance.BorderSize = 0;
            this.btnShowMain.Click += BtnShowMain_Click;
            
            // 关闭按钮
            this.btnClose = new Button
            {
                Text = "关闭监控",
                Size = new Size(110, 30),
                Location = new Point(240, 55),
                BackColor = Color.FromArgb(217, 83, 79),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9F)
            };
            this.btnClose.FlatAppearance.BorderSize = 0;
            this.btnClose.Click += BtnClose_Click;
            
            this.pnlButtons.Controls.Add(this.btnConfirmPrint);
            this.pnlButtons.Controls.Add(this.btnShowMain);
            this.pnlButtons.Controls.Add(this.btnClose);
            
            // 添加到窗体
            this.Controls.Add(this.pnlContent);
            this.Controls.Add(this.pnlHeader);
            this.Controls.Add(this.pnlButtons);
            
            this.ResumeLayout(false);
        }

        private void InitializeUI()
        {
            // 设置窗体位置到屏幕右侧
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(screen.Right - this.Width - 20, screen.Top + 50);
            
            UpdateDisplay();
        }

        public void UpdateRecord(TestRecord record)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<TestRecord>(UpdateRecord), record);
                return;
            }

            _currentRecord = record;
            UpdateDisplay();
        }

        public void UpdateAutoPrintStatus(bool isAutoMode)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(UpdateAutoPrintStatus), isAutoMode);
                return;
            }

            btnConfirmPrint.Enabled = !isAutoMode && _currentRecord != null;
            
            var statusText = isAutoMode ? 
                LanguageManager.GetString("AutoPrintEnabled") : 
                LanguageManager.GetString("ManualPrintMode");
            
            if (_currentRecord == null)
            {
                statusText = LanguageManager.GetString("WaitingForData");
            }
            
            lblStatus.Text = statusText;
        }

        private void UpdateDisplay()
        {
            if (_currentRecord == null)
            {
                lblSerialNumber.Text = LanguageManager.GetString("NoRecordYet");
                lblStatus.Text = LanguageManager.GetString("WaitingForData");
                rtbPreview.Text = LanguageManager.GetString("WaitingForData") + "\n\n" + 
                                  LanguageManager.GetString("PrintPreview");
                btnConfirmPrint.Enabled = false;
                return;
            }

            // 更新序列号显示
            lblSerialNumber.Text = _currentRecord.TR_SerialNum ?? "N/A";
            
            // 生成预览内容
            var template = PrintTemplateManager.GetDefaultTemplate();
            var preview = PrintTemplateManager.ProcessTemplate(template, _currentRecord);
            rtbPreview.Text = preview;
            
            // 更新按钮状态
            var config = ConfigurationManager.Config;
            btnConfirmPrint.Enabled = !config.Printer.AutoPrint;
            
            UpdateAutoPrintStatus(config.Printer.AutoPrint);
        }

        private void UpdateLanguage()
        {
            this.Text = LanguageManager.GetString("PrintMonitorTitle");
            btnConfirmPrint.Text = LanguageManager.GetString("ConfirmPrint");
            btnShowMain.Text = LanguageManager.GetString("ShowMainWindow");
            btnClose.Text = LanguageManager.GetString("CloseMonitor");
            
            // 如果当前没有记录，更新默认显示文本
            if (_currentRecord == null)
            {
                lblSerialNumber.Text = LanguageManager.GetString("NoRecordYet");
                lblStatus.Text = LanguageManager.GetString("WaitingForData");
                rtbPreview.Text = LanguageManager.GetString("WaitingForData") + "\n\n" + 
                                  LanguageManager.GetString("PrintPreview");
            }
        }

        private void BtnConfirmPrint_Click(object? sender, EventArgs e)
        {
            if (_currentRecord == null) return;

            try
            {
                PrintRequested?.Invoke(this, _currentRecord);
                
                // 临时显示成功消息
                var originalText = lblStatus.Text;
                lblStatus.Text = LanguageManager.GetString("PrintSuccess");
                lblStatus.ForeColor = Color.Green;
                
                var timer = new System.Windows.Forms.Timer { Interval = 2000 };
                timer.Tick += (s, ev) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    lblStatus.Text = originalText;
                    lblStatus.ForeColor = Color.Gray;
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                Logger.Error($"打印确认失败: {ex.Message}", ex);
                lblStatus.Text = LanguageManager.GetString("PrintError");
                lblStatus.ForeColor = Color.Red;
            }
        }

        private void BtnShowMain_Click(object? sender, EventArgs e)
        {
            ShowMainRequested?.Invoke(this, EventArgs.Empty);
        }

        private void BtnClose_Click(object? sender, EventArgs e)
        {
            this.Hide();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 防止窗体被关闭，只是隐藏
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                base.OnFormClosing(e);
            }
        }

        // 支持语言切换时的更新
        public void RefreshLanguage()
        {
            UpdateLanguage();
            UpdateDisplay();
        }

        // 显示通知消息
        public void ShowNotification(string message, MessageType type = MessageType.Info)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string, MessageType>(ShowNotification), message, type);
                return;
            }

            var originalText = lblStatus.Text;
            var originalColor = lblStatus.ForeColor;
            
            lblStatus.Text = message;
            lblStatus.ForeColor = type switch
            {
                MessageType.Success => Color.Green,
                MessageType.Error => Color.Red,
                MessageType.Warning => Color.Orange,
                _ => Color.Blue
            };

            var timer = new System.Windows.Forms.Timer { Interval = 3000 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                lblStatus.Text = originalText;
                lblStatus.ForeColor = originalColor;
            };
            timer.Start();
        }
    }

    public enum MessageType
    {
        Info,
        Success,
        Warning,
        Error
    }
} 