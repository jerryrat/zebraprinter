using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Forms
{
    public partial class PrintPreviewForm : Form
    {
        private readonly TestRecord _record;
        private readonly PrinterService _printerService;
        private System.Windows.Forms.Timer? _blinkTimer;
        private int _blinkCount;
        private Color _originalColor;
        
        public PrintPreviewForm(TestRecord record, PrinterService printerService)
        {
            _record = record;
            _printerService = printerService;
            
            InitializeComponent();
            InitializeBlinkTimer();
            LoadPreviewData();
            
            // 添加窗体关闭事件处理
            this.FormClosing += OnFormClosing;
        }
        
        private void InitializeBlinkTimer()
        {
            _blinkTimer = new System.Windows.Forms.Timer();
            _blinkTimer.Interval = 200; // 200ms间隔闪烁
            _blinkTimer.Tick += BlinkTimer_Tick;
            _originalColor = Color.Green; // 默认绿色
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
        
        private void StartBlinkEffect()
        {
            if (_blinkTimer != null)
            {
                _blinkTimer.Stop();
                _blinkCount = 0;
                _blinkTimer.Start();
            }
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
        }
        
        private void LoadPreviewData()
        {
            try
            {
                if (_record != null)
                {
                    // 启动闪烁效果
                    StartBlinkEffect();
                    
                    lblSerialNumber.Text = _record.TR_SerialNum ?? LanguageManager.GetString("NA");
                    
                    // 加载并显示打印内容
                    var config = ConfigurationManager.Config;
                    var templateName = config.Printer.DefaultTemplate;
                    var content = _printerService.GeneratePrintContent(_record, templateName);
                    
                    if (!string.IsNullOrEmpty(content))
                    {
                        rtbPreviewContent.Text = content;
                    }
                    else
                    {
                        rtbPreviewContent.Text = LanguageManager.GetString("NoPreviewData");
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
                }
                else
                {
                    lblSerialNumber.Text = LanguageManager.GetString("NA");
                    rtbPreviewContent.Text = LanguageManager.GetString("NoPreviewData");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{LanguageManager.GetString("LoadPreviewError")}: {ex.Message}", 
                    LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void btnConfirmPrint_Click(object sender, EventArgs e)
        {
            try
            {
                if (_record != null)
                {
                    var config = ConfigurationManager.Config;
                    var templateName = config.Printer.DefaultTemplate;
                    var printResult = _printerService.PrintRecord(_record, config.Printer.PrintFormat, templateName);
                    
                    if (printResult.Success)
                    {
                        MessageBox.Show(LanguageManager.GetString("PrintCompleted"), 
                            LanguageManager.GetString("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show($"{LanguageManager.GetString("PrintFailed")}: {printResult.ErrorMessage}", 
                            LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{LanguageManager.GetString("PrintError")}: {ex.Message}", 
                    LanguageManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void btnShowMain_Click(object sender, EventArgs e)
        {
            // 显示主窗口
            var mainForm = Application.OpenForms["MainForm"];
            if (mainForm != null)
            {
                mainForm.Show();
                mainForm.WindowState = FormWindowState.Normal;
                mainForm.BringToFront();
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
            // 启动闪烁效果提示新数据加载
            StartBlinkEffect();
            
            // 由于_record是readonly，我们需要重新创建预览内容
            RefreshPreview();
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
        

    }
} 