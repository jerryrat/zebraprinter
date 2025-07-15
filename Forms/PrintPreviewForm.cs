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
        private TestRecord? _currentRecord;
        private readonly PrinterService _printerService;
        private bool _autoPrintEnabled;
        
        public event EventHandler<TestRecord>? PrintRequested;

        public PrintPreviewForm()
        {
            InitializeComponent();
            _printerService = new PrinterService();
            UpdateButtonStates();
        }

        public void LoadRecord(TestRecord record, bool autoPrintEnabled = false)
        {
            _currentRecord = record;
            _autoPrintEnabled = autoPrintEnabled;
            
            // 更新序列号显示 - 删除emoji，只显示序列号
            if (!string.IsNullOrEmpty(record.TR_SerialNum))
            {
                lblSerialNumber.Text = record.TR_SerialNum;
            }
            else
            {
                lblSerialNumber.Text = "N/A";
            }

            // 生成预览内容
            UpdatePreviewContent();
            
            // 更新按钮状态
            UpdateButtonStates();
        }

        private void UpdatePreviewContent()
        {
            if (_currentRecord == null)
            {
                rtbPreviewContent.Text = "没有可预览的数据";
                return;
            }

            try
            {
                var config = ConfigurationManager.Config;
                var templateName = config.Printer.DefaultTemplate;
                
                // 获取打印模板并处理
                var template = PrintTemplateManager.GetTemplate(templateName);
                if (template == null)
                {
                    template = PrintTemplateManager.GetDefaultTemplate();
                }

                var processedContent = PrintTemplateManager.ProcessTemplate(template, _currentRecord);
                
                // 简化显示内容 - 只显示核心打印内容
                rtbPreviewContent.Clear();
                
                // 直接显示处理后的打印内容，使用统一的字体
                rtbPreviewContent.SelectionFont = new Font("Consolas", 11F, FontStyle.Regular);
                rtbPreviewContent.SelectionColor = Color.FromArgb(33, 37, 41);
                rtbPreviewContent.AppendText(processedContent);
                
                // 滚动到顶部
                rtbPreviewContent.SelectionStart = 0;
                rtbPreviewContent.ScrollToCaret();
            }
            catch (Exception ex)
            {
                Logger.Error($"预览内容生成失败: {ex.Message}", ex);
                rtbPreviewContent.Text = $"预览内容生成失败:\n{ex.Message}";
            }
        }

        private void UpdateButtonStates()
        {
            // 当自动打印启用时，确认打印按钮失效
            btnConfirmPrint.Enabled = !_autoPrintEnabled && _currentRecord != null;
            
            if (_autoPrintEnabled)
            {
                btnConfirmPrint.Text = "自动打印已启用";
                btnConfirmPrint.BackColor = Color.Gray;
            }
            else
            {
                btnConfirmPrint.Text = "确认打印";
                btnConfirmPrint.BackColor = Color.FromArgb(0, 122, 204);
            }
        }

        private void btnConfirmPrint_Click(object? sender, EventArgs e)
        {
            if (_currentRecord == null || _autoPrintEnabled)
            {
                return;
            }

            try
            {
                // 触发打印事件
                PrintRequested?.Invoke(this, _currentRecord);
                
                // 显示打印确认
                MessageBox.Show("打印任务已发送", "打印确认", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // 关闭窗口
                this.Close();
            }
            catch (Exception ex)
            {
                Logger.Error($"确认打印失败: {ex.Message}", ex);
                MessageBox.Show($"打印失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClose_Click(object? sender, EventArgs e)
        {
            this.Close();
        }

        private void btnShowMain_Click(object? sender, EventArgs e)
        {
            try
            {
                // 显示主窗口并将其置于前台
                if (this.Owner != null)
                {
                    this.Owner.Show();
                    this.Owner.WindowState = FormWindowState.Normal;
                    this.Owner.BringToFront();
                    this.Owner.Activate();
                    
                    Logger.Info("主窗口已显示并置于前台");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"显示主窗口失败: {ex.Message}", ex);
            }
        }

        public void SetAutoPrintMode(bool enabled)
        {
            _autoPrintEnabled = enabled;
            UpdateButtonStates();
        }

        public void RefreshPreview()
        {
            if (_currentRecord != null)
            {
                UpdatePreviewContent();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            Logger.Info("打印预览窗口已关闭");
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Logger.Info("打印预览窗口已显示");
            
            // 确保窗口在屏幕中央
            this.CenterToParent();
        }
    }
} 