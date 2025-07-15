using System;
using System.Drawing;
using System.Windows.Forms;

namespace ZebraPrinterMonitor.Forms
{
    partial class PrintPreviewForm
    {
        private System.ComponentModel.IContainer components = null;
        private Panel pnlMain;
        private Label lblSerialNumber;
        private Panel pnlContent;
        private RichTextBox rtbPreviewContent;
        private Panel pnlButtons;
        private Button btnConfirmPrint;
        private Button btnShowMain;
        private Button btnClose;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.pnlMain = new Panel();
            this.lblSerialNumber = new Label();
            this.pnlContent = new Panel();
            this.rtbPreviewContent = new RichTextBox();
            this.pnlButtons = new Panel();
            this.btnConfirmPrint = new Button();
            this.btnShowMain = new Button();
            this.btnClose = new Button();
            
            this.pnlMain.SuspendLayout();
            this.pnlContent.SuspendLayout();
            this.pnlButtons.SuspendLayout();
            this.SuspendLayout();

            // 主窗体设置 - 紧凑型手机屏幕大小
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(380, 550); // 减少高度，优化尺寸
            this.Controls.Add(this.pnlMain);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrintPreviewForm";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "🖨️ 打印预览";
            this.TopMost = true;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // 主面板
            this.pnlMain.Dock = DockStyle.Fill;
            this.pnlMain.Location = new Point(0, 0);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Size = new Size(380, 550);
            this.pnlMain.TabIndex = 0;
            this.pnlMain.Padding = new Padding(12);
            this.pnlMain.BackColor = Color.White;

            // 序列号标签 - 现代化设计
            this.lblSerialNumber.AutoSize = false;
            this.lblSerialNumber.Dock = DockStyle.Top;
            this.lblSerialNumber.Font = new Font("Segoe UI", 24F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.lblSerialNumber.ForeColor = Color.White;
            this.lblSerialNumber.Location = new Point(12, 12);
            this.lblSerialNumber.Name = "lblSerialNumber";
            this.lblSerialNumber.Size = new Size(356, 60);
            this.lblSerialNumber.TabIndex = 0;
            this.lblSerialNumber.Text = "加载中...";
            this.lblSerialNumber.TextAlign = ContentAlignment.MiddleCenter;
            this.lblSerialNumber.BackColor = Color.FromArgb(52, 152, 219);
            this.lblSerialNumber.FlatStyle = FlatStyle.Flat;

            // 内容面板
            this.pnlContent.Dock = DockStyle.Fill;
            this.pnlContent.Location = new Point(12, 72);
            this.pnlContent.Name = "pnlContent";
            this.pnlContent.Size = new Size(356, 393);
            this.pnlContent.TabIndex = 1;
            this.pnlContent.Padding = new Padding(0, 8, 0, 8);

            // 预览内容文本框 - 简化和美化
            this.rtbPreviewContent.Dock = DockStyle.Fill;
            this.rtbPreviewContent.Font = new Font("Consolas", 11F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.rtbPreviewContent.Location = new Point(0, 8);
            this.rtbPreviewContent.Name = "rtbPreviewContent";
            this.rtbPreviewContent.ReadOnly = true;
            this.rtbPreviewContent.Size = new Size(356, 377);
            this.rtbPreviewContent.TabIndex = 0;
            this.rtbPreviewContent.Text = "正在加载打印内容...";
            this.rtbPreviewContent.BackColor = Color.FromArgb(248, 249, 250);
            this.rtbPreviewContent.BorderStyle = BorderStyle.None;
            this.rtbPreviewContent.Margin = new Padding(4);

            // 按钮面板 - 现代化设计
            this.pnlButtons.Dock = DockStyle.Bottom;
            this.pnlButtons.Height = 85;
            this.pnlButtons.Location = new Point(12, 465);
            this.pnlButtons.Name = "pnlButtons";
            this.pnlButtons.Size = new Size(356, 85);
            this.pnlButtons.TabIndex = 2;
            this.pnlButtons.Padding = new Padding(8);
            this.pnlButtons.BackColor = Color.White;

            // 确认打印按钮 - 现代化设计
            this.btnConfirmPrint.BackColor = Color.FromArgb(46, 204, 113);
            this.btnConfirmPrint.FlatStyle = FlatStyle.Flat;
            this.btnConfirmPrint.FlatAppearance.BorderSize = 0;
            this.btnConfirmPrint.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.btnConfirmPrint.ForeColor = Color.White;
            this.btnConfirmPrint.Location = new Point(8, 8);
            this.btnConfirmPrint.Name = "btnConfirmPrint";
            this.btnConfirmPrint.Size = new Size(340, 40);
            this.btnConfirmPrint.TabIndex = 0;
            this.btnConfirmPrint.Text = "🖨️ 确认打印";
            this.btnConfirmPrint.UseVisualStyleBackColor = false;
            this.btnConfirmPrint.Click += new EventHandler(this.btnConfirmPrint_Click);

            // 显示主界面按钮
            this.btnShowMain.BackColor = Color.FromArgb(52, 152, 219);
            this.btnShowMain.FlatStyle = FlatStyle.Flat;
            this.btnShowMain.FlatAppearance.BorderSize = 0;
            this.btnShowMain.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.btnShowMain.ForeColor = Color.White;
            this.btnShowMain.Location = new Point(8, 54);
            this.btnShowMain.Name = "btnShowMain";
            this.btnShowMain.Size = new Size(250, 25);
            this.btnShowMain.TabIndex = 1;
            this.btnShowMain.Text = "📋 显示主界面";
            this.btnShowMain.UseVisualStyleBackColor = false;
            this.btnShowMain.Click += new EventHandler(this.btnShowMain_Click);

            // 关闭按钮 - 精简设计
            this.btnClose.BackColor = Color.FromArgb(149, 165, 166);
            this.btnClose.FlatStyle = FlatStyle.Flat;
            this.btnClose.FlatAppearance.BorderSize = 0;
            this.btnClose.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.btnClose.ForeColor = Color.White;
            this.btnClose.Location = new Point(264, 54);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new Size(84, 25);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "✖ 关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new EventHandler(this.btnClose_Click);

            // 添加控件到面板
            this.pnlMain.Controls.Add(this.pnlContent);
            this.pnlMain.Controls.Add(this.lblSerialNumber);
            this.pnlMain.Controls.Add(this.pnlButtons);
            this.pnlContent.Controls.Add(this.rtbPreviewContent);
            this.pnlButtons.Controls.Add(this.btnConfirmPrint);
            this.pnlButtons.Controls.Add(this.btnShowMain);
            this.pnlButtons.Controls.Add(this.btnClose);

            this.pnlMain.ResumeLayout(false);
            this.pnlContent.ResumeLayout(false);
            this.pnlButtons.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
} 