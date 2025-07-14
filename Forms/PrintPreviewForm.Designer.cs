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
            this.btnClose = new Button();
            
            this.pnlMain.SuspendLayout();
            this.pnlContent.SuspendLayout();
            this.pnlButtons.SuspendLayout();
            this.SuspendLayout();

            // 主窗体设置 - 手机屏幕大小 (约iPhone Plus尺寸)
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(414, 736); // iPhone Plus 尺寸
            this.Controls.Add(this.pnlMain);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrintPreviewForm";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "打印预览";
            this.TopMost = true;
            this.BackColor = Color.White;

            // 主面板
            this.pnlMain.Dock = DockStyle.Fill;
            this.pnlMain.Location = new Point(0, 0);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Size = new Size(414, 736);
            this.pnlMain.TabIndex = 0;
            this.pnlMain.Padding = new Padding(15);

            // 序列号标签 - 顶端加粗加大
            this.lblSerialNumber.AutoSize = false;
            this.lblSerialNumber.Dock = DockStyle.Top;
            this.lblSerialNumber.Font = new Font("Microsoft Sans Serif", 24F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.lblSerialNumber.ForeColor = Color.DarkBlue;
            this.lblSerialNumber.Location = new Point(15, 15);
            this.lblSerialNumber.Name = "lblSerialNumber";
            this.lblSerialNumber.Size = new Size(384, 80);
            this.lblSerialNumber.TabIndex = 0;
            this.lblSerialNumber.Text = "序列号: 加载中...";
            this.lblSerialNumber.TextAlign = ContentAlignment.MiddleCenter;
            this.lblSerialNumber.BackColor = Color.LightBlue;
            this.lblSerialNumber.BorderStyle = BorderStyle.FixedSingle;

            // 内容面板
            this.pnlContent.Dock = DockStyle.Fill;
            this.pnlContent.Location = new Point(15, 95);
            this.pnlContent.Name = "pnlContent";
            this.pnlContent.Size = new Size(384, 541);
            this.pnlContent.TabIndex = 1;
            this.pnlContent.Padding = new Padding(0, 10, 0, 10);

            // 预览内容文本框
            this.rtbPreviewContent.Dock = DockStyle.Fill;
            this.rtbPreviewContent.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.rtbPreviewContent.Location = new Point(0, 10);
            this.rtbPreviewContent.Name = "rtbPreviewContent";
            this.rtbPreviewContent.ReadOnly = true;
            this.rtbPreviewContent.Size = new Size(384, 521);
            this.rtbPreviewContent.TabIndex = 0;
            this.rtbPreviewContent.Text = "正在加载打印内容...";
            this.rtbPreviewContent.BackColor = Color.White;
            this.rtbPreviewContent.BorderStyle = BorderStyle.FixedSingle;

            // 按钮面板
            this.pnlButtons.Dock = DockStyle.Bottom;
            this.pnlButtons.Height = 85;
            this.pnlButtons.Location = new Point(15, 636);
            this.pnlButtons.Name = "pnlButtons";
            this.pnlButtons.Size = new Size(384, 85);
            this.pnlButtons.TabIndex = 2;
            this.pnlButtons.Padding = new Padding(10);

            // 确认打印按钮
            this.btnConfirmPrint.BackColor = Color.FromArgb(0, 122, 204);
            this.btnConfirmPrint.FlatStyle = FlatStyle.Flat;
            this.btnConfirmPrint.Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.btnConfirmPrint.ForeColor = Color.White;
            this.btnConfirmPrint.Location = new Point(10, 10);
            this.btnConfirmPrint.Name = "btnConfirmPrint";
            this.btnConfirmPrint.Size = new Size(280, 45);
            this.btnConfirmPrint.TabIndex = 0;
            this.btnConfirmPrint.Text = "确认打印";
            this.btnConfirmPrint.UseVisualStyleBackColor = false;
            this.btnConfirmPrint.Click += new EventHandler(this.btnConfirmPrint_Click);

            // 关闭按钮
            this.btnClose.BackColor = Color.Gray;
            this.btnClose.FlatStyle = FlatStyle.Flat;
            this.btnClose.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.btnClose.ForeColor = Color.White;
            this.btnClose.Location = new Point(300, 10);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new Size(74, 45);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new EventHandler(this.btnClose_Click);

            // 添加控件到面板
            this.pnlMain.Controls.Add(this.pnlContent);
            this.pnlMain.Controls.Add(this.lblSerialNumber);
            this.pnlMain.Controls.Add(this.pnlButtons);
            this.pnlContent.Controls.Add(this.rtbPreviewContent);
            this.pnlButtons.Controls.Add(this.btnConfirmPrint);
            this.pnlButtons.Controls.Add(this.btnClose);

            this.pnlMain.ResumeLayout(false);
            this.pnlContent.ResumeLayout(false);
            this.pnlButtons.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
} 