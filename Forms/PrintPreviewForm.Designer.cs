using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Services;

namespace ZebraPrinterMonitor.Forms
{
    partial class PrintPreviewForm
    {
        private System.ComponentModel.IContainer components = null;
        private Label lblSerialNumber;
        private RichTextBox rtbPreviewContent;
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
            this.lblSerialNumber = new Label();
            this.rtbPreviewContent = new RichTextBox();
            this.btnConfirmPrint = new Button();
            this.btnShowMain = new Button();
            this.btnClose = new Button();
            this.SuspendLayout();
            
            // 
            // PrintPreviewForm
            // 
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(340, 550);
            this.Controls.Add(this.lblSerialNumber);
            this.Controls.Add(this.rtbPreviewContent);
            this.Controls.Add(this.btnConfirmPrint);
            this.Controls.Add(this.btnShowMain);
            this.Controls.Add(this.btnClose);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrintPreviewForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = LanguageManager.GetString("PrintPreviewTitle");
            // 🔧 设置窗口置顶显示在所有程序窗口最上层
            this.TopMost = true;                     // 置于所有窗口最顶层
            this.ShowInTaskbar = true;               // 在任务栏显示，方便用户管理
            this.ShowIcon = false;
            
            // 
            // lblSerialNumber
            // 
            this.lblSerialNumber.AutoSize = false;
            this.lblSerialNumber.Font = new Font("Microsoft Sans Serif", 24F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.lblSerialNumber.ForeColor = Color.Green;
            this.lblSerialNumber.Location = new Point(20, 15);
            this.lblSerialNumber.Name = "lblSerialNumber";
            this.lblSerialNumber.Size = new Size(300, 35);
            this.lblSerialNumber.TabIndex = 0;
            this.lblSerialNumber.Text = LanguageManager.GetString("Loading");
            this.lblSerialNumber.TextAlign = ContentAlignment.MiddleCenter;
            
            // 
            // rtbPreviewContent
            // 
            this.rtbPreviewContent.Font = new Font("Arial", 14F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.rtbPreviewContent.Location = new Point(20, 60);
            this.rtbPreviewContent.Name = "rtbPreviewContent";
            this.rtbPreviewContent.ReadOnly = true;
            this.rtbPreviewContent.Size = new Size(300, 420);
            this.rtbPreviewContent.TabIndex = 1;
            this.rtbPreviewContent.Text = LanguageManager.GetString("LoadingContent");
            // 🔧 修复打印预览文字遮挡问题：添加换行和滚动条支持
            this.rtbPreviewContent.WordWrap = true;                    // 启用自动换行
            this.rtbPreviewContent.ScrollBars = RichTextBoxScrollBars.Both; // 添加滚动条
            this.rtbPreviewContent.DetectUrls = false;                 // 禁用URL检测
            this.rtbPreviewContent.Multiline = true;                   // 确保多行显示
            
            // 
            // btnConfirmPrint
            // 
            this.btnConfirmPrint.BackColor = Color.FromArgb(0, 122, 204);
            this.btnConfirmPrint.FlatAppearance.BorderSize = 0;
            this.btnConfirmPrint.FlatStyle = FlatStyle.Flat;
            this.btnConfirmPrint.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.btnConfirmPrint.ForeColor = Color.White;
            this.btnConfirmPrint.Location = new Point(20, 500);
            this.btnConfirmPrint.Name = "btnConfirmPrint";
            this.btnConfirmPrint.Size = new Size(110, 35);
            this.btnConfirmPrint.TabIndex = 2;
            this.btnConfirmPrint.Text = LanguageManager.GetString("ConfirmPrint");
            this.btnConfirmPrint.UseVisualStyleBackColor = false;
            this.btnConfirmPrint.Click += new EventHandler(this.btnConfirmPrint_Click);
            
            // 
            // btnShowMain
            // 
            this.btnShowMain.BackColor = Color.FromArgb(40, 167, 69);
            this.btnShowMain.FlatAppearance.BorderSize = 0;
            this.btnShowMain.FlatStyle = FlatStyle.Flat;
            this.btnShowMain.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.btnShowMain.ForeColor = Color.White;
            this.btnShowMain.Location = new Point(145, 500);
            this.btnShowMain.Name = "btnShowMain";
            this.btnShowMain.Size = new Size(90, 35);
            this.btnShowMain.TabIndex = 3;
            this.btnShowMain.Text = LanguageManager.GetString("ShowMainWindow");
            this.btnShowMain.UseVisualStyleBackColor = false;
            this.btnShowMain.Click += new EventHandler(this.btnShowMain_Click);
            
            // 
            // btnClose
            // 
            this.btnClose.BackColor = Color.FromArgb(220, 53, 69);
            this.btnClose.FlatAppearance.BorderSize = 0;
            this.btnClose.FlatStyle = FlatStyle.Flat;
            this.btnClose.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.btnClose.ForeColor = Color.White;
            this.btnClose.Location = new Point(250, 500);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new Size(70, 35);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = LanguageManager.GetString("Close");
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new EventHandler(this.btnClose_Click);
            
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
} 