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
            this.ClientSize = new Size(800, 600);
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
            
            // 
            // lblSerialNumber
            // 
            this.lblSerialNumber.AutoSize = true;
            this.lblSerialNumber.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.lblSerialNumber.Location = new Point(20, 20);
            this.lblSerialNumber.Name = "lblSerialNumber";
            this.lblSerialNumber.Size = new Size(100, 20);
            this.lblSerialNumber.TabIndex = 0;
            this.lblSerialNumber.Text = LanguageManager.GetString("Loading");
            
            // 
            // rtbPreviewContent
            // 
            this.rtbPreviewContent.Font = new Font("Consolas", 10F, FontStyle.Regular, GraphicsUnit.Point, 134);
            this.rtbPreviewContent.Location = new Point(20, 50);
            this.rtbPreviewContent.Name = "rtbPreviewContent";
            this.rtbPreviewContent.ReadOnly = true;
            this.rtbPreviewContent.Size = new Size(760, 480);
            this.rtbPreviewContent.TabIndex = 1;
            this.rtbPreviewContent.Text = LanguageManager.GetString("LoadingContent");
            
            // 
            // btnConfirmPrint
            // 
            this.btnConfirmPrint.BackColor = Color.FromArgb(0, 122, 204);
            this.btnConfirmPrint.FlatAppearance.BorderSize = 0;
            this.btnConfirmPrint.FlatStyle = FlatStyle.Flat;
            this.btnConfirmPrint.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            this.btnConfirmPrint.ForeColor = Color.White;
            this.btnConfirmPrint.Location = new Point(20, 550);
            this.btnConfirmPrint.Name = "btnConfirmPrint";
            this.btnConfirmPrint.Size = new Size(150, 30);
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
            this.btnShowMain.Location = new Point(325, 550);
            this.btnShowMain.Name = "btnShowMain";
            this.btnShowMain.Size = new Size(150, 30);
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
            this.btnClose.Location = new Point(630, 550);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new Size(150, 30);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = LanguageManager.GetString("Close");
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new EventHandler(this.btnClose_Click);
            
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
} 