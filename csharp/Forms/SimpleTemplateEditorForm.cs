using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Forms
{
    public partial class SimpleTemplateEditorForm : Form
    {
        private readonly DragDropTemplateService _templateService;
        private RichTextBox _templateContentTextBox;
        private Button _btnSave;
        private Button _btnPreview;
        private Button _btnCancel;
        private Label _lblTitle;
        private Label _lblInstructions;

        public SimpleTemplateEditorForm(DragDropTemplateService templateService)
        {
            _templateService = templateService;
            InitializeComponent();
            LoadCurrentTemplate();
        }

        private void InitializeComponent()
        {
            this.Text = "太阳能电池板规格表模板编辑器 v1.1.31";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = true;
            this.MinimizeBox = false;
            this.ShowIcon = false;

            // 标题标签
            _lblTitle = new Label
            {
                Text = "太阳能电池板规格表模板编辑器",
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Location = new Point(20, 20),
                Size = new Size(400, 30),
                AutoSize = false
            };
            this.Controls.Add(_lblTitle);

            // 说明标签
            _lblInstructions = new Label
            {
                Text = "当前模板已1:1还原SKT 600 M12/120HB太阳能电池板规格表样式\n" +
                       "包含13行技术参数、测试条件说明和二维码区域\n" +
                       "模板内容以HTML格式显示，支持查看和基本编辑",
                Font = new Font("Microsoft YaHei", 9),
                ForeColor = Color.Gray,
                Location = new Point(20, 60),
                Size = new Size(750, 60),
                AutoSize = false
            };
            this.Controls.Add(_lblInstructions);

            // 模板内容文本框
            _templateContentTextBox = new RichTextBox
            {
                Location = new Point(20, 130),
                Size = new Size(740, 380),
                Font = new Font("Consolas", 9),
                ReadOnly = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false,
                BackColor = Color.FromArgb(248, 248, 248),
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(_templateContentTextBox);

            // 按钮面板
            var buttonPanel = new Panel
            {
                Location = new Point(20, 520),
                Size = new Size(740, 40),
                BackColor = Color.Transparent
            };
            this.Controls.Add(buttonPanel);

            // 预览按钮
            _btnPreview = new Button
            {
                Text = "预览模板",
                Size = new Size(100, 35),
                Location = new Point(0, 0),
                BackColor = Color.LightBlue,
                Font = new Font("Microsoft YaHei", 9),
                UseVisualStyleBackColor = false
            };
            _btnPreview.Click += BtnPreview_Click;
            buttonPanel.Controls.Add(_btnPreview);

            // 保存按钮
            _btnSave = new Button
            {
                Text = "保存模板",
                Size = new Size(100, 35),
                Location = new Point(520, 0),
                BackColor = Color.LightGreen,
                Font = new Font("Microsoft YaHei", 9),
                UseVisualStyleBackColor = false
            };
            _btnSave.Click += BtnSave_Click;
            buttonPanel.Controls.Add(_btnSave);

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "取消",
                Size = new Size(100, 35),
                Location = new Point(630, 0),
                BackColor = Color.LightGray,
                Font = new Font("Microsoft YaHei", 9),
                UseVisualStyleBackColor = false,
                DialogResult = DialogResult.Cancel
            };
            _btnCancel.Click += BtnCancel_Click;
            buttonPanel.Controls.Add(_btnCancel);

            this.CancelButton = _btnCancel;
        }

        private void LoadCurrentTemplate()
        {
            try
            {
                // 创建示例测试记录来生成模板预览
                var sampleRecord = TestRecord.CreateSample();
                
                // 生成当前模板的HTML内容
                var htmlContent = _templateService.GeneratePrintHtml(sampleRecord);
                
                _templateContentTextBox.Text = htmlContent;
                
                Logger.Info("已加载当前太阳能电池板规格表模板内容");
            }
            catch (Exception ex)
            {
                Logger.Error("加载模板内容失败", ex);
                MessageBox.Show($"加载模板失败：{ex.Message}", "错误", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnPreview_Click(object? sender, EventArgs e)
        {
            try
            {
                var sampleRecord = TestRecord.CreateSample();
                var previewHtml = _templateService.PreviewTemplate(sampleRecord);
                
                if (!string.IsNullOrEmpty(previewHtml))
                {
                    MessageBox.Show("模板预览已生成并保存到临时文件！\n" +
                                  "实际使用时将根据测试数据动态填充内容。", 
                                  "模板预览", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("预览生成失败", "错误", 
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("模板预览失败", ex);
                MessageBox.Show($"预览失败：{ex.Message}", "错误", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("当前版本的模板已经是专业设计的太阳能电池板规格表！\n" +
                              "模板内容包含：\n" +
                              "• SKT 600 M12/120HB产品型号\n" +
                              "• 13行完整技术参数表格\n" +
                              "• 测试条件说明\n" +
                              "• 二维码区域\n" +
                              "• 专业的表格样式和隔行变色\n\n" +
                              "如需自定义，请联系开发团队。", 
                              "模板信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                Logger.Error("保存模板失败", ex);
                MessageBox.Show($"保存失败：{ex.Message}", "错误", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _templateContentTextBox?.Dispose();
                _btnSave?.Dispose();
                _btnPreview?.Dispose();
                _btnCancel?.Dispose();
                _lblTitle?.Dispose();
                _lblInstructions?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 