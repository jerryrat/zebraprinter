using System;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Services;

namespace ZebraPrinterMonitor.Forms
{
    public partial class HeaderFooterSettingsForm : Form
    {
        private PrintTemplate _template;
        private CheckBox _showHeaderCheckBox;
        private CheckBox _showFooterCheckBox;
        private TextBox _headerTextTextBox;
        private TextBox _footerTextTextBox;
        private TextBox _headerImageTextBox;
        private TextBox _footerImageTextBox;
        private Button _browseHeaderImageButton;
        private Button _browseFooterImageButton;
        private Button _okButton;
        private Button _cancelButton;

        public HeaderFooterSettingsForm(PrintTemplate template)
        {
            _template = template ?? throw new ArgumentNullException(nameof(template));
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "页眉页脚设置";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 页眉设置组
            var headerGroupBox = new GroupBox
            {
                Text = "页眉设置",
                Location = new Point(15, 15),
                Size = new Size(450, 120)
            };

            _showHeaderCheckBox = new CheckBox
            {
                Text = "显示页眉",
                Location = new Point(15, 25),
                Size = new Size(100, 23)
            };

            var headerTextLabel = new Label
            {
                Text = "页眉文本:",
                Location = new Point(15, 55),
                Size = new Size(70, 23)
            };

            _headerTextTextBox = new TextBox
            {
                Location = new Point(95, 52),
                Size = new Size(335, 23),
                PlaceholderText = "输入页眉文本"
            };

            var headerImageLabel = new Label
            {
                Text = "页眉图片:",
                Location = new Point(15, 85),
                Size = new Size(70, 23)
            };

            _headerImageTextBox = new TextBox
            {
                Location = new Point(95, 82),
                Size = new Size(250, 23),
                ReadOnly = true,
                PlaceholderText = "选择图片文件"
            };

            _browseHeaderImageButton = new Button
            {
                Text = "浏览",
                Location = new Point(355, 81),
                Size = new Size(75, 25)
            };

            headerGroupBox.Controls.AddRange(new Control[] {
                _showHeaderCheckBox, headerTextLabel, _headerTextTextBox,
                headerImageLabel, _headerImageTextBox, _browseHeaderImageButton
            });

            // 页脚设置组
            var footerGroupBox = new GroupBox
            {
                Text = "页脚设置",
                Location = new Point(15, 150),
                Size = new Size(450, 120)
            };

            _showFooterCheckBox = new CheckBox
            {
                Text = "显示页脚",
                Location = new Point(15, 25),
                Size = new Size(100, 23)
            };

            var footerTextLabel = new Label
            {
                Text = "页脚文本:",
                Location = new Point(15, 55),
                Size = new Size(70, 23)
            };

            _footerTextTextBox = new TextBox
            {
                Location = new Point(95, 52),
                Size = new Size(335, 23),
                PlaceholderText = "输入页脚文本"
            };

            var footerImageLabel = new Label
            {
                Text = "页脚图片:",
                Location = new Point(15, 85),
                Size = new Size(70, 23)
            };

            _footerImageTextBox = new TextBox
            {
                Location = new Point(95, 82),
                Size = new Size(250, 23),
                ReadOnly = true,
                PlaceholderText = "选择图片文件"
            };

            _browseFooterImageButton = new Button
            {
                Text = "浏览",
                Location = new Point(355, 81),
                Size = new Size(75, 25)
            };

            footerGroupBox.Controls.AddRange(new Control[] {
                _showFooterCheckBox, footerTextLabel, _footerTextTextBox,
                footerImageLabel, _footerImageTextBox, _browseFooterImageButton
            });

            // 按钮
            _okButton = new Button
            {
                Text = "确定",
                Location = new Point(310, 320),
                Size = new Size(75, 30),
                DialogResult = DialogResult.OK
            };

            _cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(390, 320),
                Size = new Size(75, 30),
                DialogResult = DialogResult.Cancel
            };

            // 绑定事件
            _browseHeaderImageButton.Click += BrowseHeaderImage_Click;
            _browseFooterImageButton.Click += BrowseFooterImage_Click;
            _okButton.Click += OkButton_Click;

            this.Controls.AddRange(new Control[] {
                headerGroupBox, footerGroupBox, _okButton, _cancelButton
            });

            this.AcceptButton = _okButton;
            this.CancelButton = _cancelButton;
        }

        private void LoadSettings()
        {
            _showHeaderCheckBox.Checked = _template.ShowHeader;
            _headerTextTextBox.Text = _template.HeaderText;
            _headerImageTextBox.Text = _template.HeaderImagePath;
            _showFooterCheckBox.Checked = _template.ShowFooter;
            _footerTextTextBox.Text = _template.FooterText;
            _footerImageTextBox.Text = _template.FooterImagePath;
        }

        private void BrowseHeaderImage_Click(object? sender, EventArgs e)
        {
            using var openFileDialog = new OpenFileDialog
            {
                Title = "选择页眉图片",
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|所有文件|*.*",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                _headerImageTextBox.Text = openFileDialog.FileName;
            }
        }

        private void BrowseFooterImage_Click(object? sender, EventArgs e)
        {
            using var openFileDialog = new OpenFileDialog
            {
                Title = "选择页脚图片",
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|所有文件|*.*",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                _footerImageTextBox.Text = openFileDialog.FileName;
            }
        }

        private void OkButton_Click(object? sender, EventArgs e)
        {
            // 更新模板设置（不直接保存，让主窗体决定何时保存）
            _template.ShowHeader = _showHeaderCheckBox.Checked;
            _template.HeaderText = _headerTextTextBox.Text;
            _template.HeaderImagePath = _headerImageTextBox.Text;
            _template.ShowFooter = _showFooterCheckBox.Checked;
            _template.FooterText = _footerTextTextBox.Text;
            _template.FooterImagePath = _footerImageTextBox.Text;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
} 