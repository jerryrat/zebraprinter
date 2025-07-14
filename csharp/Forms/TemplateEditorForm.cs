using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Forms
{
    public partial class TemplateEditorForm : Form
    {
        private PrintTemplate _currentTemplate;
        private TemplateEditorEnhancer _enhancer;
        private TemplateAutoCompleteProvider _autoCompleteProvider;
        private bool _hasUnsavedChanges = false;
        private System.Windows.Forms.Timer _validationTimer;
        private ToolTip _tooltip = new ToolTip();

        // 控件
        private RichTextBox rtbEditor;
        private Panel pnlToolbar;
        private Panel pnlVariables;
        private Panel pnlPreview;
        private Panel pnlStatus;
        private Button btnSave;
        private Button btnSaveAs;
        private Button btnNew;
        private Button btnOpen;
        private Button btnPreview;
        private Button btnValidate;
        private ComboBox cmbTemplates;
        private ComboBox cmbFormat;
        private ListBox lstVariables;
        private RichTextBox rtbPreview;
        private Label lblStatus;
        private Label lblLineColumn;
        private CheckBox chkSyntaxHighlight;
        private CheckBox chkAutoPreview;
        private Splitter splitterMain;
        private Splitter splitterRight;

        public TemplateEditorForm()
        {
            InitializeComponent();
            InitializeEnhancer();
            LoadTemplates();
            SetupValidationTimer();
        }

        public TemplateEditorForm(PrintTemplate template) : this()
        {
            LoadTemplate(template);
        }

        private void InitializeComponent()
        {
            this.Text = "打印模板编辑器";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimumSize = new Size(800, 600);

            // 创建主容器
            CreateMainLayout();
            CreateToolbar();
            CreateVariablesPanel();
            CreatePreviewPanel();
            CreateStatusPanel();

            SetupEventHandlers();
        }

        private void CreateMainLayout()
        {
            // 工具栏
            pnlToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = SystemColors.Control
            };
            this.Controls.Add(pnlToolbar);

            // 状态栏
            pnlStatus = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 25,
                BackColor = SystemColors.Control
            };
            this.Controls.Add(pnlStatus);

            // 左侧变量面板
            pnlVariables = new Panel
            {
                Dock = DockStyle.Left,
                Width = 200,
                BackColor = SystemColors.ControlLight
            };
            this.Controls.Add(pnlVariables);

            // 左侧分割线
            splitterMain = new Splitter
            {
                Dock = DockStyle.Left,
                Width = 3
            };
            this.Controls.Add(splitterMain);

            // 右侧预览面板
            pnlPreview = new Panel
            {
                Dock = DockStyle.Right,
                Width = 300,
                BackColor = SystemColors.ControlLight
            };
            this.Controls.Add(pnlPreview);

            // 右侧分割线
            splitterRight = new Splitter
            {
                Dock = DockStyle.Right,
                Width = 3
            };
            this.Controls.Add(splitterRight);

            // 编辑器面板（填充剩余空间）
            Panel pnlEditor = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            this.Controls.Add(pnlEditor);

            // 编辑器
            rtbEditor = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10),
                AcceptsTab = true,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both
            };
            pnlEditor.Controls.Add(rtbEditor);
        }

        private void CreateToolbar()
        {
            int x = 10, y = 10;

            // 文件操作
            btnNew = CreateButton("新建", x, y, btnNew_Click);
            x += 70;

            btnOpen = CreateButton("打开", x, y, btnOpen_Click);
            x += 70;

            btnSave = CreateButton("保存", x, y, btnSave_Click);
            x += 70;

            btnSaveAs = CreateButton("另存为", x, y, btnSaveAs_Click);
            x += 80;

            // 分隔线
            x += 20;

            // 模板选择
            Label lblTemplate = new Label
            {
                Text = "模板:",
                Location = new Point(x, y + 5),
                Size = new Size(40, 20)
            };
            pnlToolbar.Controls.Add(lblTemplate);
            x += 45;

            cmbTemplates = new ComboBox
            {
                Location = new Point(x, y),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            pnlToolbar.Controls.Add(cmbTemplates);
            x += 160;

            // 格式选择
            Label lblFormat = new Label
            {
                Text = "格式:",
                Location = new Point(x, y + 5),
                Size = new Size(40, 20)
            };
            pnlToolbar.Controls.Add(lblFormat);
            x += 45;

            cmbFormat = new ComboBox
            {
                Location = new Point(x, y),
                Size = new Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbFormat.Items.AddRange(Enum.GetNames(typeof(PrintFormat)));
            cmbFormat.SelectedIndex = 0;
            pnlToolbar.Controls.Add(cmbFormat);
            x += 110;

            // 第二行
            x = 10;
            y = 40;

            // 工具按钮
            btnValidate = CreateButton("验证", x, y, btnValidate_Click);
            x += 70;

            btnPreview = CreateButton("预览", x, y, btnPreview_Click);
            x += 70;

            // 选项
            chkSyntaxHighlight = new CheckBox
            {
                Text = "语法高亮",
                Location = new Point(x, y + 3),
                Size = new Size(80, 20),
                Checked = true
            };
            pnlToolbar.Controls.Add(chkSyntaxHighlight);
            x += 90;

            chkAutoPreview = new CheckBox
            {
                Text = "自动预览",
                Location = new Point(x, y + 3),
                Size = new Size(80, 20),
                Checked = true
            };
            pnlToolbar.Controls.Add(chkAutoPreview);
        }

        private Button CreateButton(string text, int x, int y, EventHandler clickHandler)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(60, 25),
                UseVisualStyleBackColor = true
            };
            button.Click += clickHandler;
            pnlToolbar.Controls.Add(button);
            return button;
        }

        private void CreateVariablesPanel()
        {
            // 标题
            Label lblVariablesTitle = new Label
            {
                Text = "可用变量",
                Dock = DockStyle.Top,
                Height = 25,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = SystemColors.ActiveCaption,
                ForeColor = SystemColors.ActiveCaptionText
            };
            pnlVariables.Controls.Add(lblVariablesTitle);

            // 变量列表
            lstVariables = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9),
                HorizontalScrollbar = true
            };
            pnlVariables.Controls.Add(lstVariables);

            // 加载变量
            LoadVariables();
        }

        private void CreatePreviewPanel()
        {
            // 标题
            Label lblPreviewTitle = new Label
            {
                Text = "实时预览",
                Dock = DockStyle.Top,
                Height = 25,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = SystemColors.ActiveCaption,
                ForeColor = SystemColors.ActiveCaptionText
            };
            pnlPreview.Controls.Add(lblPreviewTitle);

            // 预览内容
            rtbPreview = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.WhiteSmoke,
                ScrollBars = RichTextBoxScrollBars.Both
            };
            pnlPreview.Controls.Add(rtbPreview);
        }

        private void CreateStatusPanel()
        {
            lblStatus = new Label
            {
                Text = "就绪",
                Dock = DockStyle.Left,
                Width = 200,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(5, 0, 0, 0)
            };
            pnlStatus.Controls.Add(lblStatus);

            lblLineColumn = new Label
            {
                Text = "行: 1, 列: 1",
                Dock = DockStyle.Right,
                Width = 100,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 5, 0)
            };
            pnlStatus.Controls.Add(lblLineColumn);
        }

        private void InitializeEnhancer()
        {
            _enhancer = new TemplateEditorEnhancer(rtbEditor);
            _autoCompleteProvider = new TemplateAutoCompleteProvider(rtbEditor, this);
        }

        private void SetupEventHandlers()
        {
            // 编辑器事件
            rtbEditor.TextChanged += RtbEditor_TextChanged;
            rtbEditor.SelectionChanged += RtbEditor_SelectionChanged;
            rtbEditor.KeyDown += RtbEditor_KeyDown;

            // 变量列表事件
            lstVariables.DoubleClick += LstVariables_DoubleClick;
            lstVariables.MouseMove += LstVariables_MouseMove;

            // 模板选择事件
            cmbTemplates.SelectedIndexChanged += CmbTemplates_SelectedIndexChanged;

            // 选项事件
            chkSyntaxHighlight.CheckedChanged += ChkSyntaxHighlight_CheckedChanged;
            chkAutoPreview.CheckedChanged += ChkAutoPreview_CheckedChanged;

            // 窗体事件
            this.FormClosing += TemplateEditorForm_FormClosing;
        }

        private void SetupValidationTimer()
        {
            _validationTimer = new System.Windows.Forms.Timer();
            _validationTimer.Interval = 1000; // 1秒延迟
            _validationTimer.Tick += ValidationTimer_Tick;
        }

        private void LoadTemplates()
        {
            cmbTemplates.Items.Clear();
            var templates = PrintTemplateManager.GetTemplates();
            foreach (var template in templates)
            {
                cmbTemplates.Items.Add(template.Name);
            }
            
            if (cmbTemplates.Items.Count > 0)
            {
                cmbTemplates.SelectedIndex = 0;
            }
        }

        private void LoadVariables()
        {
            lstVariables.Items.Clear();
            var fields = PrintTemplateManager.GetAvailableFields();
            var descriptions = PrintTemplateManager.GetFieldDescriptions();

            foreach (var field in fields)
            {
                var description = descriptions.ContainsKey(field) ? descriptions[field] : "";
                lstVariables.Items.Add($"{field} - {description}");
            }
        }

        private void LoadTemplate(PrintTemplate template)
        {
            if (template == null) return;

            _currentTemplate = template;
            rtbEditor.Text = template.Content;
            cmbFormat.SelectedItem = template.Format.ToString();
            
            this.Text = $"打印模板编辑器 - {template.Name}";
            _hasUnsavedChanges = false;
            
            UpdatePreview();
            UpdateStatus("模板已加载");
        }

        #region 事件处理

        private void RtbEditor_TextChanged(object sender, EventArgs e)
        {
            _hasUnsavedChanges = true;
            this.Text = this.Text.TrimEnd('*') + "*";

            if (chkSyntaxHighlight.Checked)
            {
                _enhancer.ApplySyntaxHighlighting();
            }

            if (chkAutoPreview.Checked)
            {
                _validationTimer.Stop();
                _validationTimer.Start();
            }

            UpdateLineColumnInfo();
        }

        private void RtbEditor_SelectionChanged(object sender, EventArgs e)
        {
            UpdateLineColumnInfo();
        }

        private void RtbEditor_KeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl+S 保存
            if (e.Control && e.KeyCode == Keys.S)
            {
                btnSave_Click(sender, e);
                e.Handled = true;
            }
            // Ctrl+Shift+S 另存为
            else if (e.Control && e.Shift && e.KeyCode == Keys.S)
            {
                btnSaveAs_Click(sender, e);
                e.Handled = true;
            }
            // F5 预览
            else if (e.KeyCode == Keys.F5)
            {
                btnPreview_Click(sender, e);
                e.Handled = true;
            }
        }

        private void LstVariables_DoubleClick(object sender, EventArgs e)
        {
            if (lstVariables.SelectedItem != null)
            {
                string item = lstVariables.SelectedItem.ToString();
                string variable = item.Split(' ')[0]; // 获取变量名部分
                _enhancer.InsertVariable(variable.Trim('{', '}'));
                rtbEditor.Focus();
            }
        }

        private void LstVariables_MouseMove(object sender, MouseEventArgs e)
        {
            int index = lstVariables.IndexFromPoint(e.Location);
            if (index >= 0 && index < lstVariables.Items.Count)
            {
                string item = lstVariables.Items[index].ToString();
                _tooltip.SetToolTip(lstVariables, item);
            }
        }

        private void CmbTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTemplates.SelectedItem != null)
            {
                string templateName = cmbTemplates.SelectedItem.ToString();
                var template = PrintTemplateManager.GetTemplate(templateName);
                if (template != null)
                {
                    if (ConfirmUnsavedChanges())
                    {
                        LoadTemplate(template);
                    }
                    else
                    {
                        // 恢复之前的选择
                        if (_currentTemplate != null)
                        {
                            cmbTemplates.SelectedItem = _currentTemplate.Name;
                        }
                    }
                }
            }
        }

        private void ChkSyntaxHighlight_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSyntaxHighlight.Checked)
            {
                _enhancer.ApplySyntaxHighlighting();
                UpdateStatus("语法高亮已启用");
            }
            else
            {
                // 清除格式
                rtbEditor.SelectAll();
                rtbEditor.SelectionColor = Color.Black;
                rtbEditor.SelectionFont = new Font(rtbEditor.Font, FontStyle.Regular);
                rtbEditor.DeselectAll();
                UpdateStatus("语法高亮已禁用");
            }
        }

        private void ChkAutoPreview_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAutoPreview.Checked)
            {
                UpdatePreview();
                UpdateStatus("自动预览已启用");
            }
            else
            {
                UpdateStatus("自动预览已禁用");
            }
        }

        private void ValidationTimer_Tick(object sender, EventArgs e)
        {
            _validationTimer.Stop();
            UpdatePreview();
            ValidateTemplate();
        }

        private void TemplateEditorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!ConfirmUnsavedChanges())
            {
                e.Cancel = true;
            }
        }

        #endregion

        #region 按钮事件

        private void btnNew_Click(object sender, EventArgs e)
        {
            if (ConfirmUnsavedChanges())
            {
                CreateNewTemplate();
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            // 显示模板选择对话框或文件选择对话框
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "模板文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
                openFileDialog.Title = "打开模板文件";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string content = System.IO.File.ReadAllText(openFileDialog.FileName);
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                        
                        var template = new PrintTemplate
                        {
                            Name = fileName,
                            Content = content,
                            Format = (PrintFormat)Enum.Parse(typeof(PrintFormat), cmbFormat.SelectedItem.ToString())
                        };
                        
                        LoadTemplate(template);
                        UpdateStatus($"已从文件加载: {openFileDialog.FileName}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"加载文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveCurrentTemplate();
        }

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            SaveTemplateAs();
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            UpdatePreview();
            UpdateStatus("预览已更新");
        }

        private void btnValidate_Click(object sender, EventArgs e)
        {
            ValidateTemplate();
        }

        #endregion

        #region 辅助方法

        private void CreateNewTemplate()
        {
            _currentTemplate = new PrintTemplate
            {
                Name = "新模板",
                Content = "",
                Format = PrintFormat.Text,
                IsDefault = false
            };
            
            rtbEditor.Clear();
            this.Text = "打印模板编辑器 - 新模板";
            _hasUnsavedChanges = false;
            
            UpdateStatus("新模板已创建");
        }

        private bool ConfirmUnsavedChanges()
        {
            if (_hasUnsavedChanges)
            {
                var result = MessageBox.Show(
                    "当前模板有未保存的更改，是否保存？",
                    "确认",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    return SaveCurrentTemplate();
                }
                else if (result == DialogResult.Cancel)
                {
                    return false;
                }
            }
            return true;
        }

        private bool SaveCurrentTemplate()
        {
            if (_currentTemplate == null)
            {
                return SaveTemplateAs();
            }

            try
            {
                _currentTemplate.Content = rtbEditor.Text;
                _currentTemplate.Format = (PrintFormat)Enum.Parse(typeof(PrintFormat), cmbFormat.SelectedItem.ToString());
                
                PrintTemplateManager.SaveTemplate(_currentTemplate);
                
                _hasUnsavedChanges = false;
                this.Text = this.Text.TrimEnd('*');
                
                LoadTemplates(); // 刷新模板列表
                cmbTemplates.SelectedItem = _currentTemplate.Name;
                
                UpdateStatus($"模板 '{_currentTemplate.Name}' 已保存");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool SaveTemplateAs()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "另存为模板";
                dialog.Size = new Size(400, 200);
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;

                var lblName = new Label
                {
                    Text = "模板名称:",
                    Location = new Point(20, 30),
                    Size = new Size(80, 20)
                };
                dialog.Controls.Add(lblName);

                var txtName = new TextBox
                {
                    Location = new Point(110, 28),
                    Size = new Size(250, 25),
                    Text = _currentTemplate?.Name ?? "新模板"
                };
                dialog.Controls.Add(txtName);

                var chkDefault = new CheckBox
                {
                    Text = "设为默认模板",
                    Location = new Point(110, 65),
                    Size = new Size(120, 20),
                    Checked = _currentTemplate?.IsDefault ?? false
                };
                dialog.Controls.Add(chkDefault);

                var btnOK = new Button
                {
                    Text = "确定",
                    Location = new Point(200, 110),
                    Size = new Size(75, 25),
                    DialogResult = DialogResult.OK
                };
                dialog.Controls.Add(btnOK);

                var btnCancel = new Button
                {
                    Text = "取消",
                    Location = new Point(285, 110),
                    Size = new Size(75, 25),
                    DialogResult = DialogResult.Cancel
                };
                dialog.Controls.Add(btnCancel);

                dialog.AcceptButton = btnOK;
                dialog.CancelButton = btnCancel;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var template = new PrintTemplate
                        {
                            Name = txtName.Text.Trim(),
                            Content = rtbEditor.Text,
                            Format = (PrintFormat)Enum.Parse(typeof(PrintFormat), cmbFormat.SelectedItem.ToString()),
                            IsDefault = chkDefault.Checked
                        };

                        if (string.IsNullOrEmpty(template.Name))
                        {
                            MessageBox.Show("请输入模板名称", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return false;
                        }

                        PrintTemplateManager.SaveTemplate(template);
                        LoadTemplate(template);
                        LoadTemplates();
                        
                        UpdateStatus($"模板已另存为 '{template.Name}'");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
            return false;
        }

        private void UpdatePreview()
        {
            try
            {
                if (string.IsNullOrEmpty(rtbEditor.Text))
                {
                    rtbPreview.Text = "预览内容为空";
                    return;
                }

                // 创建示例测试记录
                var sampleRecord = CreateSampleTestRecord();
                
                var template = new PrintTemplate
                {
                    Content = rtbEditor.Text,
                    Format = PrintFormat.Text
                };

                string preview = PrintTemplateManager.ProcessTemplate(template, sampleRecord);
                rtbPreview.Text = preview;
            }
            catch (Exception ex)
            {
                rtbPreview.Text = $"预览错误: {ex.Message}";
            }
        }

        private TestRecord CreateSampleTestRecord()
        {
            return TestRecord.CreateSample();
        }

        private void ValidateTemplate()
        {
            var result = _enhancer.ValidateTemplate();
            UpdateStatus(result.GetSummary());
            
            // 如果有错误，可以在状态栏显示详细信息
            if (!result.IsValid || result.Warnings.Count > 0)
            {
                _tooltip.SetToolTip(lblStatus, result.GetSummary());
            }
        }

        private void UpdateStatus(string message)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = SystemColors.ControlText;
        }

        private void UpdateLineColumnInfo()
        {
            int line = rtbEditor.GetLineFromCharIndex(rtbEditor.SelectionStart) + 1;
            int column = rtbEditor.SelectionStart - rtbEditor.GetFirstCharIndexFromLine(line - 1) + 1;
            lblLineColumn.Text = $"行: {line}, 列: {column}";
        }

        #endregion

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _validationTimer?.Dispose();
                _tooltip?.Dispose();
                _autoCompleteProvider?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 