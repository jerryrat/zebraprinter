using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;
using System.Linq;

namespace ZebraPrinterMonitor.Forms
{
    public partial class TemplateDesignerForm : Form
    {
        private PrintTemplate _currentTemplate;
        private Panel _designPanel;
        private ListBox _fieldListBox;
        private TextBox _templateNameTextBox;
        private ComboBox _formatComboBox;
        private Button _saveButton;
        private Button _previewButton;
        private Button _cancelButton;
        private RichTextBox _previewTextBox;
        private Dictionary<string, string> _availableFields;
        private List<FieldControl> _fieldControls;
        private FieldControl _selectedField;
        private bool _isDragging;
        private Point _dragStartPoint;

        public TemplateDesignerForm()
        {
            InitializeFields();  // 先初始化字段
            _fieldControls = new List<FieldControl>();
            _currentTemplate = new PrintTemplate 
            { 
                Name = "新模板", 
                Content = "",
                Format = PrintFormat.Text 
            };
            InitializeComponent();
        }

        public TemplateDesignerForm(PrintTemplate template) : this()
        {
            _currentTemplate = template;
            LoadTemplate();
        }

        private void InitializeComponent()
        {
            this.Text = "模板可视化设计器";
            this.Size = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 创建主面板
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2
            };

            // 设置列宽比例
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));

            // 设置行高比例
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 85F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15F));

            // 字段列表面板
            var fieldPanel = new GroupBox
            {
                Text = "可用字段",
                Dock = DockStyle.Fill,
                Margin = new Padding(5)
            };

            var fieldContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1,
                Margin = new Padding(5)
            };

            // 设置行高
            fieldContainer.RowStyles.Add(new RowStyle(SizeType.Percent, 70F)); // 字段列表
            fieldContainer.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // 文本输入
            fieldContainer.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // 添加按钮
            fieldContainer.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // 清空按钮

            _fieldListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.One
            };

            // 填充字段列表
            foreach (var field in _availableFields)
            {
                _fieldListBox.Items.Add($"{field.Key} - {field.Value}");
            }

            _fieldListBox.MouseDown += FieldListBox_MouseDown;
            fieldContainer.Controls.Add(_fieldListBox, 0, 0);

            // 自定义文本输入框
            var customTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                PlaceholderText = "输入自定义文本",
                Margin = new Padding(2)
            };
            fieldContainer.Controls.Add(customTextBox, 0, 1);

            // 添加自定义文本按钮
            var addCustomButton = new Button
            {
                Text = "添加自定义文本",
                Dock = DockStyle.Fill,
                Margin = new Padding(2)
            };
            addCustomButton.Click += (s, e) => AddCustomText_Click(customTextBox);
            fieldContainer.Controls.Add(addCustomButton, 0, 2);

            // 清空设计面板按钮
            var clearDesignButton = new Button
            {
                Text = "清空设计面板",
                Dock = DockStyle.Fill,
                Margin = new Padding(2)
            };
            clearDesignButton.Click += ClearDesign_Click;
            fieldContainer.Controls.Add(clearDesignButton, 0, 3);

            fieldPanel.Controls.Add(fieldContainer);
            mainPanel.Controls.Add(fieldPanel, 0, 0);

            // 设计面板
            var designGroup = new GroupBox
            {
                Text = "设计区域",
                Dock = DockStyle.Fill,
                Margin = new Padding(5)
            };

            _designPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(5),
                AllowDrop = true
            };

            _designPanel.DragEnter += DesignPanel_DragEnter;
            _designPanel.DragDrop += DesignPanel_DragDrop;
            _designPanel.MouseDown += DesignPanel_MouseDown;
            _designPanel.MouseMove += DesignPanel_MouseMove;
            _designPanel.MouseUp += DesignPanel_MouseUp;
            _designPanel.Paint += DesignPanel_Paint;

            designGroup.Controls.Add(_designPanel);
            mainPanel.Controls.Add(designGroup, 1, 0);

            // 属性面板
            var propertyPanel = new GroupBox
            {
                Text = "属性和预览",
                Dock = DockStyle.Fill,
                Margin = new Padding(5)
            };

            var propertyContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 8,
                ColumnCount = 2,
                Margin = new Padding(5)
            };

            // 模板名称
            propertyContainer.Controls.Add(new Label { Text = "模板名称:", Anchor = AnchorStyles.Left }, 0, 0);
            _templateNameTextBox = new TextBox { Dock = DockStyle.Fill, Text = _currentTemplate.Name };
            propertyContainer.Controls.Add(_templateNameTextBox, 1, 0);

            // 格式选择
            propertyContainer.Controls.Add(new Label { Text = "输出格式:", Anchor = AnchorStyles.Left }, 0, 1);
            _formatComboBox = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _formatComboBox.Items.AddRange(new[] { "Text", "ZPL", "Code128", "QRCode" });
            _formatComboBox.SelectedItem = _currentTemplate.Format.ToString();
            propertyContainer.Controls.Add(_formatComboBox, 1, 1);

            // 预览按钮
            _previewButton = new Button { Text = "预览", Dock = DockStyle.Fill };
            _previewButton.Click += PreviewButton_Click;
            propertyContainer.Controls.Add(_previewButton, 0, 2);

            // 清空按钮
            var clearButton = new Button { Text = "清空", Dock = DockStyle.Fill };
            clearButton.Click += ClearButton_Click;
            propertyContainer.Controls.Add(clearButton, 1, 2);

            // 预览文本框
            _previewTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9F),
                WordWrap = false
            };
            propertyContainer.Controls.Add(_previewTextBox, 0, 3);
            propertyContainer.SetColumnSpan(_previewTextBox, 2);
            propertyContainer.SetRowSpan(_previewTextBox, 5);

            // 设置行高
            for (int i = 0; i < 3; i++)
            {
                propertyContainer.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            }
            propertyContainer.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            propertyPanel.Controls.Add(propertyContainer);
            mainPanel.Controls.Add(propertyPanel, 2, 0);

            // 按钮面板
            var buttonPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Margin = new Padding(5)
            };

            _saveButton = new Button { Text = "保存模板", Dock = DockStyle.Fill };
            _saveButton.Click += SaveButton_Click;
            buttonPanel.Controls.Add(_saveButton, 0, 0);

            var loadButton = new Button { Text = "加载模板", Dock = DockStyle.Fill };
            loadButton.Click += LoadButton_Click;
            buttonPanel.Controls.Add(loadButton, 1, 0);

            _cancelButton = new Button { Text = "取消", Dock = DockStyle.Fill };
            _cancelButton.Click += CancelButton_Click;
            buttonPanel.Controls.Add(_cancelButton, 2, 0);

            mainPanel.Controls.Add(buttonPanel, 0, 1);
            mainPanel.SetColumnSpan(buttonPanel, 3);

            this.Controls.Add(mainPanel);
        }

        private void InitializeFields()
        {
            _availableFields = new Dictionary<string, string>
            {
                { "{SerialNumber}", "序列号" },
                { "{Power}", "功率" },
                { "{Voltage}", "电压" },
                { "{Current}", "电流" },
                { "{VoltageVpm}", "Vpm电压" },
                { "{TestDateTime}", "测试时间" },
                { "{PrintCount}", "打印次数" },
                { "{CurrentTime}", "当前时间" },
                { "{CurrentDate}", "当前日期" }
            };
        }

        private void LoadTemplate()
        {
            _templateNameTextBox.Text = _currentTemplate.Name;
            _formatComboBox.SelectedItem = _currentTemplate.Format.ToString();
            
            // 解析现有模板内容并创建字段控件
            // 这里可以根据需要实现模板内容的解析逻辑
        }

        private void FieldListBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && _fieldListBox.SelectedItem != null)
            {
                var selectedText = _fieldListBox.SelectedItem.ToString();
                var fieldKey = selectedText.Split(' ')[0];
                _fieldListBox.DoDragDrop(fieldKey, DragDropEffects.Copy);
            }
        }

        private void DesignPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void DesignPanel_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
            {
                var fieldKey = e.Data.GetData(DataFormats.Text).ToString();
                var dropPoint = _designPanel.PointToClient(new Point(e.X, e.Y));
                
                CreateFieldControl(fieldKey, dropPoint);
            }
        }

        private void CreateFieldControl(string fieldKey, Point location, bool isCustomText = false)
        {
            string displayText;
            if (isCustomText)
            {
                displayText = fieldKey; // 自定义文本直接显示
            }
            else
            {
                displayText = _availableFields.ContainsKey(fieldKey) ? _availableFields[fieldKey] : fieldKey;
            }

            var fieldControl = new FieldControl(fieldKey, displayText, isCustomText)
            {
                Location = location,
                BackColor = isCustomText ? Color.LightGreen : Color.LightBlue,
                BorderStyle = BorderStyle.FixedSingle
            };

            fieldControl.MouseDown += FieldControl_MouseDown;
            fieldControl.MouseMove += FieldControl_MouseMove;
            fieldControl.MouseUp += FieldControl_MouseUp;
            fieldControl.Click += FieldControl_Click;
            fieldControl.DoubleClick += FieldControl_DoubleClick;

            _fieldControls.Add(fieldControl);
            _designPanel.Controls.Add(fieldControl);
        }

        private void FieldControl_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _selectedField = sender as FieldControl;
                _isDragging = true;
                _dragStartPoint = e.Location;
                
                // 高亮选中的字段
                foreach (var control in _fieldControls)
                {
                    control.BackColor = control == _selectedField ? Color.Yellow : Color.LightBlue;
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                // 右键删除
                var field = sender as FieldControl;
                if (field != null)
                {
                    _fieldControls.Remove(field);
                    _designPanel.Controls.Remove(field);
                    field.Dispose();
                }
            }
        }

        private void FieldControl_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging && _selectedField != null)
            {
                // 将鼠标位置转换为设计面板的坐标
                var screenPoint = _selectedField.PointToScreen(e.Location);
                var panelPoint = _designPanel.PointToClient(screenPoint);
                
                // 计算新位置，减去拖拽开始点的偏移
                var newLocation = new Point(
                    panelPoint.X - _dragStartPoint.X,
                    panelPoint.Y - _dragStartPoint.Y
                );
                
                // 确保不超出边界
                newLocation.X = Math.Max(0, Math.Min(newLocation.X, _designPanel.Width - _selectedField.Width));
                newLocation.Y = Math.Max(0, Math.Min(newLocation.Y, _designPanel.Height - _selectedField.Height));
                
                _selectedField.Location = newLocation;
            }
        }

        private void FieldControl_MouseUp(object sender, MouseEventArgs e)
        {
            _isDragging = false;
        }

        private void FieldControl_Click(object sender, EventArgs e)
        {
            _selectedField = sender as FieldControl;
            
            // 高亮选中的字段
            foreach (var control in _fieldControls)
            {
                control.BackColor = control == _selectedField ? Color.Yellow : Color.LightBlue;
            }
        }

        private void FieldControl_DoubleClick(object sender, EventArgs e)
        {
            var field = sender as FieldControl;
            if (field != null && field.IsCustomText)
            {
                // 双击编辑自定义文本
                var result = Microsoft.VisualBasic.Interaction.InputBox(
                    "编辑自定义文本:", 
                    "编辑文本", 
                    field.FieldKey);
                
                if (!string.IsNullOrWhiteSpace(result))
                {
                    field.FieldKey = result.Trim();
                    field.DisplayName = result.Trim();
                    field.Controls.OfType<Label>().First().Text = result.Trim();
                }
            }
        }

        private void DesignPanel_MouseDown(object sender, MouseEventArgs e)
        {
            // 点击空白区域取消选择
            _selectedField = null;
            foreach (var control in _fieldControls)
            {
                control.BackColor = Color.LightBlue;
            }
        }

        private void DesignPanel_MouseMove(object sender, MouseEventArgs e)
        {
            // 可以在这里添加网格对齐等功能
        }

        private void DesignPanel_MouseUp(object sender, MouseEventArgs e)
        {
            _isDragging = false;
        }

        private void DesignPanel_Paint(object sender, PaintEventArgs e)
        {
            // 绘制网格线
            var graphics = e.Graphics;
            var pen = new Pen(Color.LightGray, 1);
            
            // 绘制垂直网格线
            for (int x = 0; x < _designPanel.Width; x += 20)
            {
                graphics.DrawLine(pen, x, 0, x, _designPanel.Height);
            }
            
            // 绘制水平网格线
            for (int y = 0; y < _designPanel.Height; y += 20)
            {
                graphics.DrawLine(pen, 0, y, _designPanel.Width, y);
            }
        }

        private void PreviewButton_Click(object sender, EventArgs e)
        {
            GeneratePreview();
        }

        private void GeneratePreview()
        {
            try
            {
                // 生成模板内容
                var template = GenerateTemplateFromDesign();
                
                // 创建示例数据
                var sampleRecord = new TestRecord
                {
                    TR_SerialNum = "ABC-1234567",
                    TR_DateTime = DateTime.Now,
                    TR_Isc = 12.34m,
                    TR_Voc = 45.67m,
                    TR_Vpm = 38.90m,
                    TR_Pm = 123.45m,
                    TR_Print = 1
                };

                // 处理模板并显示预览
                var preview = PrintTemplateManager.ProcessTemplate(template, sampleRecord);
                _previewTextBox.Text = preview;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"预览生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private PrintTemplate GenerateTemplateFromDesign()
        {
            var template = new PrintTemplate
            {
                Name = _templateNameTextBox.Text,
                Format = Enum.Parse<PrintFormat>(_formatComboBox.SelectedItem.ToString()),
                Content = GenerateContentFromFields()
            };

            return template;
        }

        private string GenerateContentFromFields()
        {
            if (_fieldControls.Count == 0)
                return "";

            // 根据字段位置生成文本布局
            var sortedFields = _fieldControls.OrderBy(f => f.Location.Y).ThenBy(f => f.Location.X).ToList();
            var content = "";

            foreach (var field in sortedFields)
            {
                if (field.IsCustomText)
                {
                    // 自定义文本直接添加
                    content += field.FieldKey + "\r\n";
                }
                else
                {
                    // 预定义字段添加为模板变量
                    content += $"{field.DisplayName}: {field.FieldKey}\r\n";
                }
            }

            return content.TrimEnd('\r', '\n');
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_templateNameTextBox.Text))
                {
                    MessageBox.Show("请输入模板名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var template = GenerateTemplateFromDesign();
                PrintTemplateManager.SaveTemplate(template);
                
                MessageBox.Show("模板保存成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            var templates = PrintTemplateManager.GetTemplates();
            var templateNames = templates.Select(t => t.Name).ToArray();

            if (templateNames.Length == 0)
            {
                MessageBox.Show("没有可用的模板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 简单的模板选择对话框
            var form = new Form
            {
                Text = "选择模板",
                Size = new Size(300, 150),
                StartPosition = FormStartPosition.CenterParent
            };
            
            var comboBox = new ComboBox
            {
                Location = new Point(20, 20),
                Size = new Size(250, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            comboBox.Items.AddRange(templateNames);
            comboBox.SelectedIndex = 0;
            
            var okButton = new Button
            {
                Text = "确定",
                Location = new Point(100, 60),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            
            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(180, 60),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };
            
            form.Controls.AddRange(new Control[] { comboBox, okButton, cancelButton });
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;
            
            string selectedTemplate = null;
            if (form.ShowDialog() == DialogResult.OK)
            {
                selectedTemplate = comboBox.SelectedItem.ToString();
            }

            if (!string.IsNullOrEmpty(selectedTemplate))
            {
                var template = PrintTemplateManager.GetTemplate(selectedTemplate);
                if (template != null)
                {
                    _currentTemplate = template;
                    LoadTemplate();
                }
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            foreach (var control in _fieldControls.ToList())
            {
                _designPanel.Controls.Remove(control);
                control.Dispose();
            }
            _fieldControls.Clear();
            _previewTextBox.Clear();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void AddCustomText_Click(TextBox textBox)
        {
            if (!string.IsNullOrWhiteSpace(textBox.Text))
            {
                var customText = textBox.Text.Trim();
                // 在设计面板中心位置添加自定义文本
                var centerPoint = new Point(
                    _designPanel.Width / 2 - 60, 
                    _designPanel.Height / 2 - 12
                );
                CreateFieldControl(customText, centerPoint, true);
                textBox.Clear();
            }
            else
            {
                MessageBox.Show("请输入自定义文本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ClearDesign_Click(object sender, EventArgs e)
        {
            foreach (var control in _fieldControls.ToList())
            {
                _designPanel.Controls.Remove(control);
                control.Dispose();
            }
            _fieldControls.Clear();
            _previewTextBox.Clear();
        }
    }

    // 字段控件类
    public class FieldControl : UserControl
    {
        public string FieldKey { get; set; }
        public string DisplayName { get; set; }
        public bool IsCustomText { get; private set; }
        private Label _label;

        public FieldControl(string fieldKey, string displayName, bool isCustomText = false)
        {
            FieldKey = fieldKey;
            DisplayName = displayName;
            IsCustomText = isCustomText;
            
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(120, 25);
            this.BorderStyle = BorderStyle.FixedSingle;
            this.BackColor = IsCustomText ? Color.LightGreen : Color.LightBlue;
            this.Cursor = Cursors.SizeAll;

            _label = new Label
            {
                Text = DisplayName,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8F),
                ForeColor = Color.Black
            };

            this.Controls.Add(_label);
        }
    }
} 