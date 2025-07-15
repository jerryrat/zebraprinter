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
        private TextBox _customTextBox;
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
            this.Size = new Size(1200, 800);
            this.Text = LanguageManager.GetString("VisualDesignerTitle");
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // 创建主面板
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(10)
            };
            
            // 设置列和行的大小
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));
            
            // 字段选择面板
            var fieldPanel = new Panel { Dock = DockStyle.Fill };
            var fieldGroupBox = new GroupBox 
            { 
                Text = LanguageManager.GetString("AvailableFieldsDesc"), 
                Dock = DockStyle.Fill, 
                Padding = new Padding(10) 
            };
            
            var fieldListBox = new ListBox
            {
                Dock = DockStyle.Left,
                Width = 200,
                Items = { "{SerialNumber}", "{TestDateTime}", "{Current}", "{Voltage}", "{VoltageVpm}", "{Power}", "{PrintCount}" }
            };
            
            var customTextPanel = new Panel { Dock = DockStyle.Fill };
            _customTextBox = new TextBox
            {
                PlaceholderText = LanguageManager.GetString("CustomTextPlaceholder"),
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(10, 0, 0, 5)
            };
            
            var addCustomTextButton = new Button
            {
                Text = LanguageManager.GetString("AddCustomText"),
                Dock = DockStyle.Top,
                Height = 30,
                Margin = new Padding(10, 0, 0, 5)
            };
            
            var clearDesignButton = new Button
            {
                Text = LanguageManager.GetString("ClearDesignArea"),
                Dock = DockStyle.Top,
                Height = 30,
                Margin = new Padding(10, 0, 0, 5)
            };
            
            customTextPanel.Controls.Add(clearDesignButton);
            customTextPanel.Controls.Add(addCustomTextButton);
            customTextPanel.Controls.Add(_customTextBox);
            
            fieldGroupBox.Controls.Add(customTextPanel);
            fieldGroupBox.Controls.Add(fieldListBox);
            fieldPanel.Controls.Add(fieldGroupBox);
            
            // 设计面板
            var designPanel = new Panel { Dock = DockStyle.Fill };
            var designGroupBox = new GroupBox 
            { 
                Text = LanguageManager.GetString("DesignArea"), 
                Dock = DockStyle.Fill, 
                Padding = new Padding(10) 
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

            designGroupBox.Controls.Add(_designPanel);
            designPanel.Controls.Add(designGroupBox);
            
            // 属性面板
            var propertyPanel = new Panel { Dock = DockStyle.Fill };
            var propertyGroupBox = new GroupBox 
            { 
                Text = LanguageManager.GetString("PropertiesAndPreview"), 
                Dock = DockStyle.Fill, 
                Padding = new Padding(10) 
            };
            
            var propertyContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 80,
                ColumnCount = 2,
                RowCount = 2
            };
            
            propertyContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            propertyContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            
            propertyContainer.Controls.Add(new Label { Text = LanguageManager.GetString("TemplateNameProp"), Anchor = AnchorStyles.Left }, 0, 0);
            _templateNameTextBox = new TextBox { Dock = DockStyle.Fill, Text = _currentTemplate.Name };
            propertyContainer.Controls.Add(_templateNameTextBox, 1, 0);
            
            propertyContainer.Controls.Add(new Label { Text = LanguageManager.GetString("OutputFormat"), Anchor = AnchorStyles.Left }, 0, 1);
            _formatComboBox = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _formatComboBox.Items.AddRange(new[] { "Text", "ZPL", "Code128", "QRCode" });
            _formatComboBox.SelectedItem = _currentTemplate.Format.ToString();
            propertyContainer.Controls.Add(_formatComboBox, 1, 1);
            
            // 预览和操作按钮
            var buttonPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                ColumnCount = 2,
                RowCount = 1
            };
            
            _previewButton = new Button { Text = LanguageManager.GetString("Preview"), Dock = DockStyle.Fill };
            _previewButton.Click += PreviewButton_Click;
            buttonPanel.Controls.Add(_previewButton, 0, 0);
            
            var clearButton = new Button { Text = LanguageManager.GetString("Clear"), Dock = DockStyle.Fill };
            clearButton.Click += ClearButton_Click;
            buttonPanel.Controls.Add(clearButton, 1, 0);
            
            _previewTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9F),
                WordWrap = false
            };
            
            propertyGroupBox.Controls.Add(_previewTextBox);
            propertyGroupBox.Controls.Add(buttonPanel);
            propertyGroupBox.Controls.Add(propertyContainer);
            propertyPanel.Controls.Add(propertyGroupBox);
            
            // 底部按钮面板
            var bottomPanel = new Panel { Dock = DockStyle.Fill };
            var bottomButtonPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };
            
            _saveButton = new Button { Text = LanguageManager.GetString("SaveTemplateBtn"), Dock = DockStyle.Fill };
            _saveButton.Click += SaveButton_Click;
            bottomButtonPanel.Controls.Add(_saveButton, 0, 0);
            
            var loadButton = new Button { Text = LanguageManager.GetString("LoadTemplate"), Dock = DockStyle.Fill };
            loadButton.Click += LoadButton_Click;
            bottomButtonPanel.Controls.Add(loadButton, 1, 0);
            
            _cancelButton = new Button { Text = LanguageManager.GetString("Cancel"), Dock = DockStyle.Fill };
            _cancelButton.Click += CancelButton_Click;
            bottomButtonPanel.Controls.Add(_cancelButton, 2, 0);
            
            bottomPanel.Controls.Add(bottomButtonPanel);
            
            // 添加控件到主面板
            mainPanel.Controls.Add(fieldPanel, 0, 0);
            mainPanel.Controls.Add(propertyPanel, 1, 0);
            mainPanel.Controls.Add(designPanel, 0, 1);
            mainPanel.Controls.Add(bottomPanel, 0, 2);
            mainPanel.SetColumnSpan(bottomPanel, 2);
            
            this.Controls.Add(mainPanel);
            
            // 设置事件处理
            fieldListBox.MouseDown += FieldListBox_MouseDown;
            addCustomTextButton.Click += AddCustomText_Click;
            clearDesignButton.Click += ClearDesign_Click;
            
            _designPanel.DragEnter += DesignPanel_DragEnter;
            _designPanel.DragDrop += DesignPanel_DragDrop;
            _designPanel.MouseDown += DesignPanel_MouseDown;
            _designPanel.MouseMove += DesignPanel_MouseMove;
            _designPanel.MouseUp += DesignPanel_MouseUp;
            _designPanel.Paint += DesignPanel_Paint;
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
            
            // 清空现有控件
            ClearButton_Click(null, EventArgs.Empty);
            
            // 解析现有模板内容并创建字段控件
            if (!string.IsNullOrEmpty(_currentTemplate.Content))
            {
                ParseAndCreateFieldControls(_currentTemplate.Content);
            }
        }

        private void ParseAndCreateFieldControls(string content)
        {
            var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int yOffset = 20;
            int lineHeight = 40;
            
            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                if (string.IsNullOrEmpty(trimmedLine))
                    continue;
                
                // 简化解析逻辑：按行处理，每行创建一个控件
                ProcessLine(trimmedLine, yOffset);
                yOffset += lineHeight;
            }
        }

        private void ProcessLine(string line, int yOffset)
        {
            int xOffset = 20;
            
            // 查找所有字段变量
            var fieldMatches = System.Text.RegularExpressions.Regex.Matches(line, @"\{[^}]+\}");
            
            if (fieldMatches.Count == 0)
            {
                // 纯文本行
                CreateFieldControl(line, new Point(xOffset, yOffset), true);
                return;
            }
            
            // 处理包含字段的行
            int lastIndex = 0;
            
            foreach (System.Text.RegularExpressions.Match match in fieldMatches)
            {
                // 处理字段前的文本
                if (match.Index > lastIndex)
                {
                    var beforeText = line.Substring(lastIndex, match.Index - lastIndex);
                    if (!string.IsNullOrEmpty(beforeText.Trim()))
                    {
                        CreateFieldControl(beforeText, new Point(xOffset, yOffset), true);
                        xOffset += GetTextWidth(beforeText) + 10;
                    }
                }
                
                // 处理字段
                var fieldKey = match.Value;
                if (_availableFields.ContainsKey(fieldKey))
                {
                    CreateFieldControl(fieldKey, new Point(xOffset, yOffset), false);
                    xOffset += GetTextWidth(_availableFields[fieldKey]) + 10;
                }
                else
                {
                    // 未知字段，作为自定义文本处理
                    CreateFieldControl(fieldKey, new Point(xOffset, yOffset), true);
                    xOffset += GetTextWidth(fieldKey) + 10;
                }
                
                lastIndex = match.Index + match.Length;
            }
            
            // 处理最后一个字段后的文本
            if (lastIndex < line.Length)
            {
                var afterText = line.Substring(lastIndex);
                if (!string.IsNullOrEmpty(afterText.Trim()))
                {
                    CreateFieldControl(afterText, new Point(xOffset, yOffset), true);
                }
            }
        }

        private int GetTextWidth(string text)
        {
            // 估算文本宽度（可以根据字体大小调整）
            return text.Length * 12; // 假设每个字符宽度为12像素
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
                
                // 将选中的控件置于最顶层
                if (_selectedField != null)
                {
                    _selectedField.BringToFront();
                }
                
                // 高亮选中的字段
                foreach (var control in _fieldControls)
                {
                    control.BackColor = control == _selectedField ? Color.Yellow : (_selectedField?.IsCustomText == true ? Color.LightGreen : Color.LightBlue);
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
            if (_isDragging && _selectedField != null && sender == _selectedField)
            {
                // 计算新位置：当前控件位置 + 鼠标移动的距离
                var deltaX = e.X - _dragStartPoint.X;
                var deltaY = e.Y - _dragStartPoint.Y;
                
                var newLocation = new Point(
                    _selectedField.Location.X + deltaX,
                    _selectedField.Location.Y + deltaY
                );
                
                // 确保不超出设计面板边界
                newLocation.X = Math.Max(0, Math.Min(newLocation.X, _designPanel.Width - _selectedField.Width));
                newLocation.Y = Math.Max(0, Math.Min(newLocation.Y, _designPanel.Height - _selectedField.Height));
                
                _selectedField.Location = newLocation;
            }
        }

        private void FieldControl_MouseUp(object sender, MouseEventArgs e)
        {
            _isDragging = false;
            _selectedField = null;
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
                    // 预定义字段只添加变量引用，不包含项目名称
                    content += field.FieldKey + "\r\n";
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

        private void AddCustomText_Click(object? sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_customTextBox.Text))
            {
                var customText = _customTextBox.Text.Trim();
                // 在设计面板中心位置添加自定义文本
                var centerPoint = new Point(
                    _designPanel.Width / 2 - 60, 
                    _designPanel.Height / 2 - 12
                );
                CreateFieldControl(customText, centerPoint, true);
                _customTextBox.Clear();
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

            // 让Label透传鼠标事件到父控件
            _label.MouseDown += (s, e) => OnMouseDown(e);
            _label.MouseMove += (s, e) => OnMouseMove(e);
            _label.MouseUp += (s, e) => OnMouseUp(e);
            _label.Click += (s, e) => OnClick(e);
            _label.DoubleClick += (s, e) => OnDoubleClick(e);

            this.Controls.Add(_label);
        }
    }
} 