using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Models;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor.Forms
{
    public partial class VisualTemplateEditorForm : Form
    {
        private readonly DragDropTemplateService _templateService;
        private Panel _toolbox;
        private Panel _canvas;
        private Panel _properties;
        private Dictionary<string, Control> _designElements;
        private Control? _selectedElement;
        private bool _isDragging = false;
        private Point _dragStartPoint;
        private ContextMenuStrip _elementContextMenu;

        public VisualTemplateEditorForm(DragDropTemplateService templateService)
        {
            _templateService = templateService;
            _designElements = new Dictionary<string, Control>();
            InitializeComponent();
            InitializeToolbox();
            InitializeCanvas();
            InitializeProperties();
            InitializeContextMenu();
            LoadCurrentTemplate();
        }

        private void InitializeComponent()
        {
            this.Text = "可视化模板编辑器 v1.1.31";
            this.Size = new Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            // 创建主要面板
            _toolbox = new Panel
            {
                Dock = DockStyle.Left,
                Width = 200,
                BackColor = Color.LightGray,
                BorderStyle = BorderStyle.FixedSingle
            };

            _properties = new Panel
            {
                Dock = DockStyle.Right,
                Width = 300,
                BackColor = Color.WhiteSmoke,
                BorderStyle = BorderStyle.FixedSingle
            };

            _canvas = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                AllowDrop = true
            };

            this.Controls.Add(_canvas);
            this.Controls.Add(_properties);
            this.Controls.Add(_toolbox);

            // 添加菜单栏
            var menuStrip = new MenuStrip();
            var fileMenu = new ToolStripMenuItem("文件");
            fileMenu.DropDownItems.Add("新建", null, (s, e) => NewTemplate());
            fileMenu.DropDownItems.Add("打开", null, (s, e) => OpenTemplate());
            fileMenu.DropDownItems.Add("保存", null, (s, e) => SaveTemplate());
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add("退出", null, (s, e) => this.Close());

            var editMenu = new ToolStripMenuItem("编辑");
            editMenu.DropDownItems.Add("复制", null, (s, e) => CopyElement());
            editMenu.DropDownItems.Add("粘贴", null, (s, e) => PasteElement());
            editMenu.DropDownItems.Add("删除", null, (s, e) => DeleteElement());

            menuStrip.Items.Add(fileMenu);
            menuStrip.Items.Add(editMenu);
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }

        private void InitializeToolbox()
        {
            var titleLabel = new Label
            {
                Text = "工具箱",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.DarkGray,
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 10, FontStyle.Bold)
            };
            _toolbox.Controls.Add(titleLabel);

            // 创建工具按钮
            var tools = new[]
            {
                new { Name = "文本", Type = "text", Icon = "T" },
                new { Name = "二维码", Type = "qrcode", Icon = "QR" },
                new { Name = "条形码", Type = "barcode", Icon = "|||" },
                new { Name = "图片", Type = "image", Icon = "IMG" },
                new { Name = "直线", Type = "line", Icon = "—" },
                new { Name = "矩形", Type = "rectangle", Icon = "□" }
            };

            int y = 40;
            foreach (var tool in tools)
            {
                var button = new Button
                {
                    Text = $"{tool.Icon}\n{tool.Name}",
                    Location = new Point(10, y),
                    Size = new Size(180, 50),
                    Tag = tool.Type,
                    BackColor = Color.LightBlue,
                    FlatStyle = FlatStyle.Flat
                };
                button.Click += ToolButton_Click;
                _toolbox.Controls.Add(button);
                y += 60;
            }

            // 添加字段列表
            var fieldsLabel = new Label
            {
                Text = "可用字段",
                Location = new Point(10, y),
                Size = new Size(180, 25),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            _toolbox.Controls.Add(fieldsLabel);

            var fieldsListBox = new ListBox
            {
                Location = new Point(10, y + 30),
                Size = new Size(180, 200)
            };
            
            // 添加可用字段
            var fields = PrintTemplateManager.GetAvailableFields();
            foreach (var field in fields)
            {
                fieldsListBox.Items.Add(field);
            }
            fieldsListBox.DoubleClick += (s, e) =>
            {
                if (fieldsListBox.SelectedItem != null)
                {
                    CreateTextElement(fieldsListBox.SelectedItem.ToString()!);
                }
            };
            _toolbox.Controls.Add(fieldsListBox);
        }

        private void InitializeCanvas()
        {
            _canvas.Paint += Canvas_Paint;
            _canvas.MouseDown += Canvas_MouseDown;
            _canvas.MouseMove += Canvas_MouseMove;
            _canvas.MouseUp += Canvas_MouseUp;
            _canvas.DragEnter += Canvas_DragEnter;
            _canvas.DragDrop += Canvas_DragDrop;
        }

        private void InitializeProperties()
        {
            var titleLabel = new Label
            {
                Text = "属性面板",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.DarkGray,
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 10, FontStyle.Bold)
            };
            _properties.Controls.Add(titleLabel);
        }

        private void InitializeContextMenu()
        {
            _elementContextMenu = new ContextMenuStrip();
            _elementContextMenu.Items.Add("复制", null, (s, e) => CopyElement());
            _elementContextMenu.Items.Add("删除", null, (s, e) => DeleteElement());
            _elementContextMenu.Items.Add(new ToolStripSeparator());
            _elementContextMenu.Items.Add("置于顶层", null, (s, e) => BringToFront());
            _elementContextMenu.Items.Add("置于底层", null, (s, e) => SendToBack());
        }

        private void ToolButton_Click(object? sender, EventArgs e)
        {
            if (sender is Button button && button.Tag is string toolType)
            {
                CreateAndAddElement(toolType, new Point(100, 100));
            }
        }

        private void CreateAndAddElement(string type, Point location)
        {
            Control element = type switch
            {
                "text" => CreateTextElement("文本"),
                "qrcode" => CreateQRCodeElement(),
                "barcode" => CreateBarcodeElement(),
                "image" => CreateImageElement(),
                "line" => CreateLineElement(),
                "rectangle" => CreateRectangleElement(),
                _ => CreateTextElement("未知元素")
            };

            element.Location = location;
            element.Tag = new DragDropTemplateService.PrintElementTemplate
            {
                Id = Guid.NewGuid().ToString(),
                Type = type,
                X = location.X,
                Y = location.Y,
                Width = element.Width,
                Height = element.Height,
                Content = element.Text
            };

            _canvas.Controls.Add(element);
            _designElements[((DragDropTemplateService.PrintElementTemplate)element.Tag).Id] = element;
            SelectElement(element);
        }

        private Control CreateTextElement(string text)
        {
            var label = new Label
            {
                Text = text,
                Size = new Size(100, 30),
                BackColor = Color.LightYellow,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.SizeAll
            };
            SetupElementEvents(label);
            return label;
        }

        private Control CreateQRCodeElement()
        {
            var panel = new Panel
            {
                Size = new Size(100, 100),
                BackColor = Color.LightGreen,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.SizeAll
            };
            var label = new Label
            {
                Text = "QR",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 12, FontStyle.Bold)
            };
            panel.Controls.Add(label);
            SetupElementEvents(panel);
            return panel;
        }

        private Control CreateBarcodeElement()
        {
            var panel = new Panel
            {
                Size = new Size(150, 50),
                BackColor = Color.LightCyan,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.SizeAll
            };
            var label = new Label
            {
                Text = "|||||||||||",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 10)
            };
            panel.Controls.Add(label);
            SetupElementEvents(panel);
            return panel;
        }

        private Control CreateImageElement()
        {
            var panel = new Panel
            {
                Size = new Size(100, 100),
                BackColor = Color.LightPink,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.SizeAll
            };
            var label = new Label
            {
                Text = "图片",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(label);
            SetupElementEvents(panel);
            return panel;
        }

        private Control CreateLineElement()
        {
            var panel = new Panel
            {
                Size = new Size(100, 3),
                BackColor = Color.Black,
                Cursor = Cursors.SizeAll
            };
            SetupElementEvents(panel);
            return panel;
        }

        private Control CreateRectangleElement()
        {
            var panel = new Panel
            {
                Size = new Size(100, 60),
                BackColor = Color.Transparent,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.SizeAll
            };
            SetupElementEvents(panel);
            return panel;
        }

        private void SetupElementEvents(Control element)
        {
            element.MouseDown += Element_MouseDown;
            element.MouseMove += Element_MouseMove;
            element.MouseUp += Element_MouseUp;
            element.ContextMenuStrip = _elementContextMenu;
            element.DoubleClick += Element_DoubleClick;
        }

        private void Element_MouseDown(object? sender, MouseEventArgs e)
        {
            if (sender is Control element)
            {
                SelectElement(element);
                _isDragging = true;
                _dragStartPoint = e.Location;
            }
        }

        private void Element_MouseMove(object? sender, MouseEventArgs e)
        {
            if (_isDragging && sender is Control element && e.Button == MouseButtons.Left)
            {
                var newLocation = new Point(
                    element.Left + (e.X - _dragStartPoint.X),
                    element.Top + (e.Y - _dragStartPoint.Y)
                );
                element.Location = newLocation;
                
                // 更新元素数据
                if (element.Tag is DragDropTemplateService.PrintElementTemplate template)
                {
                    template.X = newLocation.X;
                    template.Y = newLocation.Y;
                }
                UpdatePropertyPanel();
            }
        }

        private void Element_MouseUp(object? sender, MouseEventArgs e)
        {
            _isDragging = false;
        }

        private void Element_DoubleClick(object? sender, EventArgs e)
        {
            if (sender is Control element && element.Tag is DragDropTemplateService.PrintElementTemplate template)
            {
                var dialog = new TextInputDialog("编辑内容", "请输入新内容:", template.Content);
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    template.Content = dialog.InputText;
                    if (element is Label label)
                    {
                        label.Text = template.Content;
                    }
                    UpdatePropertyPanel();
                }
            }
        }

        private void SelectElement(Control element)
        {
            // 清除其他选择
            foreach (Control ctrl in _canvas.Controls)
            {
                ctrl.BackColor = GetOriginalBackColor(ctrl);
            }

            // 选择当前元素
            _selectedElement = element;
            element.BackColor = Color.Yellow;
            UpdatePropertyPanel();
        }

        private Color GetOriginalBackColor(Control control)
        {
            if (control.Tag is DragDropTemplateService.PrintElementTemplate template)
            {
                return template.Type switch
                {
                    "text" => Color.LightYellow,
                    "qrcode" => Color.LightGreen,
                    "barcode" => Color.LightCyan,
                    "image" => Color.LightPink,
                    "line" => Color.Black,
                    "rectangle" => Color.Transparent,
                    _ => Color.White
                };
            }
            return Color.White;
        }

        private void UpdatePropertyPanel()
        {
            // 清除现有控件
            for (int i = _properties.Controls.Count - 1; i >= 1; i--)
            {
                _properties.Controls.RemoveAt(i);
            }

            if (_selectedElement?.Tag is DragDropTemplateService.PrintElementTemplate template)
            {
                int y = 40;
                
                // 添加属性控件
                AddPropertyTextBox("X位置:", template.X.ToString(), (value) => {
                    if (double.TryParse(value, out double x))
                    {
                        template.X = x;
                        _selectedElement.Left = (int)x;
                    }
                }, ref y);

                AddPropertyTextBox("Y位置:", template.Y.ToString(), (value) => {
                    if (double.TryParse(value, out double y))
                    {
                        template.Y = y;
                        _selectedElement.Top = (int)y;
                    }
                }, ref y);

                AddPropertyTextBox("宽度:", template.Width.ToString(), (value) => {
                    if (double.TryParse(value, out double width))
                    {
                        template.Width = width;
                        _selectedElement.Width = (int)width;
                    }
                }, ref y);

                AddPropertyTextBox("高度:", template.Height.ToString(), (value) => {
                    if (double.TryParse(value, out double height))
                    {
                        template.Height = height;
                        _selectedElement.Height = (int)height;
                    }
                }, ref y);

                AddPropertyTextBox("内容:", template.Content, (value) => {
                    template.Content = value;
                    if (_selectedElement is Label label)
                    {
                        label.Text = value;
                    }
                }, ref y);
            }
        }

        private void AddPropertyTextBox(string labelText, string value, Action<string> onValueChanged, ref int y)
        {
            var label = new Label
            {
                Text = labelText,
                Location = new Point(10, y),
                Size = new Size(80, 20)
            };
            _properties.Controls.Add(label);

            var textBox = new TextBox
            {
                Text = value,
                Location = new Point(100, y),
                Size = new Size(180, 20)
            };
            textBox.TextChanged += (s, e) => onValueChanged(textBox.Text);
            _properties.Controls.Add(textBox);

            y += 30;
        }

        private void Canvas_Paint(object? sender, PaintEventArgs e)
        {
            // 绘制网格
            var graphics = e.Graphics;
            var gridPen = new Pen(Color.LightGray, 1);
            
            for (int x = 0; x < _canvas.Width; x += 20)
            {
                graphics.DrawLine(gridPen, x, 0, x, _canvas.Height);
            }
            
            for (int y = 0; y < _canvas.Height; y += 20)
            {
                graphics.DrawLine(gridPen, 0, y, _canvas.Width, y);
            }
            
            gridPen.Dispose();
        }

        private void Canvas_MouseDown(object? sender, MouseEventArgs e)
        {
            // 点击空白区域取消选择
            _selectedElement = null;
            foreach (Control ctrl in _canvas.Controls)
            {
                ctrl.BackColor = GetOriginalBackColor(ctrl);
            }
            UpdatePropertyPanel();
        }

        private void Canvas_MouseMove(object? sender, MouseEventArgs e)
        {
            // 显示鼠标位置
            this.Text = $"可视化模板编辑器 v1.1.31 - 位置: ({e.X}, {e.Y})";
        }

        private void Canvas_MouseUp(object? sender, MouseEventArgs e)
        {
            // 处理画布鼠标释放事件
        }

        private void Canvas_DragEnter(object? sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void Canvas_DragDrop(object? sender, DragEventArgs e)
        {
            var point = _canvas.PointToClient(new Point(e.X, e.Y));
            if (e.Data!.GetDataPresent(DataFormats.Text))
            {
                var text = e.Data.GetData(DataFormats.Text)?.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    CreateAndAddElement("text", point);
                }
            }
        }

        private void NewTemplate()
        {
            _canvas.Controls.Clear();
            _designElements.Clear();
            _selectedElement = null;
            UpdatePropertyPanel();
        }

        private void OpenTemplate()
        {
            var openDialog = new OpenFileDialog
            {
                Filter = "模板文件 (*.json)|*.json",
                Title = "打开模板"
            };
            
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 这里可以添加加载模板的逻辑
                    MessageBox.Show("模板加载功能待实现", "提示");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载模板失败: {ex.Message}", "错误");
                }
            }
        }

        private void SaveTemplate()
        {
            var saveDialog = new SaveFileDialog
            {
                Filter = "模板文件 (*.json)|*.json",
                Title = "保存模板"
            };
            
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 这里可以添加保存模板的逻辑
                    var templates = new Dictionary<string, DragDropTemplateService.PrintElementTemplate>();
                    foreach (var kvp in _designElements)
                    {
                        if (kvp.Value.Tag is DragDropTemplateService.PrintElementTemplate template)
                        {
                            templates[kvp.Key] = template;
                        }
                    }
                    
                    // 调用服务保存模板
                    var templateName = System.IO.Path.GetFileNameWithoutExtension(saveDialog.FileName);
                    _templateService.SaveTemplate(templateName, templates);
                    
                    MessageBox.Show("模板保存成功！", "成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"保存模板失败: {ex.Message}", "错误");
                }
            }
        }

        private void LoadCurrentTemplate()
        {
            try
            {
                var currentTemplate = _templateService.GetCurrentTemplate();
                foreach (var kvp in currentTemplate)
                {
                    var template = kvp.Value;
                    Control element = template.Type switch
                    {
                        "text" => CreateTextElement(template.Content),
                        "qrcode" => CreateQRCodeElement(),
                        "barcode" => CreateBarcodeElement(),
                        "image" => CreateImageElement(),
                        "line" => CreateLineElement(),
                        "rectangle" => CreateRectangleElement(),
                        _ => CreateTextElement("未知元素")
                    };

                    element.Location = new Point((int)template.X, (int)template.Y);
                    element.Size = new Size((int)template.Width, (int)template.Height);
                    element.Tag = template;

                    _canvas.Controls.Add(element);
                    _designElements[template.Id] = element;
                    SetupElementEvents(element);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"加载当前模板失败: {ex.Message}", ex);
            }
        }

        private void CopyElement()
        {
            // 复制功能待实现
            MessageBox.Show("复制功能待实现", "提示");
        }

        private void PasteElement()
        {
            // 粘贴功能待实现
            MessageBox.Show("粘贴功能待实现", "提示");
        }

        private void DeleteElement()
        {
            if (_selectedElement != null)
            {
                if (_selectedElement.Tag is DragDropTemplateService.PrintElementTemplate template)
                {
                    _designElements.Remove(template.Id);
                }
                _canvas.Controls.Remove(_selectedElement);
                _selectedElement = null;
                UpdatePropertyPanel();
            }
        }

        private new void BringToFront()
        {
            _selectedElement?.BringToFront();
        }

        private new void SendToBack()
        {
            _selectedElement?.SendToBack();
        }
    }

    // 简单的文本输入对话框
    public class TextInputDialog : Form
    {
        public string InputText { get; private set; } = "";

        public TextInputDialog(string title, string prompt, string defaultText = "")
        {
            this.Text = title;
            this.Size = new Size(400, 150);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var label = new Label
            {
                Text = prompt,
                Location = new Point(12, 15),
                Size = new Size(360, 20)
            };

            var textBox = new TextBox
            {
                Text = defaultText,
                Location = new Point(12, 40),
                Size = new Size(360, 20)
            };

            var okButton = new Button
            {
                Text = "确定",
                Location = new Point(215, 75),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };

            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(297, 75),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            okButton.Click += (s, e) => {
                InputText = textBox.Text;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }
    }
} 