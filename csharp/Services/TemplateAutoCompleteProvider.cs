using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ZebraPrinterMonitor.Services
{
    /// <summary>
    /// 模板自动完成和智能提示服务
    /// </summary>
    public class TemplateAutoCompleteProvider
    {
        private readonly RichTextBox _textBox;
        private readonly ListBox _suggestionBox;
        private readonly Form _parentForm;
        private readonly Dictionary<string, string> _fieldDescriptions;
        private readonly List<string> _availableFields;
        private readonly List<TemplateSnippet> _snippets = new();
        private bool _isShowingSuggestions = false;

        public TemplateAutoCompleteProvider(RichTextBox textBox, Form parentForm)
        {
            _textBox = textBox;
            _parentForm = parentForm;
            _fieldDescriptions = PrintTemplateManager.GetFieldDescriptions();
            _availableFields = PrintTemplateManager.GetAvailableFields();
            
            // 创建建议列表框
            _suggestionBox = new ListBox
            {
                Visible = false,
                Font = new Font("Consolas", 9),
                BackColor = Color.LightYellow,
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                Height = 150,
                Width = 300
            };
            
            _parentForm.Controls.Add(_suggestionBox);
            _suggestionBox.BringToFront();
            
            InitializeSnippets();
            SetupEventHandlers();
        }

        private void InitializeSnippets()
        {
            _snippets.AddRange(new List<TemplateSnippet>
            {
                new TemplateSnippet("基本太阳能板规格", "basic_spec", @"型号: SKT 600 M12/120HB
额定最大功率: {Power} W
最大功率电压: {VoltageVpm} V
最大功率电流: {CurrentImp} A
开路电压: {Voltage} V
短路电流: {Current} A"),

                new TemplateSnippet("完整规格表", "full_spec", @"+---------------------------------------------------------------+
|                    SOLAR PANEL SPECIFICATION                 |
+--------------------------------------+----------------------------+
| Model Type                           | SKT 600 M12/120HB         |
| Rated Maximum Power      (Pmax)      | {Power} W                  |
| Voltage at Pmax          (Vmp)       | {VoltageVpm} V             |
| Current at Pmax          (Imp)       | {CurrentImp} A             |
| Open-Circuit Voltage     (Voc)       | {Voltage} V                |
| Short-Circuit Current    (Isc)       | {Current} A                |
| PV Module Classification             | CLASS II                   |
| Maximum System Voltage               | 1500 V                     |
| Maximum Series Fuse Rating           | 35 A                       |
| Operating Temperature                | -40~85°C                   |
| Dimensions(mm)                       | 2172x1303x40(mm)           |
| Pmax/Voc/Isc Tolerance               | ±3%                        |
+--------------------------------------+----------------------------+"),

                new TemplateSnippet("简洁规格", "simple_spec", @"序列号: {SerialNumber}
测试时间: {TestDateTime}
最大功率: {Power}W
开路电压: {Voltage}V / 短路电流: {Current}A
最大功率电压: {VoltageVpm}V / 最大功率电流: {CurrentImp}A"),

                new TemplateSnippet("表格边框", "table_border", @"+---------------------------------------------------------------+
|                                                               |
+---------------------------------------------------------------+"),

                new TemplateSnippet("测试条件说明", "test_conditions", @"测试条件: STC (Standard Test Conditions)
- 辐照度: 1000W/m²
- 光谱: AM1.5
- 电池温度: 25°C"),

                new TemplateSnippet("日期时间格式", "datetime_format", @"测试日期: {TestDateTime}
打印时间: {CurrentTime}")
            });
        }

        private void SetupEventHandlers()
        {
            _textBox.KeyDown += TextBox_KeyDown;
            _textBox.KeyPress += TextBox_KeyPress;
            _textBox.TextChanged += TextBox_TextChanged;
            _textBox.LostFocus += TextBox_LostFocus;
            
            _suggestionBox.KeyDown += SuggestionBox_KeyDown;
            _suggestionBox.DoubleClick += SuggestionBox_DoubleClick;
            _suggestionBox.Click += SuggestionBox_Click;
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (_isShowingSuggestions)
            {
                switch (e.KeyCode)
                {
                    case Keys.Escape:
                        HideSuggestions();
                        e.Handled = true;
                        break;
                        
                    case Keys.Up:
                        if (_suggestionBox.SelectedIndex > 0)
                            _suggestionBox.SelectedIndex--;
                        e.Handled = true;
                        break;
                        
                    case Keys.Down:
                        if (_suggestionBox.SelectedIndex < _suggestionBox.Items.Count - 1)
                            _suggestionBox.SelectedIndex++;
                        e.Handled = true;
                        break;
                        
                    case Keys.Enter:
                    case Keys.Tab:
                        AcceptSuggestion();
                        e.Handled = true;
                        break;
                }
            }
            else
            {
                // Ctrl+Space 手动触发自动完成
                if (e.Control && e.KeyCode == Keys.Space)
                {
                    ShowSuggestions();
                    e.Handled = true;
                }
            }
        }

        private void TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '{' || e.KeyChar == '}')
            {
                // 延迟显示建议，避免影响输入
                var timer = new System.Windows.Forms.Timer { Interval = 100 };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    CheckForAutoComplete();
                };
                timer.Start();
            }
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            if (_isShowingSuggestions)
            {
                UpdateSuggestions();
            }
        }

        private void TextBox_LostFocus(object sender, EventArgs e)
        {
            // 延迟隐藏，允许用户点击建议列表
            var timer = new System.Windows.Forms.Timer { Interval = 200 };
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                timer.Dispose();
                if (!_suggestionBox.Focused)
                {
                    HideSuggestions();
                }
            };
            timer.Start();
        }

        private void SuggestionBox_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    HideSuggestions();
                    _textBox.Focus();
                    e.Handled = true;
                    break;
                    
                case Keys.Enter:
                    AcceptSuggestion();
                    e.Handled = true;
                    break;
            }
        }

        private void SuggestionBox_DoubleClick(object sender, EventArgs e)
        {
            AcceptSuggestion();
        }

        private void SuggestionBox_Click(object sender, EventArgs e)
        {
            AcceptSuggestion();
        }

        private void CheckForAutoComplete()
        {
            int cursorPos = _textBox.SelectionStart;
            string text = _textBox.Text;
            
            // 检查是否在输入变量
            if (cursorPos > 0 && text[cursorPos - 1] == '{')
            {
                ShowVariableSuggestions();
            }
            // 检查是否在输入片段
            else if (IsAtWordBoundary(cursorPos))
            {
                string currentWord = GetCurrentWord(cursorPos);
                if (!string.IsNullOrEmpty(currentWord))
                {
                    ShowSnippetSuggestions(currentWord);
                }
            }
        }

        private void ShowSuggestions()
        {
            CheckForAutoComplete();
        }

        private void ShowVariableSuggestions()
        {
            _suggestionBox.Items.Clear();
            
            foreach (var field in _availableFields)
            {
                string description = _fieldDescriptions.ContainsKey(field) ? _fieldDescriptions[field] : "";
                _suggestionBox.Items.Add(new SuggestionItem(field, description, SuggestionType.Variable));
            }
            
            if (_suggestionBox.Items.Count > 0)
            {
                ShowSuggestionBox();
            }
        }

        private void ShowSnippetSuggestions(string prefix)
        {
            _suggestionBox.Items.Clear();
            
            var matchingSnippets = _snippets
                .Where(s => s.Name.ToLower().Contains(prefix.ToLower()) || 
                           s.Trigger.ToLower().StartsWith(prefix.ToLower()))
                .ToList();
            
            foreach (var snippet in matchingSnippets)
            {
                _suggestionBox.Items.Add(new SuggestionItem(snippet.Name, snippet.Trigger, SuggestionType.Snippet, snippet));
            }
            
            if (_suggestionBox.Items.Count > 0)
            {
                ShowSuggestionBox();
            }
        }

        private void UpdateSuggestions()
        {
            if (!_isShowingSuggestions) return;
            
            int cursorPos = _textBox.SelectionStart;
            string currentWord = GetCurrentWord(cursorPos);
            
            // 过滤现有建议
            var filteredItems = new List<SuggestionItem>();
            foreach (SuggestionItem item in _suggestionBox.Items)
            {
                if (string.IsNullOrEmpty(currentWord) || 
                    item.DisplayText.ToLower().Contains(currentWord.ToLower()))
                {
                    filteredItems.Add(item);
                }
            }
            
            if (filteredItems.Count == 0)
            {
                HideSuggestions();
            }
            else
            {
                _suggestionBox.Items.Clear();
                foreach (var item in filteredItems)
                {
                    _suggestionBox.Items.Add(item);
                }
                _suggestionBox.SelectedIndex = 0;
            }
        }

        private void ShowSuggestionBox()
        {
            if (_suggestionBox.Items.Count == 0) return;
            
            // 计算建议框位置
            Point caretPos = GetCaretPosition();
            _suggestionBox.Location = new Point(
                Math.Min(caretPos.X, _parentForm.Width - _suggestionBox.Width - 10),
                Math.Min(caretPos.Y + 20, _parentForm.Height - _suggestionBox.Height - 50)
            );
            
            _suggestionBox.SelectedIndex = 0;
            _suggestionBox.Visible = true;
            _suggestionBox.BringToFront();
            _isShowingSuggestions = true;
        }

        private void HideSuggestions()
        {
            _suggestionBox.Visible = false;
            _isShowingSuggestions = false;
        }

        private void AcceptSuggestion()
        {
            if (_suggestionBox.SelectedItem is SuggestionItem item)
            {
                int cursorPos = _textBox.SelectionStart;
                
                if (item.Type == SuggestionType.Variable)
                {
                    // 插入变量（移除开头的大括号，因为用户已经输入了）
                    string variable = item.DisplayText.Substring(1); // 移除开头的 {
                    _textBox.Text = _textBox.Text.Insert(cursorPos, variable);
                    _textBox.SelectionStart = cursorPos + variable.Length;
                }
                else if (item.Type == SuggestionType.Snippet && item.Snippet != null)
                {
                    // 插入代码片段
                    string wordStart = GetWordStart(cursorPos);
                    int replaceStart = cursorPos - wordStart.Length;
                    int replaceLength = wordStart.Length;
                    
                    _textBox.Text = _textBox.Text.Remove(replaceStart, replaceLength);
                    _textBox.Text = _textBox.Text.Insert(replaceStart, item.Snippet.Content);
                    _textBox.SelectionStart = replaceStart + item.Snippet.Content.Length;
                }
                
                HideSuggestions();
                _textBox.Focus();
            }
        }

        private Point GetCaretPosition()
        {
            var caretPos = _textBox.GetPositionFromCharIndex(_textBox.SelectionStart);
            return _textBox.PointToScreen(caretPos);
        }

        private string GetCurrentWord(int position)
        {
            string text = _textBox.Text;
            int start = position;
            
            // 向后查找单词开始
            while (start > 0 && (char.IsLetterOrDigit(text[start - 1]) || text[start - 1] == '_'))
            {
                start--;
            }
            
            return text.Substring(start, position - start);
        }

        private string GetWordStart(int position)
        {
            return GetCurrentWord(position);
        }

        private bool IsAtWordBoundary(int position)
        {
            if (position == 0) return true;
            
            string text = _textBox.Text;
            char prevChar = text[position - 1];
            
            return char.IsWhiteSpace(prevChar) || char.IsPunctuation(prevChar) || prevChar == '\n' || prevChar == '\r';
        }

        public void Dispose()
        {
            _suggestionBox?.Dispose();
        }
    }

    /// <summary>
    /// 建议项
    /// </summary>
    public class SuggestionItem
    {
        public string DisplayText { get; }
        public string Description { get; }
        public SuggestionType Type { get; }
        public TemplateSnippet? Snippet { get; }

        public SuggestionItem(string displayText, string description, SuggestionType type, TemplateSnippet? snippet = null)
        {
            DisplayText = displayText;
            Description = description;
            Type = type;
            Snippet = snippet;
        }

        public override string ToString()
        {
            return $"{DisplayText} - {Description}";
        }
    }

    /// <summary>
    /// 建议类型
    /// </summary>
    public enum SuggestionType
    {
        Variable,
        Snippet
    }

    /// <summary>
    /// 模板代码片段
    /// </summary>
    public class TemplateSnippet
    {
        public string Name { get; }
        public string Trigger { get; }
        public string Content { get; }

        public TemplateSnippet(string name, string trigger, string content)
        {
            Name = name;
            Trigger = trigger;
            Content = content;
        }
    }
} 