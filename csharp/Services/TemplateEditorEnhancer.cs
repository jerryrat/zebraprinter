using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.Generic; // Added for List

namespace ZebraPrinterMonitor.Services
{
    /// <summary>
    /// 模板编辑器增强功能
    /// 提供语法高亮、智能提示等功能
    /// </summary>
    public class TemplateEditorEnhancer
    {
        private RichTextBox _textBox;
        private bool _isUpdating = false;
        
        public TemplateEditorEnhancer(RichTextBox textBox)
        {
            _textBox = textBox;
            _textBox.TextChanged += OnTextChanged;
            _textBox.SelectionChanged += OnSelectionChanged;
        }

        /// <summary>
        /// 应用语法高亮
        /// </summary>
        public void ApplySyntaxHighlighting()
        {
            if (_isUpdating) return;
            
            _isUpdating = true;
            
            try
            {
                int currentSelectionStart = _textBox.SelectionStart;
                int currentSelectionLength = _textBox.SelectionLength;
                
                // 保存当前滚动位置
                int scrollPos = GetScrollPos();
                
                // 重置所有文本格式
                _textBox.SelectAll();
                _textBox.SelectionColor = Color.Black;
                _textBox.SelectionBackColor = Color.White;
                _textBox.SelectionFont = new Font(_textBox.Font, FontStyle.Regular);
                
                // 高亮模板变量 {Variable}
                HighlightPattern(@"\{[A-Za-z0-9_]+\}", Color.Blue, FontStyle.Bold);
                
                // 高亮注释行 (以 // 开头)
                HighlightPattern(@"//.*?(?=\r|\n|$)", Color.Green, FontStyle.Italic);
                
                // 高亮特殊字符和符号
                HighlightPattern(@"[+\-|=:]", Color.Brown, FontStyle.Regular);
                
                // 高亮数字
                HighlightPattern(@"\b\d+\.?\d*\b", Color.Purple, FontStyle.Regular);
                
                // 恢复光标位置
                _textBox.SelectionStart = currentSelectionStart;
                _textBox.SelectionLength = currentSelectionLength;
                
                // 恢复滚动位置
                SetScrollPos(scrollPos);
            }
            finally
            {
                _isUpdating = false;
            }
        }
        
        private void HighlightPattern(string pattern, Color color, FontStyle style)
        {
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            MatchCollection matches = regex.Matches(_textBox.Text);
            
            foreach (Match match in matches)
            {
                _textBox.SelectionStart = match.Index;
                _textBox.SelectionLength = match.Length;
                _textBox.SelectionColor = color;
                _textBox.SelectionFont = new Font(_textBox.Font, style);
            }
        }
        
        private void OnTextChanged(object sender, EventArgs e)
        {
            // 延迟应用语法高亮，避免频繁更新
            var timer = new System.Windows.Forms.Timer();
            timer.Interval = 500; // 500ms延迟
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                timer.Dispose();
                ApplySyntaxHighlighting();
            };
            timer.Start();
        }
        
        private void OnSelectionChanged(object sender, EventArgs e)
        {
            // 可以在这里添加智能提示功能
            ShowVariableTooltip();
        }
        
        private void ShowVariableTooltip()
        {
            // 检查光标位置是否在变量上
            int cursorPos = _textBox.SelectionStart;
            string text = _textBox.Text;
            
            // 查找当前位置的变量
            Regex regex = new Regex(@"\{([A-Za-z0-9_]+)\}");
            MatchCollection matches = regex.Matches(text);
            
            foreach (Match match in matches)
            {
                if (cursorPos >= match.Index && cursorPos <= match.Index + match.Length)
                {
                    string variableName = match.Groups[1].Value;
                    ShowTooltip(variableName, GetVariableDescription(variableName));
                    return;
                }
            }
        }
        
        private string GetVariableDescription(string variableName)
        {
            var descriptions = PrintTemplateManager.GetFieldDescriptions();
            string key = "{" + variableName + "}";
            return descriptions.ContainsKey(key) ? descriptions[key] : "未知变量";
        }
        
        private void ShowTooltip(string variable, string description)
        {
            // 这里可以实现工具提示显示
            // 可以使用ToolTip控件或自定义提示窗口
        }
        
        #region 滚动位置保存/恢复
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int GetScrollPos(IntPtr hWnd, int nBar);
        
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SetScrollPos(IntPtr hWnd, int nBar, int nPos, bool bRedraw);
        
        private int GetScrollPos()
        {
            return GetScrollPos(_textBox.Handle, 1); // SB_VERT = 1
        }
        
        private void SetScrollPos(int pos)
        {
            SetScrollPos(_textBox.Handle, 1, pos, true);
        }
        #endregion
        
        /// <summary>
        /// 插入变量到当前光标位置
        /// </summary>
        public void InsertVariable(string variableName)
        {
            int cursorPos = _textBox.SelectionStart;
            string variable = "{" + variableName + "}";
            
            _textBox.Text = _textBox.Text.Insert(cursorPos, variable);
            _textBox.SelectionStart = cursorPos + variable.Length;
            _textBox.Focus();
        }
        
        /// <summary>
        /// 获取所有可用的变量列表
        /// </summary>
        public string[] GetAvailableVariables()
        {
            var fields = PrintTemplateManager.GetAvailableFields();
            return fields.ToArray();
        }
        
        /// <summary>
        /// 验证模板语法
        /// </summary>
        public TemplateValidationResult ValidateTemplate()
        {
            var result = new TemplateValidationResult();
            string text = _textBox.Text;
            
            // 检查未闭合的大括号
            int openBraces = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '{') openBraces++;
                else if (text[i] == '}') openBraces--;
                
                if (openBraces < 0)
                {
                    result.Errors.Add($"第 {GetLineNumber(i)} 行: 未匹配的右大括号");
                    openBraces = 0;
                }
            }
            
            if (openBraces > 0)
            {
                result.Errors.Add("模板中存在未闭合的左大括号");
            }
            
            // 检查无效的变量
            Regex regex = new Regex(@"\{([A-Za-z0-9_]+)\}");
            MatchCollection matches = regex.Matches(text);
            var validFields = PrintTemplateManager.GetAvailableFields();
            
            foreach (Match match in matches)
            {
                string variable = match.Value;
                if (!validFields.Contains(variable))
                {
                    int lineNum = GetLineNumber(match.Index);
                    result.Warnings.Add($"第 {lineNum} 行: 未知变量 '{variable}'");
                }
            }
            
            result.IsValid = result.Errors.Count == 0;
            return result;
        }
        
        private int GetLineNumber(int position)
        {
            return _textBox.Text.Substring(0, position).Split('\n').Length;
        }
    }
    
    /// <summary>
    /// 模板验证结果
    /// </summary>
    public class TemplateValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        
        public string GetSummary()
        {
            if (IsValid && Warnings.Count == 0)
                return "✅ 模板语法正确";
            
            var summary = IsValid ? "⚠️ 模板可用但有警告" : "❌ 模板有错误";
            
            if (Errors.Count > 0)
                summary += $"\n错误 ({Errors.Count}): " + string.Join(", ", Errors);
            
            if (Warnings.Count > 0)
                summary += $"\n警告 ({Warnings.Count}): " + string.Join(", ", Warnings);
            
            return summary;
        }
    }
} 