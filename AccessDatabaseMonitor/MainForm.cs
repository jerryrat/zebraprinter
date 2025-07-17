using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AccessDatabaseMonitor
{
    public partial class MainForm : Form
    {
        private DatabaseMonitor? _monitor;
        private Button _selectDbButton;
        private Button _startMonitorButton;
        private Button _stopMonitorButton;
        private Label _statusLabel;
        private Label _dbPathLabel;
        private ListBox _recordsListBox;
        private TextBox _logTextBox;
        private NumericUpDown _intervalNumericUpDown;
        private Label _intervalLabel;
        private GroupBox _controlGroupBox;
        private GroupBox _dataGroupBox;
        private GroupBox _logGroupBox;

        public MainForm()
        {
            InitializeComponent();
            InitializeControls();
        }

        private void InitializeComponent()
        {
            this.Text = "Access Database Monitor - v1.0.0";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
        }

        private void InitializeControls()
        {
            // Control Group Box
            _controlGroupBox = new GroupBox
            {
                Text = "控制面板",
                Location = new Point(10, 10),
                Size = new Size(760, 150),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // Select Database Button
            _selectDbButton = new Button
            {
                Text = "选择数据库文件",
                Location = new Point(10, 25),
                Size = new Size(120, 30)
            };
            _selectDbButton.Click += SelectDbButton_Click;

            // Database Path Label
            _dbPathLabel = new Label
            {
                Text = "未选择数据库文件",
                Location = new Point(140, 30),
                Size = new Size(400, 20),
                ForeColor = Color.Gray
            };

            // Interval Label and NumericUpDown
            _intervalLabel = new Label
            {
                Text = "监控间隔(秒):",
                Location = new Point(10, 65),
                Size = new Size(80, 20)
            };

            _intervalNumericUpDown = new NumericUpDown
            {
                Location = new Point(95, 63),
                Size = new Size(60, 20),
                Minimum = 1,
                Maximum = 60,
                Value = 5
            };

            // Start Monitor Button
            _startMonitorButton = new Button
            {
                Text = "开始监控",
                Location = new Point(170, 60),
                Size = new Size(80, 30),
                Enabled = false
            };
            _startMonitorButton.Click += StartMonitorButton_Click;

            // Stop Monitor Button
            _stopMonitorButton = new Button
            {
                Text = "停止监控",
                Location = new Point(260, 60),
                Size = new Size(80, 30),
                Enabled = false
            };
            _stopMonitorButton.Click += StopMonitorButton_Click;

            // Status Label
            _statusLabel = new Label
            {
                Text = "状态: 未连接",
                Location = new Point(10, 105),
                Size = new Size(400, 20),
                ForeColor = Color.Red
            };

            // Add controls to control group box
            _controlGroupBox.Controls.AddRange(new Control[] {
                _selectDbButton, _dbPathLabel, _intervalLabel, _intervalNumericUpDown,
                _startMonitorButton, _stopMonitorButton, _statusLabel
            });

            // Data Group Box
            _dataGroupBox = new GroupBox
            {
                Text = "监控数据",
                Location = new Point(10, 170),
                Size = new Size(375, 380),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left
            };

            // Records List Box
            _recordsListBox = new ListBox
            {
                Location = new Point(10, 25),
                Size = new Size(355, 345),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Consolas", 9),
                HorizontalScrollbar = true
            };

            _dataGroupBox.Controls.Add(_recordsListBox);

            // Log Group Box
            _logGroupBox = new GroupBox
            {
                Text = "日志信息",
                Location = new Point(395, 170),
                Size = new Size(375, 380),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right
            };

            // Log Text Box
            _logTextBox = new TextBox
            {
                Location = new Point(10, 25),
                Size = new Size(355, 345),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                Font = new Font("Consolas", 9)
            };

            _logGroupBox.Controls.Add(_logTextBox);

            // Add all controls to form
            this.Controls.AddRange(new Control[] {
                _controlGroupBox, _dataGroupBox, _logGroupBox
            });
        }

        private async void SelectDbButton_Click(object? sender, EventArgs e)
        {
            using var openFileDialog = new OpenFileDialog
            {
                Filter = "Access Database Files|*.mdb;*.accdb|All Files|*.*",
                Title = "选择Access数据库文件"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var dbPath = openFileDialog.FileName;
                _dbPathLabel.Text = dbPath;
                _dbPathLabel.ForeColor = Color.Black;

                LogMessage($"选择数据库文件: {dbPath}");

                // Test connection
                _monitor?.Dispose();
                _monitor = new DatabaseMonitor(dbPath);
                _monitor.NewRecordsDetected += OnNewRecordsDetected;
                _monitor.ErrorOccurred += OnErrorOccurred;

                var connectionSuccess = await _monitor.TestConnectionAsync();
                if (connectionSuccess)
                {
                    _statusLabel.Text = "状态: 已连接";
                    _statusLabel.ForeColor = Color.Green;
                    _startMonitorButton.Enabled = true;
                    LogMessage("数据库连接成功");

                    // Initialize and load recent records
                    await _monitor.InitializeAsync();
                    var recentRecords = await _monitor.GetRecentRecordsAsync(20);
                    DisplayRecords(recentRecords, "最近20条记录");
                }
                else
                {
                    _statusLabel.Text = "状态: 连接失败";
                    _statusLabel.ForeColor = Color.Red;
                    _startMonitorButton.Enabled = false;
                    LogMessage("数据库连接失败");
                }
            }
        }

        private void StartMonitorButton_Click(object? sender, EventArgs e)
        {
            if (_monitor != null)
            {
                var interval = (int)_intervalNumericUpDown.Value;
                _monitor.StartMonitoring(interval);
                _startMonitorButton.Enabled = false;
                _stopMonitorButton.Enabled = true;
                _statusLabel.Text = "状态: 监控中...";
                _statusLabel.ForeColor = Color.Blue;
                LogMessage($"开始监控，间隔: {interval}秒");
            }
        }

        private void StopMonitorButton_Click(object? sender, EventArgs e)
        {
            _monitor?.StopMonitoring();
            _startMonitorButton.Enabled = true;
            _stopMonitorButton.Enabled = false;
            _statusLabel.Text = "状态: 已连接";
            _statusLabel.ForeColor = Color.Green;
            LogMessage("停止监控");
        }

        private void OnNewRecordsDetected(List<TestRecord> newRecords)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<List<TestRecord>>(OnNewRecordsDetected), newRecords);
                return;
            }

            DisplayRecords(newRecords, "新增记录");
            LogMessage($"检测到 {newRecords.Count} 条新记录");

            // Show notification
            if (WindowState == FormWindowState.Minimized)
            {
                ShowBalloonTip($"检测到 {newRecords.Count} 条新记录");
            }
        }

        private void OnErrorOccurred(string error)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(OnErrorOccurred), error);
                return;
            }

            LogMessage($"错误: {error}");
            _statusLabel.Text = "状态: 错误";
            _statusLabel.ForeColor = Color.Red;
        }

        private void DisplayRecords(List<TestRecord> records, string category)
        {
            if (records.Any())
            {
                _recordsListBox.Items.Add($"=== {category} - {DateTime.Now:HH:mm:ss} ===");
                foreach (var record in records)
                {
                    _recordsListBox.Items.Add($"  {record}");
                }
                _recordsListBox.Items.Add("");

                // Auto scroll to bottom
                _recordsListBox.TopIndex = _recordsListBox.Items.Count - 1;

                // Limit items count
                while (_recordsListBox.Items.Count > 1000)
                {
                    _recordsListBox.Items.RemoveAt(0);
                }
            }
        }

        private void LogMessage(string message)
        {
            var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            _logTextBox.AppendText(logEntry + Environment.NewLine);

            // Auto scroll to bottom
            _logTextBox.SelectionStart = _logTextBox.Text.Length;
            _logTextBox.ScrollToCaret();

            // Limit log size
            var lines = _logTextBox.Lines;
            if (lines.Length > 1000)
            {
                _logTextBox.Lines = lines.Skip(200).ToArray();
            }
        }

        private void ShowBalloonTip(string message)
        {
            // For simplicity, we'll just flash the window
            // You could implement system tray notifications here if needed
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
                BringToFront();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _monitor?.Dispose();
            base.OnFormClosing(e);
        }
    }
} 