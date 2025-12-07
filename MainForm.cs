using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Text.Json;

namespace SapiReader;

public partial class MainForm : Form
{
    #region Win32 API

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const int HOTKEY_ID = 0x1234;
    private const int WM_HOTKEY = 0x0312;

    private const int KEYEVENTF_KEYDOWN = 0x0000;
    private const int KEYEVENTF_KEYUP = 0x0002;
    private const byte VK_CONTROL = 0x11;
    private const byte VK_C = 0x43;

    #endregion

    private SpeechSynthesizer _voice = null!;
    private NotifyIcon _trayIcon = null!;
    private bool _isSpeaking = false;
    private string? _previousClipboard = null;
    private List<InstalledVoice> _installedVoices = new();

    // 朗读记录（最多1000条，超出覆盖）
    private const int MAX_HISTORY = 1000;
    private List<ReadingRecord> _readingHistory = new();

    // 当前热键配置（无修饰键）
    private uint _hotkeyKey = 0x70; // F1
    private uint _stopHotkeyKey = 0x71; // F2
    private const int STOP_HOTKEY_ID = 2;

    // UI 控件
    private ComboBox _keyCombo = null!;
    private ComboBox _stopKeyCombo = null!;
    private Button _applyButton = null!;
    private Button _topMostButton = null!;
    private Label _statusLabel = null!;
    private TextBox _rateTextBox = null!;
    private TextBox _volumeTextBox = null!;
    private Label _rateLabel = null!;
    private Label _volumeLabel = null!;
    private ComboBox _voiceCombo = null!;
    private Button _refreshVoiceButton = null!;
    private ListBox _historyListBox = null!;
    private Button _clearHistoryButton = null!;
    private TextBox _escapeRulesTextBox = null!;
    private Button _saveConfigButton = null!;

    // 转义字符规则
    private Dictionary<string, string> _escapeRules = new();

    // 配置文件路径 - 使用项目根目录
    private string _configPath = "";
    private bool _isLoadingConfig = false;

    public MainForm()
    {
        // 计算配置文件路径（单文件发布时使用 AppContext.BaseDirectory）
        var exeDir = AppContext.BaseDirectory;
        _configPath = Path.Combine(exeDir, "config.json");
        
        _isLoadingConfig = true; // 在初始化之前设置为 true
        
        InitializeComponent();
        InitializeVoice();
        InitializeTrayIcon();
        LoadConfig();
        RegisterCurrentHotkey();
        
        _isLoadingConfig = false; // 初始化完成后设置为 false
    }

    private void InitializeComponent()
    {
        this.Text = "SAPI5 文本朗读工具";
        this.Size = new Size(1000, 700);
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = true;
        this.MinimumSize = new Size(600, 700);
        this.StartPosition = FormStartPosition.CenterScreen;

        // 热键设置区域（简化为只选择功能键）
        var hotkeyGroup = new GroupBox
        {
            Text = "快捷键设置",
            Location = new Point(15, 15),
            Size = new Size(955, 55)
        };

        var keyLabel = new Label { Text = "朗读热键:", Location = new Point(10, 22), AutoSize = true };
        _keyCombo = new ComboBox
        {
            Location = new Point(175, 19),
            Size = new Size(80, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        for (int i = 1; i <= 12; i++) _keyCombo.Items.Add($"F{i}");
        _keyCombo.SelectedIndex = 0;

        var stopKeyLabel = new Label { Text = "停止热键:", Location = new Point(280, 22), AutoSize = true };
        _stopKeyCombo = new ComboBox
        {
            Location = new Point(445, 19),
            Size = new Size(80, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        for (int i = 1; i <= 12; i++) _stopKeyCombo.Items.Add($"F{i}");
        _stopKeyCombo.SelectedIndex = 1; // 默认 F2

        _applyButton = new Button { Text = "应用", Location = new Point(535, 18), Size = new Size(50, 25) };
        _applyButton.Click += ApplyHotkey_Click;

        _topMostButton = new Button { Text = "置顶", Location = new Point(595, 18), Size = new Size(60, 25) };
        _topMostButton.Click += TopMostButton_Click;

        var hotkeyTip = new Label { Text = "（直接按功能键即可触发）", Location = new Point(665, 22), AutoSize = true, ForeColor = Color.Gray };

        hotkeyGroup.Controls.AddRange(new Control[] { keyLabel, _keyCombo, stopKeyLabel, _stopKeyCombo, _applyButton, _topMostButton, hotkeyTip });

        // 发音人设置区域
        var voiceGroup = new GroupBox
        {
            Text = "发音人设置",
            Location = new Point(15, 80),
            Size = new Size(955, 55)
        };

        var voiceLabel = new Label { Text = "发音人:", Location = new Point(10, 22), AutoSize = true };
        _voiceCombo = new ComboBox
        {
            Location = new Point(165, 19),
            Size = new Size(700, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _voiceCombo.SelectedIndexChanged += VoiceCombo_SelectedIndexChanged;

        _refreshVoiceButton = new Button
        {
            Text = "刷新",
            Location = new Point(880, 18),
            Size = new Size(60, 25)
        };
        _refreshVoiceButton.Click += RefreshVoices_Click;

        voiceGroup.Controls.AddRange(new Control[] { voiceLabel, _voiceCombo, _refreshVoiceButton });

        // 语音参数设置
        var paramGroup = new GroupBox
        {
            Text = "语音参数",
            Location = new Point(15, 145),
            Size = new Size(955, 80)
        };

        _rateLabel = new Label { Text = "语速 (-10 到 10):", Location = new Point(10, 25), AutoSize = true };
        _rateTextBox = new TextBox
        {
            Location = new Point(220, 22),
            Size = new Size(100, 25),
            Text = "0"
        };
        _rateTextBox.KeyPress += (s, e) =>
        {
            // 只允许数字、负号和控制键
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '-')
            {
                e.Handled = true;
            }
            // 负号只能在开头
            if (e.KeyChar == '-' && ((TextBox)s!).Text.Length > 0)
            {
                e.Handled = true;
            }
        };
        _rateTextBox.TextChanged += (s, e) =>
        {
            if (int.TryParse(_rateTextBox.Text, out int value))
            {
                if (value >= -10 && value <= 10)
                {
                    if (_voice != null) _voice.Rate = value;
                    if (!_isLoadingConfig) SaveConfig();
                    _rateTextBox.ForeColor = Color.Black;
                }
                else
                {
                    _rateTextBox.ForeColor = Color.Red;
                }
            }
            else if (string.IsNullOrEmpty(_rateTextBox.Text) || _rateTextBox.Text == "-")
            {
                _rateTextBox.ForeColor = Color.Black;
            }
            else
            {
                _rateTextBox.ForeColor = Color.Red;
            }
        };

        var rateHint = new Label { Text = "范围: -10 ~ 10", Location = new Point(330, 25), AutoSize = true, ForeColor = Color.Gray };

        _volumeLabel = new Label { Text = "音量 (0 到 100):", Location = new Point(500, 25), AutoSize = true };
        _volumeTextBox = new TextBox
        {
            Location = new Point(660, 22),
            Size = new Size(100, 25),
            Text = "100"
        };
        _volumeTextBox.KeyPress += (s, e) =>
        {
            // 只允许数字和控制键
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        };
        _volumeTextBox.TextChanged += (s, e) =>
        {
            if (int.TryParse(_volumeTextBox.Text, out int value))
            {
                if (value >= 0 && value <= 100)
                {
                    if (_voice != null) _voice.Volume = value;
                    if (!_isLoadingConfig) SaveConfig();
                    _volumeTextBox.ForeColor = Color.Black;
                }
                else
                {
                    _volumeTextBox.ForeColor = Color.Red;
                }
            }
            else if (string.IsNullOrEmpty(_volumeTextBox.Text))
            {
                _volumeTextBox.ForeColor = Color.Black;
            }
            else
            {
                _volumeTextBox.ForeColor = Color.Red;
            }
        };

        var volumeHint = new Label { Text = "范围: 0 ~ 100", Location = new Point(770, 25), AutoSize = true, ForeColor = Color.Gray };

        paramGroup.Controls.AddRange(new Control[] { _rateLabel, _rateTextBox, rateHint, _volumeLabel, _volumeTextBox, volumeHint });

        // 朗读记录区域
        var historyGroup = new GroupBox
        {
            Text = "朗读记录（最近1000条）",
            Location = new Point(15, 235),
            Size = new Size(955, 165)
        };

        _historyListBox = new ListBox
        {
            Location = new Point(10, 22),
            Size = new Size(935, 115),
            HorizontalScrollbar = true
        };

        _clearHistoryButton = new Button
        {
            Text = "清空记录",
            Location = new Point(870, 145),
            Size = new Size(75, 25)
        };
        _clearHistoryButton.Click += (s, e) => { _readingHistory.Clear(); _historyListBox.Items.Clear(); };

        historyGroup.Controls.AddRange(new Control[] { _historyListBox, _clearHistoryButton });

        // 转义字符设置区域
        var escapeGroup = new GroupBox
        {
            Text = "转义字符设置（文本替换规则）",
            Location = new Point(15, 410),
            Size = new Size(955, 150)
        };

        var escapeHint = new Label
        {
            Text = "语法：\"\u539f\u6587\u672c1\",\"替\u6362\u6587\u672c1\"; \"\u539f\u6587\u672c2\",\"替\u6362\u6587\u672c2\"; ...",
            Location = new Point(10, 20),
            Size = new Size(935, 20),
            ForeColor = Color.Gray
        };

        _escapeRulesTextBox = new TextBox
        {
            Location = new Point(10, 45),
            Size = new Size(935, 75),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 9F)
        };
        _escapeRulesTextBox.TextChanged += EscapeRulesTextBox_TextChanged;

        var escapeExample = new Label
        {
            Text = "示例: \"\\n\",\" \"; \"(\",\"\"; \")\",\"\"; \"\u3010\",\"\"; \"\u3011\",\"\"",
            Location = new Point(10, 125),
            Size = new Size(935, 20),
            ForeColor = Color.DarkGreen
        };

        _saveConfigButton = new Button
        {
            Text = "保存设置",
            Location = new Point(850, 120),
            Size = new Size(95, 25)
        };
        _saveConfigButton.Click += (s, e) =>
        {
            SaveConfig();
            _statusLabel.Text = "设置已保存";
            _statusLabel.ForeColor = Color.Green;
        };

        escapeGroup.Controls.AddRange(new Control[] { escapeHint, _escapeRulesTextBox, escapeExample, _saveConfigButton });

        // 状态栏
        _statusLabel = new Label
        {
            Text = "就绪 - 按 F1 朗读选中文本，按 F2 停止",
            Location = new Point(15, 570),
            Size = new Size(955, 20),
            ForeColor = Color.Green
        };

        // 说明标签
        var infoLabel = new Label
        {
            Text = "使用说明：选中任意文本，按快捷键即可朗读。",
            Location = new Point(15, 595),
            Size = new Size(955, 40),
            ForeColor = Color.Gray
        };

        this.Controls.AddRange(new Control[] { hotkeyGroup, voiceGroup, paramGroup, historyGroup, escapeGroup, _statusLabel, infoLabel });

        this.Resize += MainForm_Resize;
    }

    private void InitializeVoice()
    {
        _voice = new SpeechSynthesizer();
        RefreshVoiceList();
    }

    private void RefreshVoiceList()
    {
        string? currentVoiceName = null;
        if (_voiceCombo.SelectedIndex >= 0 && _voiceCombo.SelectedIndex < _installedVoices.Count)
        {
            currentVoiceName = _installedVoices[_voiceCombo.SelectedIndex].VoiceInfo.Name;
        }
        
        _voiceCombo.Items.Clear();

        // 重新创建 SpeechSynthesizer 以获取最新的发音人列表
        _voice?.Dispose();
        _voice = new SpeechSynthesizer();

        // 加载所有可用的发音人（包括第三方 SAPI5 发音人）
        _installedVoices = _voice.GetInstalledVoices().Where(v => v.Enabled).ToList();
        foreach (var voice in _installedVoices)
        {
            string displayName = $"{voice.VoiceInfo.Name} ({voice.VoiceInfo.Culture.DisplayName})";
            _voiceCombo.Items.Add(displayName);
        }

        // 尝试恢复之前选择的发音人
        if (_voiceCombo.Items.Count > 0)
        {
            int idx = -1;
            if (currentVoiceName != null)
            {
                for (int i = 0; i < _installedVoices.Count; i++)
                {
                    if (_installedVoices[i].VoiceInfo.Name == currentVoiceName)
                    {
                        idx = i;
                        break;
                    }
                }
            }
            _voiceCombo.SelectedIndex = idx >= 0 ? idx : 0;
        }

        _statusLabel.Text = $"已加载 {_installedVoices.Count} 个发音人";
        _statusLabel.ForeColor = Color.Green;
    }

    private void RefreshVoices_Click(object? sender, EventArgs e)
    {
        RefreshVoiceList();
    }

    private void VoiceCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_voiceCombo.SelectedIndex >= 0 && _voiceCombo.SelectedIndex < _installedVoices.Count)
        {
            try
            {
                var selectedVoice = _installedVoices[_voiceCombo.SelectedIndex];
                _voice.SelectVoice(selectedVoice.VoiceInfo.Name);
                if (!_isLoadingConfig) SaveConfig();
            }
            catch
            {
                // 如果选择失败，使用默认语音
            }
        }
    }

    private void InitializeTrayIcon()
    {
        _trayIcon = new NotifyIcon
        {
            Icon = SystemIcons.Application,
            Text = "SAPI5 朗读工具",
            Visible = true
        };

        var menu = new ContextMenuStrip();
        menu.Items.Add("显示主窗口", null, (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; });
        menu.Items.Add("停止朗读", null, (s, e) => StopSpeaking());
        menu.Items.Add("-");
        menu.Items.Add("退出", null, (s, e) => { _trayIcon.Visible = false; Application.Exit(); });

        _trayIcon.ContextMenuStrip = menu;
        _trayIcon.DoubleClick += (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; };
    }

    private void MainForm_Resize(object? sender, EventArgs e)
    {
        if (this.WindowState == FormWindowState.Minimized)
        {
            this.Hide();
            _trayIcon.ShowBalloonTip(1000, "SAPI5 朗读工具", "程序已最小化到托盘", ToolTipIcon.Info);
        }
    }

    private void RegisterCurrentHotkey()
    {
        // 无修饰键，直接注册功能键
        bool ok = RegisterHotKey(this.Handle, HOTKEY_ID, 0, _hotkeyKey);
        if (!ok)
        {
            _statusLabel.Text = "朗读热键注册失败！可能被其他程序占用";
            _statusLabel.ForeColor = Color.Red;
        }
        
        // 注册停止热键
        bool stopOk = RegisterHotKey(this.Handle, STOP_HOTKEY_ID, 0, _stopHotkeyKey);
        if (!stopOk)
        {
            _statusLabel.Text = "停止热键注册失败！可能被其他程序占用";
            _statusLabel.ForeColor = Color.Red;
        }
        
        if (ok && stopOk)
        {
            UpdateStatusLabel();
        }
    }

    private void UpdateStatusLabel()
    {
        string keyStr = _keyCombo.SelectedItem?.ToString() ?? "F1";
        string stopKeyStr = _stopKeyCombo.SelectedItem?.ToString() ?? "F2";
        _statusLabel.Text = $"就绪 - 按 {keyStr} 朗读选中文本，按 {stopKeyStr} 停止";
        _statusLabel.ForeColor = Color.Green;
    }

    private void ApplyHotkey_Click(object? sender, EventArgs e)
    {
        // 先取消当前热键
        UnregisterHotKey(this.Handle, HOTKEY_ID);
        UnregisterHotKey(this.Handle, STOP_HOTKEY_ID);

        // 解析功能键 (F1=0x70, F2=0x71, ..., F12=0x7B)
        _hotkeyKey = (uint)(0x70 + _keyCombo.SelectedIndex);
        _stopHotkeyKey = (uint)(0x70 + _stopKeyCombo.SelectedIndex);

        // 注册新热键
        RegisterCurrentHotkey();
        SaveConfig();
    }

    private void TopMostButton_Click(object? sender, EventArgs e)
    {
        this.TopMost = !this.TopMost;
        if (this.TopMost)
        {
            _topMostButton.Text = "取消置顶";
            _topMostButton.BackColor = Color.LightBlue;
            _statusLabel.Text = "窗口已置顶";
            _statusLabel.ForeColor = Color.Green;
        }
        else
        {
            _topMostButton.Text = "置顶";
            _topMostButton.BackColor = SystemColors.Control;
            _statusLabel.Text = "已取消置顶";
            _statusLabel.ForeColor = Color.Green;
        }
        SaveConfig();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY)
        {
            if (m.WParam.ToInt32() == HOTKEY_ID)
            {
                HandleHotkey();
            }
            else if (m.WParam.ToInt32() == STOP_HOTKEY_ID)
            {
                StopSpeaking();
            }
        }
        base.WndProc(ref m);
    }

    private async void HandleHotkey()
    {
        // 如果正在朗读，立即停止并重新朗读
        if (_isSpeaking)
        {
            _voice.SpeakAsyncCancelAll();
            _isSpeaking = false;
            // 等待语音引擎真正停止
            await Task.Delay(100);
        }

        try
        {
            // 保存当前剪贴板内容
            _previousClipboard = null;
            if (Clipboard.ContainsText())
            {
                _previousClipboard = Clipboard.GetText();
            }

            // 模拟 Ctrl+C
            SimulateCtrlC();

            // 等待复制完成
            await Task.Delay(150);

            // 读取剪贴板
            string text = "";
            if (Clipboard.ContainsText())
            {
                text = Clipboard.GetText();
            }

            // 如果复制的内容与之前相同，说明没有新选中的文本
            if (string.IsNullOrWhiteSpace(text) || text == _previousClipboard)
            {
                _statusLabel.Text = "未检测到选中的文本";
                _statusLabel.ForeColor = Color.Orange;
                await Task.Delay(2000);
                UpdateStatusLabel();
                return;
            }

            // 应用转义规则
            string processedText = ApplyEscapeRules(text);

            // 添加到朗读记录（保存转义后的文本到内存）
            AddReadingRecord(processedText);

            // 立即清理剪贴板（恢复之前的内容或清空）
            if (_previousClipboard != null)
            {
                Clipboard.SetText(_previousClipboard);
            }
            else
            {
                Clipboard.Clear();
            }

            // 开始朗读（从内存中读取）
            _isSpeaking = true;
            _statusLabel.Text = "正在朗读...";
            _statusLabel.ForeColor = Color.Blue;

            // 异步朗读
            _voice.SpeakAsync(processedText);
            
            // 等待朗读完成
            while (_isSpeaking && _voice.State == SynthesizerState.Speaking)
            {
                await Task.Delay(50); // 减少轮询间隔，提高响应速度
            }

            // 只有在正常完成时才重置状态
            if (_isSpeaking)
            {
                _isSpeaking = false;
                UpdateStatusLabel();
            }
        }
        catch (Exception ex)
        {
            _isSpeaking = false;
            // 忽略取消异常
            if (!ex.Message.Contains("canceled") && !ex.Message.Contains("cancelled"))
            {
                _statusLabel.Text = $"错误: {ex.Message}";
                _statusLabel.ForeColor = Color.Red;
            }
            else
            {
                UpdateStatusLabel();
            }
        }
    }

    private void StopSpeaking()
    {
        // 强制中断语音输出，不考虑后续流程
        _voice.SpeakAsyncCancelAll();
        _isSpeaking = false;
        
        // 确保语音引擎已完全停止
        while (_voice.State == SynthesizerState.Speaking)
        {
            System.Threading.Thread.Sleep(10);
        }
        
        UpdateStatusLabel();
    }

    private void SimulateCtrlC()
    {
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYDOWN, 0);
        keybd_event(VK_C, 0, KEYEVENTF_KEYDOWN, 0);
        keybd_event(VK_C, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        SaveConfig();
        UnregisterHotKey(this.Handle, HOTKEY_ID);
        UnregisterHotKey(this.Handle, STOP_HOTKEY_ID);
        _trayIcon.Visible = false;
        base.OnFormClosed(e);
    }

    private void AddReadingRecord(string text)
    {
        var record = new ReadingRecord
        {
            Time = DateTime.Now,
            Text = text.Length > 100 ? text.Substring(0, 100) + "..." : text
        };

        // 超过1000条则移除最旧的
        if (_readingHistory.Count >= MAX_HISTORY)
        {
            _readingHistory.RemoveAt(0);
            _historyListBox.Items.RemoveAt(0);
        }

        _readingHistory.Add(record);
        _historyListBox.Items.Add($"[{record.Time:HH:mm:ss}] {record.Text}");
        _historyListBox.TopIndex = _historyListBox.Items.Count - 1; // 滚动到最新
    }

    private void EscapeRulesTextBox_TextChanged(object? sender, EventArgs e)
    {
        ParseEscapeRules();
        if (!_isLoadingConfig) SaveConfig();
    }

    private void ParseEscapeRules()
    {
        _escapeRules.Clear();
        try
        {
            string input = _escapeRulesTextBox.Text;
            if (string.IsNullOrWhiteSpace(input)) return;

            // 解析规则："\u539f\u6587\u672c","\u66ff\u6362\u6587\u672c";
            var rules = input.Split(';');
            foreach (var rule in rules)
            {
                var trimmedRule = rule.Trim();
                if (string.IsNullOrWhiteSpace(trimmedRule)) continue;

                // 匹配 "xxx","yyy" 格式
                var match = System.Text.RegularExpressions.Regex.Match(trimmedRule, @"""([^""]*)"",""([^""]*)""");
                if (match.Success)
                {
                    string original = match.Groups[1].Value;
                    string replacement = match.Groups[2].Value;
                    
                    // 处理转义字符
                    original = original.Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t");
                    replacement = replacement.Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t");
                    
                    if (!_escapeRules.ContainsKey(original))
                    {
                        _escapeRules.Add(original, replacement);
                    }
                }
            }
        }
        catch
        {
            // 解析失败，忽略
        }
    }

    private string ApplyEscapeRules(string text)
    {
        if (_escapeRules.Count == 0) return text;

        string result = text;
        foreach (var rule in _escapeRules)
        {
            result = result.Replace(rule.Key, rule.Value);
        }
        return result;
    }

    private void SaveConfig()
    {
        try
        {
            // 确保配置目录存在
            var configDir = Path.GetDirectoryName(_configPath);
            if (!string.IsNullOrEmpty(configDir) && !Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }

            // 获取当前选择的发音人的实际名称
            string voiceName = "";
            if (_voiceCombo.SelectedIndex >= 0 && _voiceCombo.SelectedIndex < _installedVoices.Count)
            {
                voiceName = _installedVoices[_voiceCombo.SelectedIndex].VoiceInfo.Name;
            }

            var config = new AppConfig
            {
                HotkeyIndex = _keyCombo.SelectedIndex,
                StopHotkeyIndex = _stopKeyCombo.SelectedIndex,
                VoiceName = voiceName,
                Rate = int.TryParse(_rateTextBox.Text, out int rate) ? rate : 0,
                Volume = int.TryParse(_volumeTextBox.Text, out int volume) ? volume : 100,
                EscapeRules = _escapeRulesTextBox.Text,
                TopMost = this.TopMost
            };

            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_configPath, json);
        }
        catch
        {
            // 保存失败，忽略
        }
    }

    private void LoadConfig()
    {
        try
        {
            if (!File.Exists(_configPath)) return;

            var json = File.ReadAllText(_configPath);
            var config = JsonSerializer.Deserialize<AppConfig>(json);
            if (config == null) return;

            // 恢复热键设置
            if (config.HotkeyIndex >= 0 && config.HotkeyIndex < _keyCombo.Items.Count)
            {
                _keyCombo.SelectedIndex = config.HotkeyIndex;
                _hotkeyKey = (uint)(0x70 + config.HotkeyIndex);
            }
            
            if (config.StopHotkeyIndex >= 0 && config.StopHotkeyIndex < _stopKeyCombo.Items.Count)
            {
                _stopKeyCombo.SelectedIndex = config.StopHotkeyIndex;
                _stopHotkeyKey = (uint)(0x70 + config.StopHotkeyIndex);
            }

            // 恢复语音参数
            if (config.Rate >= -10 && config.Rate <= 10)
            {
                _rateTextBox.Text = config.Rate.ToString();
            }

            if (config.Volume >= 0 && config.Volume <= 100)
            {
                _volumeTextBox.Text = config.Volume.ToString();
            }

            // 恢复转义规则
            if (!string.IsNullOrEmpty(config.EscapeRules))
            {
                _escapeRulesTextBox.Text = config.EscapeRules;
            }

            // 恢复发音人（精确匹配 VoiceInfo.Name）
            if (!string.IsNullOrEmpty(config.VoiceName))
            {
                for (int i = 0; i < _installedVoices.Count; i++)
                {
                    if (_installedVoices[i].VoiceInfo.Name == config.VoiceName)
                    {
                        _voiceCombo.SelectedIndex = i;
                        break;
                    }
                }
            }

            // 恢复置顶状态
            this.TopMost = config.TopMost;
            if (this.TopMost)
            {
                _topMostButton.Text = "取消置顶";
                _topMostButton.BackColor = Color.LightBlue;
            }
            else
            {
                _topMostButton.Text = "置顶";
                _topMostButton.BackColor = SystemColors.Control;
            }
        }
        catch
        {
            // 加载失败，使用默认设置
        }
    }
}

// 配置类
public class AppConfig
{
    public int HotkeyIndex { get; set; }
    public int StopHotkeyIndex { get; set; } = 1; // 默认 F2
    public string VoiceName { get; set; } = "";
    public int Rate { get; set; }
    public int Volume { get; set; } = 100;
    public string EscapeRules { get; set; } = "";
    public bool TopMost { get; set; } = false;
}

// 朗读记录类
public class ReadingRecord
{
    public DateTime Time { get; set; }
    public string Text { get; set; } = "";
}
