using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Text.Json;
using System.Diagnostics;
using System.Media;
using System.Text.Json.Serialization;

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
    private const int AUTO_READ_HOTKEY_ID = 0x5678;
    private const int WM_HOTKEY = 0x0312;

    private const int KEYEVENTF_KEYDOWN = 0x0000;
    private const int KEYEVENTF_KEYUP = 0x0002;
    private const byte VK_CONTROL = 0x11;
    private const byte VK_C = 0x43;
    private const byte VK_SHIFT = 0x10;
    private const byte VK_HOME = 0x24;
    private const byte VK_END = 0x23;
    private const byte VK_DOWN = 0x28;

    #endregion

    private SpeechSynthesizer _voice = null!;
    private NotifyIcon _trayIcon = null!;
    private bool _isSpeaking = false;
    private bool _isAutoReading = false;
    private IDataObject? _previousClipboard = null;
    private List<InstalledVoice> _installedVoices = new();
    private string _chineseVoiceName = "";
    private string _englishVoiceName = "";
    private bool _autoSwitchVoice = true;
    private int _chineseVoiceRate = 0;
    private int _englishVoiceRate = 0;
    private int _mixedChineseMinChars = 2;
    private int _mixedEnglishMinLetters = 3;
    private bool _warmUpVoices = true;
    private bool _preventSpeakerSleep = false;
    private SoundPlayer? _speakerKeepAlivePlayer;
    private MemoryStream? _speakerKeepAliveStream;

    // 朗读记录（最多1000条，超出覆盖）
    private int _maxHistory = 1000;
    private List<ReadingRecord> _readingHistory = new();

    // 当前热键配置（无修饰键）
    private uint _hotkeyKey = 0x70; // F1
    private uint _stopHotkeyKey = 0x71; // F2
    private uint _autoReadHotkeyKey = 0x70; // F1 (Ctrl+Shift+F1)
    private uint _autoReadModifiers = MOD_CONTROL | MOD_SHIFT; // 默认修饰键
    private string _customAutoReadHotKey = ""; // 当前生效的热键字符串
    private const int STOP_HOTKEY_ID = 2;

    // UI 控件 (新版)
    private TextBox _configEditor = null!;
    private Button _saveConfigButton = null!;
    private Button _listVoicesButton = null!;
    private ListBox _historyListBox = null!;
    private Button _clearHistoryButton = null!;
    private Label _statusLabel = null!;

    // 配置相关字段
    private bool _onlyReadFirstLine = false;
    private bool _clearClipboard = true;
    private bool _removeSpaces = false;
    private int _autoReadDelay = 1000;
    
    // 转义字符规则
    private Dictionary<string, string> _escapeRules = new();
    private List<KeyValuePair<string, string>> _escapeRulesOrdered = new();

    // 配置文件路径 - 使用项目根目录
    private string _configPath = "";
    private FileSystemWatcher? _configWatcher;
    private bool _restartScheduled = false;
    private bool _selfSaving = false;

    public MainForm()
    {
        Log("MainForm Constructor Started");
        // 计算配置文件路径
        // 优先查找当前工作目录（方便开发调试和便携使用）
        string currentDirConfig = Path.Combine(Directory.GetCurrentDirectory(), "config.json");
        if (File.Exists(currentDirConfig))
        {
            _configPath = currentDirConfig;
        }
        else
        {
            // 否则使用应用程序所在目录
            _configPath = Path.Combine(AppContext.BaseDirectory, "config.json");
        }
        
        InitializeComponent();
        Log("InitializeComponent Done");
        InitializeVoice();
        Log("InitializeVoice Done");
        InitializeTrayIcon();
        Log("InitializeTrayIcon Done");
        LoadConfig();
        Log("LoadConfig Done");
        InitializeConfigWatcher();
    }

    private void Log(string message)
    {
        try { File.AppendAllText("debug.log", $"[{DateTime.Now}] {message}\n"); } catch {}
    }

    protected override void OnLoad(EventArgs e)
    {
        Log("OnLoad Started");
        base.OnLoad(e);
        Log("OnLoad Done");
    }

    protected override void OnShown(EventArgs e)
    {
        Log("OnShown Started");
        base.OnShown(e);
        Log("OnShown Done");
    }

    private void InitializeComponent()
    {
        this.Text = "SAPI5 文本朗读工具 (Config Editor Mode)";
        this.Size = new Size(1000, 700);
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = true;
        this.MinimumSize = new Size(600, 700);
        this.StartPosition = FormStartPosition.CenterScreen;

        // 1. 配置编辑器区域
        var editorLabel = new Label
        {
            Text = "配置文件编辑器 (config.json):",
            Location = new Point(15, 10),
            AutoSize = true,
            Font = new Font(this.Font, FontStyle.Bold)
        };

        _configEditor = new TextBox
        {
            Location = new Point(15, 35),
            Size = new Size(955, 300),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 10F),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        var editorMenu = new ContextMenuStrip();
        var openDefault = new ToolStripMenuItem("用默认程序打开 config.json");
        openDefault.Click += (s, e) => OpenConfigWithDefault();
        var openWith = new ToolStripMenuItem("选择其他程序打开 config.json");
        openWith.Click += (s, e) => OpenConfigWithDialog();
        var openFolder = new ToolStripMenuItem("在资源管理器中定位 config.json");
        openFolder.Click += (s, e) => OpenConfigInFolder();
        editorMenu.Items.AddRange(new ToolStripItem[] { openDefault, openWith, openFolder });
        _configEditor.ContextMenuStrip = editorMenu;

        // 2. 按钮区域
        _saveConfigButton = new Button
        {
            Text = "保存并应用配置",
            Location = new Point(15, 345),
            Size = new Size(150, 30),
            BackColor = Color.LightGreen
        };
        _saveConfigButton.Click += (s, e) => SaveConfig();

        _listVoicesButton = new Button
        {
            Text = "发音人列表(点击复制)▼",
            Location = new Point(180, 345),
            Size = new Size(180, 30)
        };
        _listVoicesButton.Click += ListVoices_Click;

        var helpLabel = new Label
        {
            Text = "提示：直接修改上方 JSON 内容，点击保存即可生效。所有设置均通过此文件管理。",
            Location = new Point(340, 350),
            AutoSize = true,
            ForeColor = Color.Gray
        };

        // 3. 朗读记录区域
        var historyLabel = new Label
        {
            Text = "朗读记录:",
            Location = new Point(15, 390),
            AutoSize = true
        };

        _historyListBox = new ListBox
        {
            Location = new Point(15, 415),
            Size = new Size(955, 180),
            HorizontalScrollbar = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        _clearHistoryButton = new Button
        {
            Text = "清空记录",
            Location = new Point(895, 385),
            Size = new Size(75, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _clearHistoryButton.Click += (s, e) => { _readingHistory.Clear(); _historyListBox.Items.Clear(); };

        // 4. 状态栏
        _statusLabel = new Label
        {
            Text = "就绪",
            Dock = DockStyle.Bottom,
            AutoSize = false,
            Height = 25,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(5, 0, 0, 0),
            BackColor = SystemColors.ControlLight,
            ForeColor = Color.Green
        };

        this.Controls.AddRange(new Control[] { 
            editorLabel, _configEditor, 
            _saveConfigButton, _listVoicesButton, helpLabel,
            historyLabel, _historyListBox, _clearHistoryButton,
            _statusLabel 
        });

        this.Resize += MainForm_Resize;
    }

    private void InitializeVoice()
    {
        _voice = new SpeechSynthesizer();
        // 获取发音人列表供查询
        _installedVoices = _voice.GetInstalledVoices().Where(v => v.Enabled).ToList();
    }

    private void ListVoices_Click(object? sender, EventArgs e)
    {
        var menu = new ContextMenuStrip();
        
        if (_installedVoices.Count == 0)
        {
             menu.Items.Add("未检测到发音人").Enabled = false;
        }
        else
        {
             // 添加提示项
             var titleItem = menu.Items.Add("点击以下名称即可复制:");
             titleItem.Enabled = false;
             titleItem.BackColor = Color.WhiteSmoke;
             menu.Items.Add(new ToolStripSeparator());

             foreach (var voice in _installedVoices)
             {
                 string name = voice.VoiceInfo.Name;
                 menu.Items.Add(name, null, (s, args) => 
                 {
                     try 
                     {
                         Clipboard.SetText(name);
                         _statusLabel.Text = $"已复制发音人: \"{name}\"";
                         _statusLabel.ForeColor = Color.Green;
                         
                        if (_configEditor.Text.Contains("\"ChineseVoiceName\": \"\"")
                            || _configEditor.Text.Contains("\"EnglishVoiceName\": \"\"")
                            || _configEditor.Text.Contains("\"VoiceName\": \"\""))
                        {
                            MessageBox.Show($"已复制: {name}\r\n可粘贴到 ChineseVoiceName / EnglishVoiceName / VoiceName 字段的双引号中 (Ctrl+V)", "复制成功");
                        }
                     }
                     catch (Exception ex)
                     {
                         MessageBox.Show($"复制失败: {ex.Message}", "错误");
                     }
                 });
             }
        }
        
        menu.Show(_listVoicesButton, 0, _listVoicesButton.Height);
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
        menu.Items.Add("开始自动朗读", null, (s, e) => { 
            Task.Run(async () => {
                await Task.Delay(3000); // 3秒倒计时
                this.Invoke(() => StartAutoRead());
            });
            _trayIcon.ShowBalloonTip(3000, "准备开始", "请在3秒内切换到要朗读的文本窗口...", ToolTipIcon.Info);
        });
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
        // 先取消当前热键
        UnregisterHotKey(this.Handle, HOTKEY_ID);
        UnregisterHotKey(this.Handle, STOP_HOTKEY_ID);
        UnregisterHotKey(this.Handle, AUTO_READ_HOTKEY_ID);

        // 注册朗读热键
        bool ok = RegisterHotKey(this.Handle, HOTKEY_ID, 0, _hotkeyKey);
        if (!ok)
        {
            _statusLabel.Text = $"朗读热键(F{_hotkeyKey - 0x6F})注册失败！可能被占用";
            _statusLabel.ForeColor = Color.Red;
        }
        
        // 注册停止热键
        bool stopOk = RegisterHotKey(this.Handle, STOP_HOTKEY_ID, 0, _stopHotkeyKey);
        if (!stopOk)
        {
            _statusLabel.Text = $"停止热键(F{_stopHotkeyKey - 0x6F})注册失败！可能被占用";
            _statusLabel.ForeColor = Color.Red;
        }
        
        // 注册自动朗读热键
        bool autoReadOk = RegisterHotKey(this.Handle, AUTO_READ_HOTKEY_ID, _autoReadModifiers, _autoReadHotkeyKey);
        
        if (ok && stopOk)
        {
            UpdateStatusLabel();
        }
    }

    private void UpdateStatusLabel()
    {
        string keyStr = $"F{_hotkeyKey - 0x6F}";
        string stopKeyStr = $"F{_stopHotkeyKey - 0x6F}";
        _statusLabel.Text = $"配置已应用 - 按 {keyStr} 朗读，按 {stopKeyStr} 停止";
        _statusLabel.ForeColor = Color.Green;
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
            else if (m.WParam.ToInt32() == AUTO_READ_HOTKEY_ID)
            {
                if (_isAutoReading)
                {
                    StopSpeaking();
                }
                else
                {
                    StartAutoRead();
                }
            }
        }
        base.WndProc(ref m);
    }

    private async void HandleHotkey()
    {
        // 1. 立即中断之前的朗读 (提高响应速度)
        if (_isSpeaking || _voice.State == SynthesizerState.Speaking)
        {
            _voice.SpeakAsyncCancelAll();
            _isSpeaking = false;
            // 不在这里等待，利用获取剪贴板的时间作为缓冲
        }

        try
        {
            _previousClipboard = null;
            try
            {
                var data = Clipboard.GetDataObject();
                if (data != null) _previousClipboard = data;
            }
            catch {}

            // 优化: 先清空剪贴板，以便检测新内容是否已复制
            try { Clipboard.Clear(); } catch {}

            SimulateCtrlC();
            
            // 优化: 轮询检查剪贴板，代替固定延迟 (提高响应速度)
            string text = "";
            for (int i = 0; i < 20; i++) // 最多等待 400ms
            {
                await Task.Delay(20);
                if (Clipboard.ContainsText())
                {
                    string t = Clipboard.GetText();
                    if (!string.IsNullOrWhiteSpace(t))
                    {
                        text = t;
                        break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                _statusLabel.Text = "剪贴板为空";
                _statusLabel.ForeColor = Color.Orange;
                await Task.Delay(2000);
                UpdateStatusLabel();
                return;
            }
            
            if (_onlyReadFirstLine)
            {
                int newlineIndex = text.IndexOfAny(new[] { '\r', '\n' });
                if (newlineIndex >= 0) text = text.Substring(0, newlineIndex);
            }

            if (_removeSpaces)
            {
                text = text.Replace(" ", "").Replace("\t", "").Replace("\u3000", "");
            }

            string processedText = ApplyEscapeRules(text);

            AddReadingRecord(processedText);

            if (_clearClipboard)
            {
                try
                {
                    if (_previousClipboard != null) Clipboard.SetDataObject(_previousClipboard, true);
                    else Clipboard.Clear();
                }
                catch {}
            }

            // 关键修复: 确保之前的朗读已完全停止 (解决长文本朗读时无法切换新内容的问题)
            int safetyWait = 0;
            while (_voice.State == SynthesizerState.Speaking && safetyWait < 50)
            {
                await Task.Delay(10);
                safetyWait++;
            }

            _isSpeaking = true;
            _statusLabel.Text = "正在朗读...";
            _statusLabel.ForeColor = Color.Blue;

            SpeakText(processedText);
            
            while (_isSpeaking && _voice.State == SynthesizerState.Speaking)
            {
                await Task.Delay(50);
            }

            if (_isSpeaking)
            {
                _isSpeaking = false;
                UpdateStatusLabel();
            }
        }
        catch (Exception ex)
        {
            _isSpeaking = false;
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
        _isAutoReading = false;
        _voice.SpeakAsyncCancelAll();
        _isSpeaking = false;
        while (_voice.State == SynthesizerState.Speaking)
        {
            System.Threading.Thread.Sleep(10);
        }
        UpdateStatusLabel();
    }

    private async void StartAutoRead()
    {
        if (_isSpeaking)
        {
            StopSpeaking();
            await Task.Delay(100);
        }

        _isAutoReading = true;
        _statusLabel.Text = "自动逐行朗读模式...";
        _statusLabel.ForeColor = Color.Blue;

        int emptyLineCount = 0;

        try
        {
            while (_isAutoReading)
            {
                bool readSuccess = await ReadCurrentSelectionAsync();

                if (!readSuccess)
                {
                    emptyLineCount++;
                    if (emptyLineCount >= 2)
                    {
                        _statusLabel.Text = "连续两次无效内容，自动朗读停止";
                        _statusLabel.ForeColor = Color.Orange;
                        _isAutoReading = false;
                        break;
                    }
                }
                else
                {
                    emptyLineCount = 0;
                    
                    // 等待朗读开始 (最多等待 1 秒)
                    int waitStart = 0;
                    while (_isAutoReading && _isSpeaking && _voice.State != SynthesizerState.Speaking && waitStart < 20)
                    {
                        await Task.Delay(50);
                        waitStart++;
                    }

                    // 等待朗读结束
                    while (_isAutoReading && _isSpeaking && _voice.State == SynthesizerState.Speaking)
                    {
                        await Task.Delay(50);
                    }
                    
                    if (_isSpeaking) _isSpeaking = false;
                    if (!_isAutoReading) break;
                }

                int delay = _autoReadDelay;
                if (delay < 0) delay = 1000;
                await Task.Delay(delay);

                if (!_isAutoReading) break;

                keybd_event(VK_DOWN, 0, KEYEVENTF_KEYDOWN, 0);
                keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
                await Task.Delay(50);
            }
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"自动朗读错误: {ex.Message}";
            _statusLabel.ForeColor = Color.Red;
            _isAutoReading = false;
        }
        finally
        {
            if (!_isAutoReading)
            {
                UpdateStatusLabel();
            }
        }
    }

    private async Task<bool> ReadCurrentSelectionAsync()
    {
        try
        {
            _previousClipboard = null;
            try
            {
                var data = Clipboard.GetDataObject();
                if (data != null) _previousClipboard = data;
            }
            catch {}

            // 2. 复制
            SimulateCtrlC();
            await Task.Delay(150);

            string text = "";
            if (Clipboard.ContainsText())
            {
                text = Clipboard.GetText();
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                return false; 
            }

            if (_onlyReadFirstLine)
            {
                int newlineIndex = text.IndexOfAny(new[] { '\r', '\n' });
                if (newlineIndex >= 0) text = text.Substring(0, newlineIndex);
            }

            if (_removeSpaces)
            {
                text = text.Replace(" ", "").Replace("\t", "").Replace("\u3000", "");
            }

            string processedText = ApplyEscapeRules(text);

            AddReadingRecord(processedText);
            
            _isSpeaking = true;
            _statusLabel.Text = "正在朗读...";
            _statusLabel.ForeColor = Color.Blue;

            SpeakText(processedText);
            
            return true;
        }
        catch (Exception)
        {
            return false;
        }
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
        UnregisterHotKey(this.Handle, HOTKEY_ID);
        UnregisterHotKey(this.Handle, STOP_HOTKEY_ID);
        UnregisterHotKey(this.Handle, AUTO_READ_HOTKEY_ID);
        _trayIcon.Visible = false;
        StopSpeakerKeepAlive();
        base.OnFormClosed(e);
    }

    private void AddReadingRecord(string text)
    {
        var record = new ReadingRecord
        {
            Time = DateTime.Now,
            Text = text.Length > 100 ? text.Substring(0, 100) + "..." : text
        };

        if (_readingHistory.Count >= _maxHistory)
        {
            _readingHistory.RemoveAt(0);
            _historyListBox.Items.RemoveAt(0);
        }

        _readingHistory.Add(record);
        _historyListBox.Items.Add($"[{record.Time:HH:mm:ss}] {record.Text}");
        _historyListBox.TopIndex = _historyListBox.Items.Count - 1;
    }

    private void LoadEscapeRules(object input)
    {
        _escapeRules.Clear();
        _escapeRulesOrdered.Clear();
        try
        {
            if (input is JsonElement element)
            {
                if (element.ValueKind == JsonValueKind.Object)
                {
                    foreach (var property in element.EnumerateObject())
                    {
                        string key = property.Name;
                        string value = property.Value.GetString() ?? "";
                        if (!_escapeRules.ContainsKey(key))
                        {
                            _escapeRules[key] = value;
                        }
                    }
                    _escapeRulesOrdered = _escapeRules
                        .OrderByDescending(kv => kv.Key.Length)
                        .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                        .ToList();
                    return;
                }
                else if (element.ValueKind == JsonValueKind.String)
                {
                    input = element.GetString() ?? "";
                }
            }

            if (input is string strInput)
            {
                if (string.IsNullOrWhiteSpace(strInput)) return;
                var rules = strInput.Split(';');
                foreach (var rule in rules)
                {
                    var trimmedRule = rule.Trim();
                    if (string.IsNullOrWhiteSpace(trimmedRule)) continue;
                    var match = System.Text.RegularExpressions.Regex.Match(trimmedRule, @"""([^""]*)"",""([^""]*)""");
                    if (match.Success)
                    {
                        string original = match.Groups[1].Value;
                        string replacement = match.Groups[2].Value;
                        original = original.Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t");
                        replacement = replacement.Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t");
                        if (!_escapeRules.ContainsKey(original))
                        {
                            _escapeRules[original] = replacement;
                        }
                    }
                }
                _escapeRulesOrdered = _escapeRules
                    .OrderByDescending(kv => kv.Key.Length)
                    .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                    .ToList();
            }
        }
        catch { }
    }

    private string ApplyEscapeRules(string text)
    {
        if (_escapeRulesOrdered.Count == 0 && _escapeRules.Count > 0)
        {
            _escapeRulesOrdered = _escapeRules
                .OrderByDescending(kv => kv.Key.Length)
                .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                .ToList();
        }
        if (_escapeRulesOrdered.Count == 0) return text;
        string result = text;
        foreach (var rule in _escapeRulesOrdered)
        {
            result = result.Replace(rule.Key, rule.Value);
        }
        return result;
    }

    private string TryFixJson(string json)
    {
        try
        {
            // 尝试修复未加引号的 AutoReadHotkeyIndex 值
            // 匹配 "AutoReadHotkeyIndex": Value (排除引号/数字/布尔/null开头)
            var regex = new System.Text.RegularExpressions.Regex(
                @"""AutoReadHotkeyIndex""\s*:\s*(?!""|\d|true|false|null)(?<val>[^,\r\n}]+)", 
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            
            return regex.Replace(json, match => 
            {
                string val = match.Groups["val"].Value.Trim();
                return $"\"AutoReadHotkeyIndex\": \"{val}\"";
            });
        }
        catch
        {
            return json;
        }
    }

    private void SaveConfig()
    {
        try
        {
            var configDir = Path.GetDirectoryName(_configPath);
            if (!string.IsNullOrEmpty(configDir) && !Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }

            string jsonContent = _configEditor.Text;
            
            // 验证 JSON 格式
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };

            AppConfig? config = null;
            try
            {
                config = JsonSerializer.Deserialize<AppConfig>(jsonContent, options);
            }
            catch (JsonException)
            {
                // 尝试自动修复常见错误 (如忘记加引号)
                string fixedJson = TryFixJson(jsonContent);
                if (fixedJson != jsonContent)
                {
                    try
                    {
                        config = JsonSerializer.Deserialize<AppConfig>(fixedJson, options);
                        // 如果修复并解析成功，更新编辑器内容
                        jsonContent = fixedJson;
                        _configEditor.Text = fixedJson;
                    }
                    catch { /* 忽略第二次错误 */ }
                }
                
                if (config == null) throw;
            }
            
            if (config != null)
            {
                // 如果验证通过，写入文件
                _selfSaving = true;
                File.WriteAllText(_configPath, jsonContent);
                _selfSaving = false;
                
                // 应用配置
                ApplyConfig(config);
                
                _statusLabel.Text = "配置已保存并生效";
                _statusLabel.ForeColor = Color.Green;
            }
            else
            {
                MessageBox.Show("JSON 格式错误，请检查！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存失败: {ex.Message}\r\n\r\n提示: 字符串值必须用双引号括起来，例如 \"Ctrl+Alt+S\"", "配置错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LoadConfig()
    {
        try
        {
            if (!File.Exists(_configPath)) return;

            var json = File.ReadAllText(_configPath);
            _configEditor.Text = json; // 显示在编辑器中

            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };
            var config = JsonSerializer.Deserialize<AppConfig>(json, options);
            
            if (config != null)
            {
                ApplyConfig(config);
            }
        }
        catch
        {
            // 加载失败
        }
    }
    
    private void InitializeConfigWatcher()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_configPath)) return;
            var dir = Path.GetDirectoryName(_configPath);
            var name = Path.GetFileName(_configPath);
            if (string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(name)) return;
            _configWatcher?.Dispose();
            _configWatcher = new FileSystemWatcher(dir, name)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName
            };
            _configWatcher.Changed += (_, __) => OnExternalConfigChanged();
            _configWatcher.Created += (_, __) => OnExternalConfigChanged();
            _configWatcher.Renamed += (_, __) => OnExternalConfigChanged();
            _configWatcher.EnableRaisingEvents = true;
        }
        catch { }
    }
    
    private void OnExternalConfigChanged()
    {
        if (_selfSaving) return;
        if (_restartScheduled) return;
        _restartScheduled = true;
        Task.Run(async () =>
        {
            await Task.Delay(800);
            try
            {
                BeginInvoke(new Action(() =>
                {
                    try
                    {
                        Application.Restart();
                    }
                    catch
                    {
                        try
                        {
                            var exe = Application.ExecutablePath;
                            var args = string.Join(" ", Environment.GetCommandLineArgs().Skip(1).Select(QuoteArg));
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = exe,
                                Arguments = args,
                                UseShellExecute = true
                            });
                        }
                        catch { }
                    }
                    try { Environment.Exit(0); } catch { }
                }));
            }
            catch
            {
                try { Environment.Exit(0); } catch { }
            }
        });
    }
    
    private static string QuoteArg(string s)
    {
        if (string.IsNullOrEmpty(s)) return "\"\"";
        if (s.Any(char.IsWhiteSpace) || s.Contains('"')) return $"\"{s.Replace("\"", "\\\"")}\"";
        return s;
    }
    
    private void OpenConfigWithDefault()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _configPath,
                UseShellExecute = true
            });
        }
        catch { }
    }
    
    private void OpenConfigWithDialog()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "rundll32.exe",
                Arguments = $"shell32.dll,OpenAs_RunDLL \"{_configPath}\"",
                UseShellExecute = true
            });
        }
        catch { }
    }
    
    private void OpenConfigInFolder()
    {
        try
        {
            Process.Start("explorer.exe", $"/select,\"{_configPath}\"");
        }
        catch { }
    }
    
    private void ApplyConfig(AppConfig config)
    {
        // 1. 热键设置
        if (config.HotkeyIndex >= 0 && config.HotkeyIndex < 12)
        {
            _hotkeyKey = (uint)(0x70 + config.HotkeyIndex);
        }
        
        if (config.StopHotkeyIndex >= 0 && config.StopHotkeyIndex < 12)
        {
            _stopHotkeyKey = (uint)(0x70 + config.StopHotkeyIndex);
        }

        // AutoReadHotkeyIndex logic
        string? autoReadStr = null;
        int autoReadIdx = -1;

        if (config.AutoReadHotkeyIndex is JsonElement element)
        {
            if (element.ValueKind == JsonValueKind.String) autoReadStr = element.GetString();
            else if (element.ValueKind == JsonValueKind.Number) autoReadIdx = element.GetInt32();
        }
        else if (config.AutoReadHotkeyIndex is string s) autoReadStr = s;
        else if (config.AutoReadHotkeyIndex is int i) autoReadIdx = i;
        else try { autoReadIdx = Convert.ToInt32(config.AutoReadHotkeyIndex); } catch { }

        if (!string.IsNullOrEmpty(autoReadStr))
        {
            try
            {
                uint mods = 0;
                uint key = 0;
                var parts = autoReadStr.ToLower().Split('+');
                foreach (var part in parts)
                {
                    var p = part.Trim();
                    if (p == "ctrl" || p == "control") mods |= MOD_CONTROL;
                    else if (p == "alt") mods |= MOD_ALT;
                    else if (p == "shift") mods |= MOD_SHIFT;
                    else
                    {
                        if (p.Length == 1) key = (uint)p.ToUpper()[0];
                        else if (p.StartsWith("f") && int.TryParse(p.Substring(1), out int fNum))
                        {
                            if (fNum >= 1 && fNum <= 24) key = (uint)(0x70 + fNum - 1);
                        }
                    }
                }

                if (key != 0)
                {
                    _autoReadModifiers = mods;
                    _autoReadHotkeyKey = key;
                    _customAutoReadHotKey = autoReadStr;
                }
            }
            catch { }
        }
        else if (autoReadIdx >= 0 && autoReadIdx < 12)
        {
            _autoReadModifiers = MOD_CONTROL | MOD_SHIFT;
            _autoReadHotkeyKey = (uint)(0x70 + autoReadIdx);
            _customAutoReadHotKey = $"Ctrl+Shift+F{autoReadIdx + 1}";
        }

        RegisterCurrentHotkey();

        // 2. 语音参数
        if (_voice != null)
        {
            _voice.Rate = ClampRate(config.Rate);
            if (config.Volume >= 0 && config.Volume <= 100) _voice.Volume = config.Volume;
            
            if (!string.IsNullOrEmpty(config.VoiceName))
            {
                try
                {
                    _voice.SelectVoice(config.VoiceName);
                }
                catch {}
            }
        }

        // 3. 其他选项
        _onlyReadFirstLine = config.OnlyReadFirstLine;
        _clearClipboard = config.ClearClipboard;
        _removeSpaces = config.RemoveSpaces;
        _autoReadDelay = config.AutoReadDelay;
        
        if (config.RecordNumber > 0)
        {
            _maxHistory = config.RecordNumber;
            // Trim history
            while (_readingHistory.Count > _maxHistory)
            {
                _readingHistory.RemoveAt(0);
                _historyListBox.Items.RemoveAt(0);
            }
        }

        LoadEscapeRules(config.EscapeRules);
        this.TopMost = config.TopMost;

        _autoSwitchVoice = config.AutoSwitchVoice;
        var desiredChinese = string.IsNullOrWhiteSpace(config.ChineseVoiceName) ? config.VoiceName : config.ChineseVoiceName;
        var desiredEnglish = string.IsNullOrWhiteSpace(config.EnglishVoiceName) ? TryGetDefaultEnglishVoiceName() : config.EnglishVoiceName;
        _chineseVoiceName = ResolveVoiceName(desiredChinese);
        _englishVoiceName = ResolveVoiceName(desiredEnglish);
        _chineseVoiceRate = ClampRate(config.ChineseVoiceRate ?? config.Rate);
        _englishVoiceRate = ClampRate(config.EnglishVoiceRate ?? config.Rate);
        _mixedChineseMinChars = config.MixedChineseMinChars is int zhMin ? Math.Max(1, zhMin) : 2;
        _mixedEnglishMinLetters = config.MixedEnglishMinLetters is int enMin ? Math.Max(1, enMin) : 3;
        _warmUpVoices = config.WarmUpVoices ?? true;
        if (_warmUpVoices)
        {
            try
            {
                BeginInvoke(new Action(() => WarmUpVoices()));
            }
            catch { }
        }

        _preventSpeakerSleep = config.PreventSpeakerFromGoingIntoSleepMode ?? false;
        if (_preventSpeakerSleep) StartSpeakerKeepAlive();
        else StopSpeakerKeepAlive();
    }

    private void SpeakText(string text)
    {
        if (!_autoSwitchVoice)
        {
            _voice.SpeakAsync(text);
            return;
        }

        if (string.IsNullOrWhiteSpace(_chineseVoiceName) || string.IsNullOrWhiteSpace(_englishVoiceName))
        {
            _voice.SpeakAsync(text);
            return;
        }

        bool hasChinese = ContainsChinese(text);
        bool hasEnglish = ContainsLatinLetter(text);
        if (hasChinese && !hasEnglish)
        {
            SelectVoiceSafe(_chineseVoiceName);
            _voice.Rate = _chineseVoiceRate;
            _voice.SpeakAsync(text);
            return;
        }
        if (hasEnglish && !hasChinese)
        {
            SelectVoiceSafe(_englishVoiceName);
            _voice.Rate = _englishVoiceRate;
            _voice.SpeakAsync(text);
            return;
        }
        if (!hasChinese && !hasEnglish)
        {
            _voice.SpeakAsync(text);
            return;
        }

        var prompt = BuildMixedVoicePrompt(text);
        if (prompt == null)
        {
            _voice.SpeakAsync(text);
            return;
        }

        _voice.SpeakAsync(prompt);
    }

    private PromptBuilder? BuildMixedVoicePrompt(string text)
    {
        try
        {
            var segments = SplitByLanguage(text);
            segments = SmoothSegments(segments);
            if (segments.Count == 0) return null;

            var prompt = new PromptBuilder();
            foreach (var segment in segments)
            {
                var voiceName = segment.Kind == TextKind.Chinese ? _chineseVoiceName : _englishVoiceName;
                var rate = segment.Kind == TextKind.Chinese ? _chineseVoiceRate : _englishVoiceRate;
                if (string.IsNullOrWhiteSpace(voiceName))
                {
                    prompt.AppendText(segment.Text);
                    continue;
                }

                try
                {
                    prompt.StartVoice(voiceName);
                    var style = new PromptStyle
                    {
                        Rate = MapPromptRate(rate)
                    };
                    prompt.StartStyle(style);
                    prompt.AppendText(segment.Text);
                    prompt.EndStyle();
                    prompt.EndVoice();
                }
                catch
                {
                    prompt.AppendText(segment.Text);
                }
            }
            return prompt;
        }
        catch
        {
            return null;
        }
    }

    private enum TextKind
    {
        Chinese,
        English
    }

    private readonly struct TextSegment
    {
        public TextSegment(TextKind kind, string text)
        {
            Kind = kind;
            Text = text;
        }

        public TextKind Kind { get; }
        public string Text { get; }
    }

    private List<TextSegment> SplitByLanguage(string text)
    {
        var segments = new List<TextSegment>();

        TextKind? currentKind = null;
        var current = new System.Text.StringBuilder();
        var pendingNeutral = new System.Text.StringBuilder();

        foreach (var ch in text)
        {
            var kind = GetCharKind(ch);
            if (kind == null)
            {
                if (currentKind == null) pendingNeutral.Append(ch);
                else current.Append(ch);
                continue;
            }

            if (currentKind == null)
            {
                currentKind = kind;
                if (pendingNeutral.Length > 0)
                {
                    current.Append(pendingNeutral);
                    pendingNeutral.Clear();
                }
                current.Append(ch);
                continue;
            }

            if (currentKind.Value == kind.Value)
            {
                current.Append(ch);
                continue;
            }

            if (current.Length > 0)
            {
                segments.Add(new TextSegment(currentKind.Value, current.ToString()));
                current.Clear();
            }
            currentKind = kind;
            current.Append(ch);
        }

        if (pendingNeutral.Length > 0)
        {
            if (currentKind != null) current.Append(pendingNeutral);
            else return segments;
        }

        if (currentKind != null && current.Length > 0)
        {
            segments.Add(new TextSegment(currentKind.Value, current.ToString()));
        }

        return segments;
    }

    private static TextKind? GetCharKind(char ch)
    {
        if (IsChineseChar(ch)) return TextKind.Chinese;
        if (IsLatinLetter(ch)) return TextKind.English;
        return null;
    }

    private static bool IsLatinLetter(char ch) => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');

    private static bool IsChineseChar(char ch)
    {
        return (ch >= '\u4E00' && ch <= '\u9FFF')
            || (ch >= '\u3400' && ch <= '\u4DBF')
            || (ch >= '\uF900' && ch <= '\uFAFF')
            || (ch >= '\u2E80' && ch <= '\u2EFF')
            || (ch >= '\u3000' && ch <= '\u303F');
    }

    private static bool ContainsChinese(string text)
    {
        foreach (var ch in text)
        {
            if (IsChineseChar(ch)) return true;
        }
        return false;
    }

    private static bool ContainsLatinLetter(string text)
    {
        foreach (var ch in text)
        {
            if (IsLatinLetter(ch)) return true;
        }
        return false;
    }

    private string TryGetDefaultEnglishVoiceName()
    {
        try
        {
            foreach (var voice in _installedVoices)
            {
                if (voice.VoiceInfo.Culture.TwoLetterISOLanguageName.Equals("en", StringComparison.OrdinalIgnoreCase))
                {
                    return voice.VoiceInfo.Name;
                }
            }
        }
        catch { }
        return "";
    }

    private void WarmUpVoices()
    {
        try
        {
            using var warm = new SpeechSynthesizer();
            warm.SetOutputToNull();

            if (!string.IsNullOrWhiteSpace(_chineseVoiceName))
            {
                try
                {
                    warm.SelectVoice(_chineseVoiceName);
                    warm.Rate = _chineseVoiceRate;
                    warm.Speak(" ");
                }
                catch { }
            }

            if (!string.IsNullOrWhiteSpace(_englishVoiceName))
            {
                try
                {
                    warm.SelectVoice(_englishVoiceName);
                    warm.Rate = _englishVoiceRate;
                    warm.Speak(" ");
                }
                catch { }
            }
        }
        catch { }
    }

    private void StartSpeakerKeepAlive()
    {
        try
        {
            if (_speakerKeepAlivePlayer != null) return;

            _speakerKeepAliveStream = new MemoryStream(CreateNearSilentWavBytes(sampleRate: 16000, seconds: 2));
            _speakerKeepAlivePlayer = new SoundPlayer(_speakerKeepAliveStream);
            _speakerKeepAlivePlayer.Load();
            _speakerKeepAlivePlayer.PlayLooping();
        }
        catch
        {
            StopSpeakerKeepAlive();
        }
    }

    private void StopSpeakerKeepAlive()
    {
        try { _speakerKeepAlivePlayer?.Stop(); } catch { }
        try { _speakerKeepAlivePlayer?.Dispose(); } catch { }
        _speakerKeepAlivePlayer = null;
        try { _speakerKeepAliveStream?.Dispose(); } catch { }
        _speakerKeepAliveStream = null;
    }

    private static byte[] CreateNearSilentWavBytes(int sampleRate, int seconds)
    {
        if (sampleRate <= 0) sampleRate = 16000;
        if (seconds <= 0) seconds = 1;

        short channels = 1;
        short bitsPerSample = 16;
        short blockAlign = (short)(channels * (bitsPerSample / 8));
        int byteRate = sampleRate * blockAlign;
        int sampleCount = sampleRate * seconds;
        int dataSize = sampleCount * blockAlign;
        int riffSize = 36 + dataSize;

        var bytes = new byte[44 + dataSize];
        int offset = 0;

        void WriteAscii(string s)
        {
            for (int i = 0; i < s.Length; i++) bytes[offset++] = (byte)s[i];
        }

        void WriteInt32(int v)
        {
            bytes[offset++] = (byte)(v & 0xFF);
            bytes[offset++] = (byte)((v >> 8) & 0xFF);
            bytes[offset++] = (byte)((v >> 16) & 0xFF);
            bytes[offset++] = (byte)((v >> 24) & 0xFF);
        }

        void WriteInt16(short v)
        {
            bytes[offset++] = (byte)(v & 0xFF);
            bytes[offset++] = (byte)((v >> 8) & 0xFF);
        }

        WriteAscii("RIFF");
        WriteInt32(riffSize);
        WriteAscii("WAVE");
        WriteAscii("fmt ");
        WriteInt32(16);
        WriteInt16(1);
        WriteInt16(channels);
        WriteInt32(sampleRate);
        WriteInt32(byteRate);
        WriteInt16(blockAlign);
        WriteInt16(bitsPerSample);
        WriteAscii("data");
        WriteInt32(dataSize);

        short amp = 1;
        for (int i = 0; i < sampleCount; i++)
        {
            short sample = (short)((i & 1) == 0 ? amp : -amp);
            WriteInt16(sample);
        }

        return bytes;
    }

    private string ResolveVoiceName(string desired)
    {
        if (string.IsNullOrWhiteSpace(desired)) return "";
        try
        {
            var desiredTrim = desired.Trim();
            var exact = _installedVoices.FirstOrDefault(v => string.Equals(v.VoiceInfo.Name, desiredTrim, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact.VoiceInfo.Name;

            var startsWith = _installedVoices.FirstOrDefault(v => v.VoiceInfo.Name.StartsWith(desiredTrim, StringComparison.OrdinalIgnoreCase));
            if (startsWith != null) return startsWith.VoiceInfo.Name;

            var contains = _installedVoices.FirstOrDefault(v => v.VoiceInfo.Name.IndexOf(desiredTrim, StringComparison.OrdinalIgnoreCase) >= 0);
            if (contains != null) return contains.VoiceInfo.Name;
        }
        catch { }
        return desired.Trim();
    }

    private void SelectVoiceSafe(string voiceName)
    {
        if (string.IsNullOrWhiteSpace(voiceName)) return;
        try { _voice.SelectVoice(voiceName); } catch { }
    }

    private static int ClampRate(int rate)
    {
        if (rate < -10) return -10;
        if (rate > 10) return 10;
        return rate;
    }

    private static PromptRate MapPromptRate(int rate)
    {
        if (rate <= -6) return PromptRate.ExtraSlow;
        if (rate <= -2) return PromptRate.Slow;
        if (rate <= 1) return PromptRate.Medium;
        if (rate <= 5) return PromptRate.Fast;
        return PromptRate.ExtraFast;
    }

    private List<TextSegment> SmoothSegments(List<TextSegment> segments)
    {
        if (segments.Count <= 1) return segments;

        string leading = "";
        int startIndex = 0;

        if (segments.Count > 0)
        {
            var first = segments[0];
            if (first.Kind == TextKind.Chinese && CountChineseChars(first.Text) < _mixedChineseMinChars)
            {
                leading = first.Text;
                startIndex = 1;
            }
            else if (first.Kind == TextKind.English && CountLatinLetters(first.Text) < _mixedEnglishMinLetters)
            {
                leading = first.Text;
                startIndex = 1;
            }
        }

        var result = new List<TextSegment>();
        for (int i = startIndex; i < segments.Count; i++)
        {
            var seg = segments[i];
            if (result.Count == 0 && !string.IsNullOrEmpty(leading))
            {
                seg = new TextSegment(seg.Kind, leading + seg.Text);
                leading = "";
            }

            bool isShort = seg.Kind == TextKind.Chinese
                ? CountChineseChars(seg.Text) < _mixedChineseMinChars
                : CountLatinLetters(seg.Text) < _mixedEnglishMinLetters;

            if (isShort && result.Count > 0)
            {
                var prev = result[result.Count - 1];
                result[result.Count - 1] = new TextSegment(prev.Kind, prev.Text + seg.Text);
                continue;
            }

            if (result.Count > 0)
            {
                var prev = result[result.Count - 1];
                if (prev.Kind == seg.Kind)
                {
                    result[result.Count - 1] = new TextSegment(prev.Kind, prev.Text + seg.Text);
                    continue;
                }
            }

            result.Add(seg);
        }

        if (!string.IsNullOrEmpty(leading))
        {
            if (result.Count == 0) return segments;
            var first = result[0];
            result[0] = new TextSegment(first.Kind, leading + first.Text);
        }

        return result;
    }

    private static int CountLatinLetters(string text)
    {
        int count = 0;
        foreach (var ch in text)
        {
            if (IsLatinLetter(ch)) count++;
        }
        return count;
    }

    private static int CountChineseChars(string text)
    {
        int count = 0;
        foreach (var ch in text)
        {
            if (IsChineseChar(ch)) count++;
        }
        return count;
    }
}

// 配置类
public class AppConfig
{
    public int HotkeyIndex { get; set; }
    public int StopHotkeyIndex { get; set; } = 1; // 默认 F2
    public object AutoReadHotkeyIndex { get; set; } = 0; // 支持 string (自定义) 或 int (F1-F12 索引)
    public string VoiceName { get; set; } = "";
    public string ChineseVoiceName { get; set; } = "";
    public string EnglishVoiceName { get; set; } = "";
    public bool AutoSwitchVoice { get; set; } = true;
    public int Rate { get; set; }
    public int? ChineseVoiceRate { get; set; }
    public int? EnglishVoiceRate { get; set; }
    public int? MixedChineseMinChars { get; set; }
    public int? MixedEnglishMinLetters { get; set; }
    public bool? WarmUpVoices { get; set; }
    [JsonPropertyName("Prevent the speaker from going into sleep mode")]
    public bool? PreventSpeakerFromGoingIntoSleepMode { get; set; }
    public int Volume { get; set; } = 100;
    public object EscapeRules { get; set; } = "";
    public bool TopMost { get; set; } = false;
    public bool OnlyReadFirstLine { get; set; } = false;
    public bool ClearClipboard { get; set; } = true;
    public int RecordNumber { get; set; } = 1000;
    public bool RemoveSpaces { get; set; } = false;
    public int AutoReadDelay { get; set; } = 1000;
}

// 朗读记录类
public class ReadingRecord
{
    public DateTime Time { get; set; }
    public string Text { get; set; } = "";
}
