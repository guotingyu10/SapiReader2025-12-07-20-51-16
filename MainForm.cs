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

    // 配置文件路径 - 使用项目根目录
    private string _configPath = "";

    public MainForm()
    {
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
        InitializeVoice();
        InitializeTrayIcon();
        LoadConfig();
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
                         
                         // 简单的视觉反馈，让用户知道操作成功
                         if (_configEditor.Text.Contains("\"VoiceName\": \"\""))
                         {
                              // 如果是空配置，提示用户可以粘贴
                              MessageBox.Show($"已复制: {name}\r\n请在 VoiceName 字段的双引号中粘贴 (Ctrl+V)", "复制成功");
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

            _voice.SpeakAsync(processedText);
            
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

            _voice.SpeakAsync(processedText);
            
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

    private void ParseEscapeRules(string input)
    {
        _escapeRules.Clear();
        try
        {
            if (string.IsNullOrWhiteSpace(input)) return;
            var rules = input.Split(';');
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
                        _escapeRules.Add(original, replacement);
                    }
                }
            }
        }
        catch { }
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
                File.WriteAllText(_configPath, jsonContent);
                
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
            if (config.Rate >= -10 && config.Rate <= 10) _voice.Rate = config.Rate;
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

        ParseEscapeRules(config.EscapeRules);
        this.TopMost = config.TopMost;
    }
}

// 配置类
public class AppConfig
{
    public int HotkeyIndex { get; set; }
    public int StopHotkeyIndex { get; set; } = 1; // 默认 F2
    public object AutoReadHotkeyIndex { get; set; } = 0; // 支持 string (自定义) 或 int (F1-F12 索引)
    public string VoiceName { get; set; } = "";
    public int Rate { get; set; }
    public int Volume { get; set; } = 100;
    public string EscapeRules { get; set; } = "";
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
