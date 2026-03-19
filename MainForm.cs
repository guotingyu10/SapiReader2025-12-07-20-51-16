using System.Runtime.InteropServices;
using System.Speech.AudioFormat;
using System.Speech.Synthesis;
using System.Text.Json;
using System.Diagnostics;
using System.Media;
using System.Net;
using System.Net.Sockets;
using System.Text;
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

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO
    {
        public uint cbSize;
        public uint flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    [DllImport("user32.dll")]
    private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLengthW(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, out int wParam, out int lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

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
    private const int EM_GETSEL = 0x00B0;
    private const int SCI_GETSELTEXT = 2161;
    private const int WM_GETTEXT = 0x000D;
    private const int WM_GETTEXTLENGTH = 0x000E;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    #endregion

    private SpeechSynthesizer _voice = null!;
    private SpeechSynthesizer? _standbyVoice;
    private int _standbySynthInProgress = 0;
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
    private bool _mixedGranularOutput = true;
    private double _pronunciationInAdvanceSeconds = 0.0;
    private double _pronunciationEnglishInAdvanceSeconds = 0.0;
    private double _pronunciationChineseInAdvanceSeconds = 0.0;
    private bool _warmUpVoices = true;
    private bool _preventSpeakerSleep = false;
    private SoundPlayer? _speakerKeepAlivePlayer;
    private MemoryStream? _speakerKeepAliveStream;
    private CancellationTokenSource? _operationCts;
    private bool _prebufferMixedAudio = false;
    private int _prebufferMinTextLength = 400;
    private int _prebufferMaxSegmentChars = 200;
    private CancellationTokenSource? _bufferedPlaybackCts;
    private SoundPlayer? _bufferedCurrentPlayer;
    private SoundPlayer? _bufferedReusablePlayer;
    private WaveOutPlayer? _bufferedWavePlayer;
    private int _forceResetSynthAfterStopMs = 800;
    private AppConfig? _lastAppliedConfig;
    private CancellationTokenSource? _stopResetCts;
    private int _disposeOldSynthTimeoutMs = 50;

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
    private readonly object _logLock = new();
    private string _logPath = "debug.log";
    private long _logMaxBytes = 5L * 1024 * 1024;
    private bool _logEnabled = true;
    private TcpListener? _controlListener;
    private CancellationTokenSource? _controlListenerCts;
    private int _recreateSynthInProgress = 0;
    private TaskCompletionSource<bool>? _recreateSynthTcs;
    private int _hotkeyInProgress = 0;

    public MainForm()
    {
        _configPath = ResolveConfigPath();

        try
        {
            var logDir = Path.GetDirectoryName(_configPath);
            if (string.IsNullOrWhiteSpace(logDir)) logDir = AppContext.BaseDirectory;
            _logPath = Path.Combine(logDir, "sapi-reader.log");
        }
        catch { }

        TryLoadLogEnabledFromConfig();
        Log($"AppStart configPath={_configPath}");
        
        InitializeComponent();
        Log("InitializeComponent Done");
        InitializeVoice();
        Log("InitializeVoice Done");
        InitializeTrayIcon();
        Log("InitializeTrayIcon Done");
        LoadConfig();
        Log("LoadConfig Done");
        InitializeConfigWatcher();
        StartControlApi();
    }

    private static string ResolveConfigPath()
    {
        try
        {
            var env = Environment.GetEnvironmentVariable("SAPI_READER_CONFIG");
            if (!string.IsNullOrWhiteSpace(env))
            {
                var p = Path.GetFullPath(env);
                if (File.Exists(p)) return p;
            }
        }
        catch { }

        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddCandidatesFrom(string dir, int maxLevels)
        {
            try
            {
                var current = dir;
                for (int i = 0; i < maxLevels && !string.IsNullOrWhiteSpace(current); i++)
                {
                    var path = Path.Combine(current, "config.json");
                    if (File.Exists(path))
                    {
                        try { candidates.Add(Path.GetFullPath(path)); } catch { candidates.Add(path); }
                    }
                    var parent = Directory.GetParent(current);
                    if (parent == null) break;
                    current = parent.FullName;
                }
            }
            catch { }
        }

        AddCandidatesFrom(Directory.GetCurrentDirectory(), 6);
        AddCandidatesFrom(AppContext.BaseDirectory, 6);

        if (candidates.Count == 0)
        {
            return Path.Combine(AppContext.BaseDirectory, "config.json");
        }

        string best = candidates.First();
        DateTime bestTime = DateTime.MinValue;
        foreach (var c in candidates)
        {
            DateTime t;
            try { t = File.GetLastWriteTimeUtc(c); }
            catch { continue; }
            if (t > bestTime)
            {
                bestTime = t;
                best = c;
            }
        }
        return best;
    }

    private void Log(string message)
    {
        if (!_logEnabled) return;
        try
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [T{Environment.CurrentManagedThreadId}] {message}\n";
            lock (_logLock)
            {
                try
                {
                    var fi = new FileInfo(_logPath);
                    if (fi.Exists && fi.Length >= _logMaxBytes)
                    {
                        string rotated = _logPath + ".1";
                        try { if (File.Exists(rotated)) File.Delete(rotated); } catch { }
                        try { File.Move(_logPath, rotated); } catch { }
                    }
                }
                catch { }

                try { File.AppendAllText(_logPath, line); } catch { }
            }
        }
        catch { }
    }

    private void TryLoadLogEnabledFromConfig()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_configPath)) return;
            if (!File.Exists(_configPath)) return;
            var json = File.ReadAllText(_configPath);
            var options = new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip
            };
            using var doc = JsonDocument.Parse(json, options);
            if (doc.RootElement.ValueKind != JsonValueKind.Object) return;
            if (doc.RootElement.TryGetProperty("sapi-reader.log", out var prop)
                && (prop.ValueKind == JsonValueKind.True || prop.ValueKind == JsonValueKind.False))
            {
                _logEnabled = prop.GetBoolean();
            }
        }
        catch { }
    }

    protected override void OnLoad(EventArgs e)
    {
        Log("OnLoad Started");
        TrySetRealtimePriorityOnStartup();
        base.OnLoad(e);
        Log("OnLoad Done");
    }

    private void TrySetRealtimePriorityOnStartup()
    {
        try
        {
            var proc = Process.GetCurrentProcess();
            var before = proc.PriorityClass;

            try
            {
                proc.PriorityClass = ProcessPriorityClass.RealTime;
            }
            catch
            {
                try { proc.PriorityClass = ProcessPriorityClass.High; }
                catch { }
            }

            var after = proc.PriorityClass;
            Log($"ProcessPriority before={before} after={after}");
        }
        catch (Exception ex)
        {
            Log($"ProcessPriority error {ex.GetType().Name}: {ex.Message}");
        }
    }

    protected override void OnShown(EventArgs e)
    {
        Log("OnShown Started");
        base.OnShown(e);
        Log("OnShown Done");
    }

    private void StartControlApi()
    {
        try
        {
            if (_controlListener != null) return;
            _controlListenerCts = new CancellationTokenSource();
            _controlListener = new TcpListener(IPAddress.Loopback, 32123);
            _controlListener.Start();
            Log("ControlApi started 127.0.0.1:32123");
            var token = _controlListenerCts.Token;
            Task.Run(() => ControlAcceptLoop(token), token);
        }
        catch (Exception ex)
        {
            Log($"ControlApi start error {ex.GetType().Name}: {ex.Message}");
        }
    }

    private void StopControlApi()
    {
        try { _controlListenerCts?.Cancel(); } catch { }
        try { _controlListener?.Stop(); } catch { }
        _controlListener = null;
        try { _controlListenerCts?.Dispose(); } catch { }
        _controlListenerCts = null;
    }

    private async Task ControlAcceptLoop(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            TcpClient? client = null;
            try
            {
                if (_controlListener == null) return;
                client = await _controlListener.AcceptTcpClientAsync(token);
                _ = Task.Run(() => HandleControlClient(client, token), token);
            }
            catch (OperationCanceledException)
            {
                try { client?.Dispose(); } catch { }
                return;
            }
            catch (Exception ex)
            {
                try { client?.Dispose(); } catch { }
                Log($"ControlApi accept error {ex.GetType().Name}: {ex.Message}");
                await Task.Delay(200, CancellationToken.None);
            }
        }
    }

    private async Task HandleControlClient(TcpClient client, CancellationToken token)
    {
        try
        {
            client.NoDelay = true;
            using var stream = client.GetStream();
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: 8192, leaveOpen: true);
            using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), bufferSize: 8192, leaveOpen: true)
            {
                AutoFlush = true
            };

            while (!token.IsCancellationRequested)
            {
                string? line;
                try
                {
                    line = await reader.ReadLineAsync(token);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
                if (line == null) return;
                if (string.IsNullOrWhiteSpace(line)) continue;

                string cmd = "";
                JsonElement root;
                try
                {
                    using var doc = JsonDocument.Parse(line);
                    root = doc.RootElement.Clone();
                    if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("cmd", out var cmdProp) && cmdProp.ValueKind == JsonValueKind.String)
                    {
                        cmd = cmdProp.GetString() ?? "";
                    }
                }
                catch (Exception ex)
                {
                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = false, error = $"{ex.GetType().Name}: {ex.Message}" }));
                    continue;
                }

                if (cmd == "ping")
                {
                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true }));
                    continue;
                }

                if (cmd == "get_state")
                {
                    try
                    {
                        var state = _voice != null ? _voice.State.ToString() : "null";
                        await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true, voiceState = state, isSpeaking = _isSpeaking, isAutoReading = _isAutoReading }));
                    }
                    catch
                    {
                        await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true, voiceState = "unknown", isSpeaking = _isSpeaking, isAutoReading = _isAutoReading }));
                    }
                    continue;
                }

                if (cmd == "set_keepalive")
                {
                    bool enable = false;
                    try
                    {
                        if (root.TryGetProperty("enable", out var p) && (p.ValueKind == JsonValueKind.True || p.ValueKind == JsonValueKind.False))
                        {
                            enable = p.GetBoolean();
                        }
                    }
                    catch { }

                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            _preventSpeakerSleep = enable;
                            if (enable) StartSpeakerKeepAlive();
                            else StopSpeakerKeepAlive();
                        }));
                    }
                    catch { }

                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true }));
                    continue;
                }

                if (cmd == "stop")
                {
                    bool hardReset = true;
                    int waitMs = 0;
                    try
                    {
                        if (root.TryGetProperty("hardReset", out var p) && (p.ValueKind == JsonValueKind.True || p.ValueKind == JsonValueKind.False))
                        {
                            hardReset = p.GetBoolean();
                        }
                        if (root.TryGetProperty("waitMs", out var w) && w.ValueKind == JsonValueKind.Number)
                        {
                            waitMs = w.GetInt32();
                        }
                    }
                    catch { }

                    try
                    {
                        BeginInvoke(new Action(() => StopSpeakingInternal(updateStatus: false, hardReset: hardReset)));
                    }
                    catch { }
                    if (waitMs > 0)
                    {
                        try { await Task.Delay(waitMs, token); } catch { }
                    }
                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true }));
                    continue;
                }

                if (cmd == "set_voice")
                {
                    string voiceName = "";
                    int? volume = null;
                    int? rate = null;
                    bool? autoSwitch = null;

                    try
                    {
                        if (root.TryGetProperty("voice", out var v) && v.ValueKind == JsonValueKind.String) voiceName = v.GetString() ?? "";
                        if (root.TryGetProperty("volume", out var vol) && vol.ValueKind == JsonValueKind.Number) volume = vol.GetInt32();
                        if (root.TryGetProperty("rate", out var r) && r.ValueKind == JsonValueKind.Number) rate = r.GetInt32();
                        if (root.TryGetProperty("autoSwitchVoice", out var asv) && (asv.ValueKind == JsonValueKind.True || asv.ValueKind == JsonValueKind.False))
                        {
                            autoSwitch = asv.GetBoolean();
                        }
                    }
                    catch { }

                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            if (autoSwitch.HasValue) _autoSwitchVoice = autoSwitch.Value;
                            if (!string.IsNullOrWhiteSpace(voiceName)) SelectVoiceSafe(voiceName);
                            if (volume.HasValue && volume.Value >= 0 && volume.Value <= 100) _voice.Volume = volume.Value;
                            if (rate.HasValue) _voice.Rate = ClampRate(rate.Value);
                        }));
                    }
                    catch { }
                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true }));
                    continue;
                }

                if (cmd == "set_mixed_voices")
                {
                    string chineseVoice = "";
                    string englishVoice = "";
                    int? chineseRate = null;
                    int? englishRate = null;
                    int? volume = null;
                    int? mixedChineseMinChars = null;
                    int? mixedEnglishMinLetters = null;
                    bool warmUp = true;

                    try
                    {
                        if (root.TryGetProperty("chineseVoice", out var c) && c.ValueKind == JsonValueKind.String) chineseVoice = c.GetString() ?? "";
                        if (root.TryGetProperty("englishVoice", out var e) && e.ValueKind == JsonValueKind.String) englishVoice = e.GetString() ?? "";
                        if (root.TryGetProperty("chineseRate", out var cr) && cr.ValueKind == JsonValueKind.Number) chineseRate = cr.GetInt32();
                        if (root.TryGetProperty("englishRate", out var er) && er.ValueKind == JsonValueKind.Number) englishRate = er.GetInt32();
                        if (root.TryGetProperty("volume", out var v) && v.ValueKind == JsonValueKind.Number) volume = v.GetInt32();
                        if (root.TryGetProperty("mixedChineseMinChars", out var m1) && m1.ValueKind == JsonValueKind.Number) mixedChineseMinChars = m1.GetInt32();
                        if (root.TryGetProperty("mixedEnglishMinLetters", out var m2) && m2.ValueKind == JsonValueKind.Number) mixedEnglishMinLetters = m2.GetInt32();
                        if (root.TryGetProperty("warmUp", out var w) && (w.ValueKind == JsonValueKind.True || w.ValueKind == JsonValueKind.False)) warmUp = w.GetBoolean();
                    }
                    catch { }

                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            _autoSwitchVoice = true;
                            if (!string.IsNullOrWhiteSpace(chineseVoice)) _chineseVoiceName = ResolveVoiceName(chineseVoice);
                            if (!string.IsNullOrWhiteSpace(englishVoice)) _englishVoiceName = ResolveVoiceName(englishVoice);
                            if (chineseRate.HasValue) _chineseVoiceRate = ClampRate(chineseRate.Value);
                            if (englishRate.HasValue) _englishVoiceRate = ClampRate(englishRate.Value);
                            if (mixedChineseMinChars.HasValue) _mixedChineseMinChars = Math.Max(1, mixedChineseMinChars.Value);
                            if (mixedEnglishMinLetters.HasValue) _mixedEnglishMinLetters = Math.Max(1, mixedEnglishMinLetters.Value);
                            if (volume.HasValue && volume.Value >= 0 && volume.Value <= 100) _voice.Volume = volume.Value;
                        }));
                    }
                    catch { }

                    if (warmUp)
                    {
                        RunStaBackground(() => WarmUpVoices());
                    }

                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true }));
                    continue;
                }

                if (cmd == "speak")
                {
                    string text = "";
                    bool skipProcessing = false;
                    bool skipHistory = false;
                    try
                    {
                        if (root.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) text = t.GetString() ?? "";
                        if (root.TryGetProperty("skipProcessing", out var sp) && (sp.ValueKind == JsonValueKind.True || sp.ValueKind == JsonValueKind.False))
                        {
                            skipProcessing = sp.GetBoolean();
                        }
                        if (root.TryGetProperty("skipHistory", out var sh) && (sh.ValueKind == JsonValueKind.True || sh.ValueKind == JsonValueKind.False))
                        {
                            skipHistory = sh.GetBoolean();
                        }
                    }
                    catch { }

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = false, error = "empty_text" }));
                        continue;
                    }

                    try
                    {
                        BeginInvoke(new Action(() => SpeakFromApi(text, skipProcessing, skipHistory)));
                    }
                    catch { }
                    await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = true }));
                    continue;
                }

                await writer.WriteLineAsync(JsonSerializer.Serialize(new { ok = false, error = "unknown_cmd" }));
            }
        }
        catch (Exception ex)
        {
            Log($"ControlApi client error {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            try { client.Dispose(); } catch { }
        }
    }

    private async void SpeakFromApi(string text, bool skipProcessing, bool skipHistory)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            _isAutoReading = false;
            CancelCurrentOperation();
            StopBufferedPlayback();
            CancelPendingStopReset();
            try { _voice?.SpeakAsyncCancelAll(); } catch { }
            _isSpeaking = false;
            string rawText = text;
            if (_onlyReadFirstLine)
            {
                int newlineIndex = rawText.IndexOfAny(new[] { '\r', '\n' });
                if (newlineIndex >= 0) rawText = rawText.Substring(0, newlineIndex);
            }
            if (_removeSpaces)
            {
                rawText = rawText.Replace(" ", "").Replace("\t", "").Replace("\u3000", "");
            }

            var cts = new CancellationTokenSource();
            _operationCts = cts;
            var token = cts.Token;

            string processedText;
            if (skipProcessing)
            {
                _statusLabel.Text = "正在朗读...";
                _statusLabel.ForeColor = Color.Blue;
                processedText = rawText;
            }
            else
            {
                _statusLabel.Text = "正在处理文本...";
                _statusLabel.ForeColor = Color.Blue;
                processedText = await Task.Run(() => ApplyEscapeRules(rawText, token), token);
                if (token.IsCancellationRequested) return;
            }

            if (!skipHistory) AddReadingRecord(processedText);
            SpeakText(processedText);
            Log($"ApiSpeak len={processedText.Length} skipProcessing={skipProcessing} ms={sw.ElapsedMilliseconds}");
        }
        catch (Exception ex)
        {
            _isSpeaking = false;
            Log($"ApiSpeak error {ex.GetType().Name}: {ex.Message}");
            _statusLabel.Text = $"错误: {ex.Message}";
            _statusLabel.ForeColor = Color.Red;
        }
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
        try { _voice.SetOutputToDefaultAudioDevice(); } catch { }
        // 获取发音人列表供查询
        _installedVoices = _voice.GetInstalledVoices().Where(v => v.Enabled).ToList();
        AttachVoiceEvents(_voice);
        EnsureStandbySynth();
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
        Log($"RegisterHotKey read ok={ok} vk={VkToDisplay(_hotkeyKey)}");
        if (!ok)
        {
            _statusLabel.Text = $"朗读热键({VkToDisplay(_hotkeyKey)})注册失败！可能被占用";
            _statusLabel.ForeColor = Color.Red;
        }
        
        // 注册停止热键
        bool stopOk = RegisterHotKey(this.Handle, STOP_HOTKEY_ID, 0, _stopHotkeyKey);
        Log($"RegisterHotKey stop ok={stopOk} vk={VkToDisplay(_stopHotkeyKey)}");
        if (!stopOk)
        {
            _statusLabel.Text = $"停止热键({VkToDisplay(_stopHotkeyKey)})注册失败！可能被占用";
            _statusLabel.ForeColor = Color.Red;
        }
        
        // 注册自动朗读热键
        bool autoReadOk = RegisterHotKey(this.Handle, AUTO_READ_HOTKEY_ID, _autoReadModifiers, _autoReadHotkeyKey);
        Log($"RegisterHotKey autoRead ok={autoReadOk} mods={_autoReadModifiers} vk={VkToDisplay(_autoReadHotkeyKey)}");
        
        if (ok && stopOk)
        {
            UpdateStatusLabel();
        }
    }

    private void UpdateStatusLabel()
    {
        string keyStr = VkToDisplay(_hotkeyKey);
        string stopKeyStr = VkToDisplay(_stopHotkeyKey);
        _statusLabel.Text = $"配置已应用 - 按 {keyStr} 朗读，按 {stopKeyStr} 停止";
        _statusLabel.ForeColor = Color.Green;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY)
        {
            int id = m.WParam.ToInt32();
            Log($"WM_HOTKEY id={id} isSpeaking={_isSpeaking} state={_voice?.State} isAutoReading={_isAutoReading}");
            if (id == HOTKEY_ID)
            {
                HandleHotkey();
            }
            else if (id == STOP_HOTKEY_ID)
            {
                StopSpeaking();
            }
            else if (id == AUTO_READ_HOTKEY_ID)
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
        if (Interlocked.Exchange(ref _hotkeyInProgress, 1) == 1)
        {
            Log("HandleHotkey skipped reason=in_progress");
            return;
        }
        var sw = Stopwatch.StartNew();
        try
        {
            Log("HandleHotkey start");
            var stopTask = StopSpeakingFastAsync(updateStatus: false);

            string text = await GetSelectedTextPreferUiaAsync(timeoutMs: 350);

            if (string.IsNullOrWhiteSpace(text))
            {
                _statusLabel.Text = "未获取到选中文本";
                _statusLabel.ForeColor = Color.Orange;
                Log("HandleHotkey selection empty");
                await Task.Delay(2000);
                UpdateStatusLabel();
                return;
            }
            Log($"HandleHotkey selection length={text.Length} selection_ms={sw.ElapsedMilliseconds}");
            
            if (_onlyReadFirstLine)
            {
                int newlineIndex = text.IndexOfAny(new[] { '\r', '\n' });
                if (newlineIndex >= 0) text = text.Substring(0, newlineIndex);
            }

            if (_removeSpaces)
            {
                text = text.Replace(" ", "").Replace("\t", "").Replace("\u3000", "");
            }

            await TrySetClipboardTextWithRetryAsync(text, retries: 15, delayMs: 30);

            CancelCurrentOperation();
            var cts = new CancellationTokenSource();
            _operationCts = cts;
            var token = cts.Token;

            _statusLabel.Text = "正在处理文本...";
            _statusLabel.ForeColor = Color.Blue;

            string processedText = await Task.Run(() => ApplyEscapeRules(text, token), token);
            if (token.IsCancellationRequested) return;
            Log($"HandleHotkey processed length={processedText.Length} process_ms={sw.ElapsedMilliseconds}");

            AddReadingRecord(processedText);

            await stopTask;
            Log($"HandleHotkey stop_ms={sw.ElapsedMilliseconds}");

            SpeakText(processedText);
            Log($"HandleHotkey speak_called_ms={sw.ElapsedMilliseconds}");
        }
        catch (Exception ex)
        {
            _isSpeaking = false;
            Log($"HandleHotkey error {ex.GetType().Name}: {ex.Message}");
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
        finally
        {
            Interlocked.Exchange(ref _hotkeyInProgress, 0);
        }
    }

    private async Task<string> GetSelectedTextPreferUiaAsync(int timeoutMs)
    {
        string? text = null;
        try
        {
            var task = Task.Run(() => GetSelectedTextFromFocusedControlWithTimeout(timeoutMs));
            var completed = await Task.WhenAny(task, Task.Delay(timeoutMs));
            if (completed == task) text = await task;
        }
        catch { }
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return await GetSelectedTextViaClipboardFallbackAsync();
    }

    private string? GetSelectedTextFromFocusedControlWithTimeout(int timeoutMs)
    {
        try
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return null;
            var threadId = GetWindowThreadProcessId(hwnd, out _);
            var info = new GUITHREADINFO { cbSize = (uint)Marshal.SizeOf<GUITHREADINFO>() };
            if (!GetGUIThreadInfo(threadId, ref info)) return null;
            if (info.hwndFocus == IntPtr.Zero) return null;

            var cls = new StringBuilder(256);
            _ = GetClassNameW(info.hwndFocus, cls, cls.Capacity);
            var className = cls.ToString();
            if (className.IndexOf("Scintilla", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return GetSelectedTextFromScintilla(info.hwndFocus, timeoutMs);
            }
            if (className.Equals("Edit", StringComparison.OrdinalIgnoreCase) || className.IndexOf("RichEdit", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return GetSelectedTextFromEditLike(info.hwndFocus, timeoutMs);
            }
        }
        catch { }

        return null;
    }

    private string? GetSelectedTextFromEditLike(IntPtr hwnd, int timeoutMs)
    {
        try
        {
            var startPtr = Marshal.AllocHGlobal(sizeof(int));
            var endPtr = Marshal.AllocHGlobal(sizeof(int));
            try
            {
                Marshal.WriteInt32(startPtr, 0);
                Marshal.WriteInt32(endPtr, 0);
                if (SendMessageTimeoutW(hwnd, EM_GETSEL, startPtr, endPtr, SMTO_ABORTIFHUNG, (uint)timeoutMs, out _) == IntPtr.Zero)
                {
                    return null;
                }
                int start = Marshal.ReadInt32(startPtr);
                int end = Marshal.ReadInt32(endPtr);
            if (end <= start) return null;

                if (SendMessageTimeoutW(hwnd, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, (uint)timeoutMs, out var lenRes) == IntPtr.Zero)
                {
                    return null;
                }
                int len = lenRes.ToInt32();
                if (len <= 0) return null;
                var buf = Marshal.AllocHGlobal((len + 1) * 2);
                try
                {
                    if (SendMessageTimeoutW(hwnd, WM_GETTEXT, (IntPtr)(len + 1), buf, SMTO_ABORTIFHUNG, (uint)timeoutMs, out _) == IntPtr.Zero)
                    {
                        return null;
                    }
                    var full = Marshal.PtrToStringUni(buf, len) ?? "";

            if (start < 0) start = 0;
            if (end > full.Length) end = full.Length;
            if (end <= start) return null;
            return full.Substring(start, end - start);
                }
                finally
                {
                    Marshal.FreeHGlobal(buf);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(startPtr);
                Marshal.FreeHGlobal(endPtr);
            }
        }
        catch
        {
            return null;
        }
    }

    private string? GetSelectedTextFromScintilla(IntPtr hwnd, int timeoutMs)
    {
        try
        {
            if (SendMessageTimeoutW(hwnd, SCI_GETSELTEXT, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, (uint)timeoutMs, out var lenRes) == IntPtr.Zero)
            {
                return null;
            }
            int byteLen = lenRes.ToInt32();
            if (byteLen <= 1) return null;
            var buf = Marshal.AllocHGlobal(byteLen + 2);
            try
            {
                if (SendMessageTimeoutW(hwnd, SCI_GETSELTEXT, IntPtr.Zero, buf, SMTO_ABORTIFHUNG, (uint)timeoutMs, out _) == IntPtr.Zero)
                {
                    return null;
                }
                var bytes = new byte[byteLen];
                Marshal.Copy(buf, bytes, 0, byteLen);
                var s = Encoding.UTF8.GetString(bytes).TrimEnd('\0');
                if (string.IsNullOrWhiteSpace(s)) return null;
                return s;
            }
            finally
            {
                Marshal.FreeHGlobal(buf);
            }
        }
        catch
        {
            return null;
        }
    }

    private async Task<string> GetSelectedTextViaClipboardFallbackAsync()
    {
        try { SimulateCtrlC(); } catch { }

        for (int i = 0; i < 25; i++)
        {
            await Task.Delay(20);
            try
            {
                if (Clipboard.ContainsText())
                {
                    var t = Clipboard.GetText();
                    if (!string.IsNullOrWhiteSpace(t)) return t;
                }
            }
            catch (ExternalException)
            {
            }
            catch
            {
            }
        }

        return "";
    }

    private async Task<bool> TrySetClipboardTextWithRetryAsync(string text, int retries, int delayMs)
    {
        for (int i = 0; i < retries; i++)
        {
            try
            {
                Clipboard.SetText(text);
                return true;
            }
            catch (ExternalException)
            {
                await Task.Delay(delayMs);
            }
            catch
            {
                await Task.Delay(delayMs);
            }
        }
        return false;
    }

    private async Task StopSpeakingFastAsync(bool updateStatus)
    {
        _isAutoReading = false;
        CancelCurrentOperation();
        StopBufferedPlayback();
        CancelPendingStopReset();

        if (Volatile.Read(ref _recreateSynthInProgress) == 1)
        {
            await WaitForRecreateSynthAsync(timeoutMs: 6000);
        }

        var voice = _voice;
        bool isSpeakingNow = false;
        try { if (voice != null) isSpeakingNow = voice.State == SynthesizerState.Speaking; } catch { }

        if (isSpeakingNow)
        {
            TryMuteSynthImmediate(voice);
            if (TrySwapInStandbySynth())
            {
                EnsureStandbySynth();
            }
            else
            {
                EnsureStandbySynth();
                bool okStandby = await WaitForStandbySynthAsync(timeoutMs: 6000);
                if (okStandby && TrySwapInStandbySynth())
                {
                    EnsureStandbySynth();
                }
                else
                {
                    RecreateSpeechSynthesizer();
                    bool ok = await WaitForRecreateSynthAsync(timeoutMs: 6000);
                    if (!ok) Log("StopSpeakingFastAsync hard_reset_timeout");
                    EnsureStandbySynth();
                }
            }
        }
        else
        {
            var cancelTask = Task.Run(() =>
            {
                try { voice?.SpeakAsyncCancelAll(); } catch { }
            });

            var completed = await Task.WhenAny(cancelTask, Task.Delay(250));
            if (completed != cancelTask)
            {
                Log("StopSpeakingFastAsync soft_cancel_timeout -> hard_reset_synth");
                TryMuteSynthImmediate(voice);
                if (TrySwapInStandbySynth())
                {
                    EnsureStandbySynth();
                }
                else
                {
                    EnsureStandbySynth();
                    bool okStandby = await WaitForStandbySynthAsync(timeoutMs: 6000);
                    if (okStandby && TrySwapInStandbySynth())
                    {
                        EnsureStandbySynth();
                    }
                    else
                    {
                        RecreateSpeechSynthesizer();
                        bool ok = await WaitForRecreateSynthAsync(timeoutMs: 6000);
                        if (!ok) Log("StopSpeakingFastAsync hard_reset_timeout");
                        EnsureStandbySynth();
                    }
                }
            }
            else
            {
                ScheduleForceResetSynthIfStuck();
            }
        }

        _isSpeaking = false;
        if (updateStatus) UpdateStatusLabel();
    }

    private void TryMuteSynthImmediate(SpeechSynthesizer? synth)
    {
        try { if (synth != null) synth.Volume = 0; } catch { }

        Task.Run(() =>
        {
            try { synth?.SetOutputToNull(); } catch { }
            try { synth?.SpeakAsyncCancelAll(); } catch { }
        });
    }

    private async Task<bool> WaitForRecreateSynthAsync(int timeoutMs)
    {
        try
        {
            var tcs = Volatile.Read(ref _recreateSynthTcs);
            if (tcs == null) return true;
            var completed = await Task.WhenAny(tcs.Task, Task.Delay(timeoutMs));
            if (completed == tcs.Task)
            {
                await tcs.Task;
                return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    private void StopSpeaking()
    {
        StopSpeakingInternal(updateStatus: true, hardReset: true);
    }

    private void StopSpeakingInternal(bool updateStatus, bool hardReset = false)
    {
        _isAutoReading = false;
        Log($"StopSpeakingInternal updateStatus={updateStatus} hardReset={hardReset} isSpeaking={_isSpeaking} state={_voice?.State}");
        CancelCurrentOperation();
        StopBufferedPlayback();
        CancelPendingStopReset();
        if (hardReset)
        {
            Log("StopSpeakingInternal action=hard_reset_synth");
            TryMuteSynthImmediate(_voice);
            if (TrySwapInStandbySynth())
            {
                EnsureStandbySynth();
            }
            else
            {
                RecreateSpeechSynthesizer();
                EnsureStandbySynth();
            }
            _isSpeaking = false;
            if (updateStatus) UpdateStatusLabel();
            return;
        }

        try { _voice?.SpeakAsyncCancelAll(); } catch { }
        ScheduleForceResetSynthIfStuck();
        _isSpeaking = false;
        if (updateStatus) UpdateStatusLabel();
    }

    private void ScheduleForceResetSynthIfStuck()
    {
        try
        {
            int delay = _forceResetSynthAfterStopMs;
            if (delay <= 0) delay = 800;
            try { _stopResetCts?.Cancel(); } catch { }
            try { _stopResetCts?.Dispose(); } catch { }
            _stopResetCts = new CancellationTokenSource();
            var token = _stopResetCts.Token;
            Task.Run(async () =>
            {
                await Task.Delay(delay);
                if (token.IsCancellationRequested) return;
                try
                {
                    BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            if (token.IsCancellationRequested) return;
                            if (_voice != null && _voice.State == SynthesizerState.Speaking)
                            {
                                Log($"ForceResetSynth reason=stuck_after_stop state={_voice.State}");
                                RecreateSpeechSynthesizer();
                            }
                        }
                        catch { }
                    }));
                }
                catch { }
            }, token);
        }
        catch { }
    }

    private void RunStaBackground(Action action)
    {
        try
        {
            var thread = new Thread(() =>
            {
                try { action(); }
                catch (Exception ex) { Log($"StaBackground error {ex.GetType().Name}: {ex.Message}"); }
            })
            {
                IsBackground = true
            };
            try { thread.SetApartmentState(ApartmentState.STA); } catch { }
            thread.Start();
        }
        catch { }
    }

    private void EnsureStandbySynth()
    {
        if (_standbyVoice != null) return;
        if (Interlocked.Exchange(ref _standbySynthInProgress, 1) == 1) return;

        RunStaBackground(() =>
        {
            SpeechSynthesizer? standby = null;
            try
            {
                standby = new SpeechSynthesizer();
                try { standby.SetOutputToDefaultAudioDevice(); } catch { }
                AttachVoiceEvents(standby);
                ApplyBaseConfigToSynth(standby);
            }
            catch
            {
                try { standby?.Dispose(); } catch { }
                standby = null;
            }
            finally
            {
                if (standby != null)
                {
                    var old = Interlocked.Exchange(ref _standbyVoice, standby);
                    DisposeOldSynthAsync(old);
                    Log("StandbySynth ready");
                }
                Interlocked.Exchange(ref _standbySynthInProgress, 0);
            }
        });
    }

    private bool TrySwapInStandbySynth()
    {
        var standby = Interlocked.Exchange(ref _standbyVoice, null);
        if (standby == null) return false;

        var old = Interlocked.Exchange(ref _voice, standby);
        TryMuteSynthImmediate(old);
        DisposeOldSynthAsync(old);
        Log("SwapInStandbySynth ok");
        return true;
    }

    private async Task<bool> WaitForStandbySynthAsync(int timeoutMs)
    {
        if (_standbyVoice != null) return true;
        if (timeoutMs <= 0) return false;

        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            if (_standbyVoice != null) return true;
            await Task.Delay(50);
        }
        return _standbyVoice != null;
    }

    private void RecreateSpeechSynthesizer()
    {
        if (Interlocked.Exchange(ref _recreateSynthInProgress, 1) == 1) return;
        var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        _recreateSynthTcs = tcs;
        var old = _voice;
        RunStaBackground(() =>
        {
            SpeechSynthesizer? next = null;
            try { old?.SetOutputToNull(); } catch { }
            try { if (old != null) old.Volume = 0; } catch { }
            try { old?.SpeakAsyncCancelAll(); } catch { }

            try
            {
                next = new SpeechSynthesizer();
                try { next.SetOutputToDefaultAudioDevice(); } catch { }
                AttachVoiceEvents(next);
                ApplyBaseConfigToSynth(next);
            }
            catch { }

            try
            {
                try
                {
                    if (next != null)
                    {
                        Interlocked.Exchange(ref _voice, next);
                        if (_installedVoices.Count == 0)
                        {
                            try { _installedVoices = next.GetInstalledVoices().Where(v => v.Enabled).ToList(); } catch { }
                        }
                        DisposeOldSynthAsync(old);
                        Log($"RecreateSpeechSynthesizer done state={_voice.State}");
                        tcs.TrySetResult(true);
                    }
                    else
                    {
                        Log("RecreateSpeechSynthesizer failed next=null");
                        tcs.TrySetResult(false);
                    }
                }
                finally
                {
                    Interlocked.Exchange(ref _recreateSynthInProgress, 0);
                }
            }
            catch
            {
                tcs.TrySetResult(false);
                Interlocked.Exchange(ref _recreateSynthInProgress, 0);
            }
        });
    }

    private void CancelCurrentOperation()
    {
        try { _operationCts?.Cancel(); } catch { }
        try { _operationCts?.Dispose(); } catch { }
        _operationCts = null;
    }

    private void CancelPendingStopReset()
    {
        try { _stopResetCts?.Cancel(); } catch { }
        try { _stopResetCts?.Dispose(); } catch { }
        _stopResetCts = null;
    }

    private void AttachVoiceEvents(SpeechSynthesizer synth)
    {
        synth.SpeakStarted += (_, __) =>
        {
            try
            {
                BeginInvoke(new Action(() =>
                {
                    _isSpeaking = true;
                    _statusLabel.Text = "正在朗读...";
                    _statusLabel.ForeColor = Color.Blue;
                }));
            }
            catch { }
        };
        synth.SpeakCompleted += (_, __) =>
        {
            try
            {
                BeginInvoke(new Action(() =>
                {
                    _isSpeaking = false;
                    if (!_isAutoReading) UpdateStatusLabel();
                }));
            }
            catch { }
        };
    }

    private void ApplyBaseConfigToSynth(SpeechSynthesizer synth)
    {
        if (_lastAppliedConfig == null) return;
        try { synth.Volume = _lastAppliedConfig.Volume; } catch { }
        try
        {
            var desiredChinese = string.IsNullOrWhiteSpace(_lastAppliedConfig.ChineseVoiceName) ? _lastAppliedConfig.VoiceName : _lastAppliedConfig.ChineseVoiceName;
            var desiredSingle = !_lastAppliedConfig.AutoSwitchVoice && !string.IsNullOrWhiteSpace(desiredChinese) ? desiredChinese : _lastAppliedConfig.VoiceName;

            try { synth.Rate = ClampRate(_lastAppliedConfig.AutoSwitchVoice ? _lastAppliedConfig.Rate : (_lastAppliedConfig.ChineseVoiceRate ?? _lastAppliedConfig.Rate)); } catch { }

            if (!string.IsNullOrWhiteSpace(desiredSingle))
            {
                try { synth.SelectVoice(desiredSingle); } catch { }
            }
        }
        catch { }
    }

    private void DisposeOldSynthAsync(SpeechSynthesizer? old)
    {
        if (old == null) return;

        try
        {
            var thread = new Thread(() =>
            {
                try { old.Dispose(); }
                catch (Exception ex) { Log($"DisposeOldSynth error {ex.GetType().Name}: {ex.Message}"); }
            })
            {
                IsBackground = true
            };
            try { thread.SetApartmentState(ApartmentState.STA); } catch { }
            thread.Start();

            int timeout = _disposeOldSynthTimeoutMs;
            if (timeout > 0)
            {
                try { thread.Join(timeout); } catch { }
            }
        }
        catch { }
    }

    private void StopBufferedPlayback()
    {
        try { _bufferedPlaybackCts?.Cancel(); } catch { }
        try { _bufferedPlaybackCts?.Dispose(); } catch { }
        _bufferedPlaybackCts = null;

        var wavePlayer = Interlocked.Exchange(ref _bufferedWavePlayer, null);
        Log($"StopBufferedPlayback wavePlayer={(wavePlayer != null)}");
        try { wavePlayer?.Stop(); } catch { }

        try { _bufferedCurrentPlayer?.Stop(); } catch { }
        _bufferedCurrentPlayer = null;
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
            string text = await GetSelectedTextPreferUiaAsync(timeoutMs: 250);

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

            CancelCurrentOperation();
            var cts = new CancellationTokenSource();
            _operationCts = cts;
            var token = cts.Token;
            string processedText = await Task.Run(() => ApplyEscapeRules(text, token), token);
            if (token.IsCancellationRequested) return false;

            AddReadingRecord(processedText);
            
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
        CancelCurrentOperation();
        StopControlApi();
        StopSpeakerKeepAlive();
        base.OnFormClosed(e);
    }

    private void AddReadingRecord(string text)
    {
        // var record = new ReadingRecord
        // {
        //     Time = DateTime.Now,
        //     Text = text.Length > 100 ? text.Substring(0, 100) + "..." : text
        // };
        //
        // if (_readingHistory.Count >= _maxHistory)
        // {
        //     _readingHistory.RemoveAt(0);
        //     _historyListBox.Items.RemoveAt(0);
        // }
        //
        // _readingHistory.Add(record);
        // _historyListBox.Items.Add($"[{record.Time:HH:mm:ss}] {record.Text}");
        // _historyListBox.TopIndex = _historyListBox.Items.Count - 1;
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

    private string ApplyEscapeRules(string text, CancellationToken token)
    {
        if (token.IsCancellationRequested) return "";

        KeyValuePair<string, string>[] rulesSnapshot;
        if (_escapeRulesOrdered.Count == 0 && _escapeRules.Count > 0)
        {
            rulesSnapshot = _escapeRules
                .OrderByDescending(kv => kv.Key.Length)
                .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                .ToArray();
        }
        else
        {
            rulesSnapshot = _escapeRulesOrdered.ToArray();
        }

        if (rulesSnapshot.Length == 0) return text;

        string result = text;
        foreach (var rule in rulesSnapshot)
        {
            if (token.IsCancellationRequested) return "";
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
        _lastAppliedConfig = config;
        // 1. 热键设置
        if (TryParseNoModifierHotkey(config.HotkeyIndex, out uint readVk)) _hotkeyKey = readVk;
        if (TryParseNoModifierHotkey(config.StopHotkeyIndex, out uint stopVk)) _stopHotkeyKey = stopVk;

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
            var desiredChineseForSingle = string.IsNullOrWhiteSpace(config.ChineseVoiceName) ? config.VoiceName : config.ChineseVoiceName;
            var desiredSingle = !config.AutoSwitchVoice && !string.IsNullOrWhiteSpace(desiredChineseForSingle) ? desiredChineseForSingle : config.VoiceName;

            _voice.Rate = ClampRate(config.AutoSwitchVoice ? config.Rate : (config.ChineseVoiceRate ?? config.Rate));
            if (config.Volume >= 0 && config.Volume <= 100) _voice.Volume = config.Volume;
            
            if (!string.IsNullOrEmpty(desiredSingle))
            {
                try
                {
                    _voice.SelectVoice(desiredSingle);
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
        _mixedGranularOutput = config.MixedGranularOutput ?? true;
        _warmUpVoices = config.WarmUpVoices ?? true;
        if (_warmUpVoices)
        {
            try
            {
                RunStaBackground(() => WarmUpVoices());
            }
            catch { }
        }

        _preventSpeakerSleep = config.PreventSpeakerFromGoingIntoSleepMode ?? false;
        if (_preventSpeakerSleep) StartSpeakerKeepAlive();
        else StopSpeakerKeepAlive();

        _prebufferMixedAudio = config.PrebufferMixedAudio ?? false;
        _prebufferMinTextLength = config.PrebufferMinTextLength is int minLen ? Math.Max(1, minLen) : 400;
        _prebufferMaxSegmentChars = config.PrebufferMaxSegmentChars is int maxChars ? Math.Max(50, maxChars) : 200;
        try
        {
            double sec = config.PronunciationInAdvance ?? 0.0;
            if (double.IsNaN(sec) || double.IsInfinity(sec)) sec = 0.0;
            if (sec < 0) sec = 0.0;
            if (sec > 0.25) sec = 0.25;
            _pronunciationInAdvanceSeconds = sec;
        }
        catch
        {
            _pronunciationInAdvanceSeconds = 0.0;
        }

        try
        {
            double sec = config.PronunciationEnglishInAdvance ?? config.PronunciationInAdvance ?? 0.0;
            if (double.IsNaN(sec) || double.IsInfinity(sec)) sec = 0.0;
            if (sec < 0) sec = 0.0;
            if (sec > 0.25) sec = 0.25;
            _pronunciationEnglishInAdvanceSeconds = sec;
        }
        catch
        {
            _pronunciationEnglishInAdvanceSeconds = _pronunciationInAdvanceSeconds;
        }

        try
        {
            double sec = config.PronunciationChineseInAdvance ?? config.PronunciationInAdvance ?? 0.0;
            if (double.IsNaN(sec) || double.IsInfinity(sec)) sec = 0.0;
            if (sec < 0) sec = 0.0;
            if (sec > 0.25) sec = 0.25;
            _pronunciationChineseInAdvanceSeconds = sec;
        }
        catch
        {
            _pronunciationChineseInAdvanceSeconds = _pronunciationInAdvanceSeconds;
        }

        _logEnabled = config.SapiReaderLog ?? true;

        Log($"ApplyConfig readKey={VkToDisplay(_hotkeyKey)} stopKey={VkToDisplay(_stopHotkeyKey)} autoRead={_customAutoReadHotKey}");
        Log($"ApplyConfig zhVoice={_chineseVoiceName} enVoice={_englishVoiceName} zhRate={_chineseVoiceRate} enRate={_englishVoiceRate}");
        Log($"ApplyConfig prebufferMixed={_prebufferMixedAudio} minLen={_prebufferMinTextLength} maxSeg={_prebufferMaxSegmentChars} keepAlive={_preventSpeakerSleep} advance={_pronunciationInAdvanceSeconds:0.###} enAdvance={_pronunciationEnglishInAdvanceSeconds:0.###} zhAdvance={_pronunciationChineseInAdvanceSeconds:0.###}");

        var oldStandby = Interlocked.Exchange(ref _standbyVoice, null);
        DisposeOldSynthAsync(oldStandby);
        EnsureStandbySynth();
    }

    private void SpeakText(string text)
    {
        CancelPendingStopReset();
        try { _voice.SpeakAsyncCancelAll(); } catch { }
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
        if (hasChinese && hasEnglish)
        {
            StartBufferedMixedPlayback(text);
            return;
        }
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

    private void StartBufferedMixedPlayback(string text)
    {
        CancelPendingStopReset();
        StopBufferedPlayback();
        try { _voice.SpeakAsyncCancelAll(); } catch { }

        var cts = new CancellationTokenSource();
        _bufferedPlaybackCts = cts;
        var token = cts.Token;

        try
        {
            BeginInvoke(new Action(() =>
            {
                _isSpeaking = true;
                _statusLabel.Text = "正在缓冲音频...";
                _statusLabel.ForeColor = Color.Blue;
            }));
        }
        catch { }

        Task.Run(() => BufferedPlaybackWorker(text, token), token);
    }

    private void BufferedPlaybackWorker(string text, CancellationToken token)
    {
        WaveOutPlayer? player = null;
        byte[]? carryTail = null;
        WaveFormat carryFmt = default;
        try
        {
            var segments = SmoothSegments(SplitByLanguage(text));
            segments = NormalizeMixedBoundaryWhitespace(segments);
            segments = SplitLongSegments(segments, _prebufferMaxSegmentChars);
            if (segments.Count == 0) return;

            const int prefetchDepth = 3;
            var tasks = new Queue<Task<PreparedAudio?>>();
            int seed = Math.Min(prefetchDepth, segments.Count);
            for (int s = 0; s < seed; s++)
            {
                var seg = segments[s];
                tasks.Enqueue(Task.Run(() => SynthesizeSegmentAudio(seg, token), token));
            }

            string? abortReason = null;
            for (int i = 0; i < segments.Count; i++)
            {
                if (token.IsCancellationRequested) return;
                PreparedAudio? prepared = null;
                try
                {
                    var task = tasks.Dequeue();
                    if (!task.Wait(12000))
                    {
                        abortReason = $"synth_timeout i={i} kind={segments[i].Kind}";
                        break;
                    }
                    prepared = task.GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    abortReason = $"dequeue_failed i={i} ex={ex.GetType().Name}";
                    break;
                }
                if (prepared == null)
                {
                    abortReason = $"synth_null i={i} kind={segments[i].Kind}";
                    break;
                }

                int nextIndex = i + seed;
                if (nextIndex < segments.Count)
                {
                    var seg = segments[nextIndex];
                    tasks.Enqueue(Task.Run(() => SynthesizeSegmentAudio(seg, token), token));
                }

                if (token.IsCancellationRequested) return;
                if (!TryExtractWav(prepared.WavStream, out var fmt, out var pcm))
                {
                    abortReason = $"extract_failed i={i} kind={segments[i].Kind}";
                    break;
                }
                if (pcm.Length == 0) continue;

                player ??= new WaveOutPlayer(fmt);
                _bufferedWavePlayer = player;
                if (!player.IsSameFormat(fmt))
                {
                    abortReason = $"format_mismatch i={i}";
                    break;
                }

                if (carryTail != null)
                {
                    if (carryFmt.IsValid && carryFmt.IsPcm16 && fmt.IsPcm16 && player.IsSameFormat(carryFmt))
                    {
                        int m = Math.Min(carryTail.Length, pcm.Length);
                        m = (m / carryFmt.BlockAlign) * carryFmt.BlockAlign;
                        if (m > 0)
                        {
                            MixPcm16TailWithHead(carryTail, pcm, m, carryFmt);
                            player.Enqueue(carryTail);
                            var swBack = Stopwatch.StartNew();
                            while (!token.IsCancellationRequested && player.InFlight > 4)
                            {
                                Thread.Sleep(10);
                                if (swBack.ElapsedMilliseconds > 3000) break;
                            }

                            var remain = new byte[pcm.Length - m];
                            Buffer.BlockCopy(pcm, m, remain, 0, remain.Length);
                            pcm = remain;
                        }
                        else
                        {
                            player.Enqueue(carryTail);
                            var swBack = Stopwatch.StartNew();
                            while (!token.IsCancellationRequested && player.InFlight > 4)
                            {
                                Thread.Sleep(10);
                                if (swBack.ElapsedMilliseconds > 3000) break;
                            }
                        }
                    }
                    else
                    {
                        player.Enqueue(carryTail);
                        var swBack = Stopwatch.StartNew();
                        while (!token.IsCancellationRequested && player.InFlight > 4)
                        {
                            Thread.Sleep(10);
                            if (swBack.ElapsedMilliseconds > 3000) break;
                        }
                    }

                    carryTail = null;
                }

                if (pcm.Length == 0) continue;

                int holdBackBytes = 0;
                if (i + 1 < segments.Count && segments[i].Kind != segments[i + 1].Kind && fmt.IsPcm16)
                {
                    double sec = segments[i + 1].Kind == TextKind.English ? _pronunciationEnglishInAdvanceSeconds : _pronunciationChineseInAdvanceSeconds;
                    long frames = (long)(fmt.SamplesPerSec * sec);
                    if (frames > 0)
                    {
                        long bytes = frames * fmt.BlockAlign;
                        if (bytes > int.MaxValue) bytes = int.MaxValue;
                        int b = (int)bytes;
                        b = (b / fmt.BlockAlign) * fmt.BlockAlign;
                        if (b > 0) holdBackBytes = b;
                    }
                }

                if (holdBackBytes > 0)
                {
                    int m = Math.Min(holdBackBytes, pcm.Length);
                    m = (m / fmt.BlockAlign) * fmt.BlockAlign;
                    int mainLen = pcm.Length - m;

                    if (mainLen > 0)
                    {
                        var main = new byte[mainLen];
                        Buffer.BlockCopy(pcm, 0, main, 0, mainLen);
                        player.Enqueue(main);
                        var swBack = Stopwatch.StartNew();
                        while (!token.IsCancellationRequested && player.InFlight > 4)
                        {
                            Thread.Sleep(10);
                            if (swBack.ElapsedMilliseconds > 3000) break;
                        }
                    }

                    if (m > 0)
                    {
                        carryTail = new byte[m];
                        Buffer.BlockCopy(pcm, mainLen, carryTail, 0, m);
                        carryFmt = fmt;
                    }
                }
                else
                {
                    player.Enqueue(pcm);
                    var swBack = Stopwatch.StartNew();
                    while (!token.IsCancellationRequested && player.InFlight > 4)
                    {
                        Thread.Sleep(10);
                        if (swBack.ElapsedMilliseconds > 3000) break;
                    }
                }
            }

            if (!token.IsCancellationRequested && player != null && carryTail != null && carryTail.Length > 0)
            {
                Log($"BufferedPlaybackWorker flush_carry bytes={carryTail.Length}");
                player.Enqueue(carryTail);
            }

            if (player != null)
            {
                var swDrain = Stopwatch.StartNew();
                while (!token.IsCancellationRequested && player.InFlight > 0)
                {
                    Thread.Sleep(10);
                    if (swDrain.ElapsedMilliseconds > 15000)
                    {
                        Log("BufferedPlaybackWorker drain_timeout");
                        try { player.Stop(); } catch { }
                        break;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(abortReason))
            {
                Log($"BufferedPlaybackWorker abort {abortReason}");
            }
        }
        finally
        {
            try { player?.Dispose(); } catch { }
            if (ReferenceEquals(_bufferedWavePlayer, player)) _bufferedWavePlayer = null;
            try
            {
                BeginInvoke(new Action(() =>
                {
                    _isSpeaking = false;
                    if (!_isAutoReading) UpdateStatusLabel();
                }));
            }
            catch { }
        }
    }

    private static void MixPcm16TailWithHead(byte[] tail, byte[] head, int mixBytes, WaveFormat fmt)
    {
        if (!fmt.IsPcm16) return;
        int frameSize = fmt.BlockAlign;
        if (frameSize <= 0) return;

        int tailStart = tail.Length - mixBytes;
        if (tailStart < 0) tailStart = 0;
        int len = Math.Min(mixBytes, Math.Min(tail.Length - tailStart, head.Length));
        len = (len / frameSize) * frameSize;
        if (len <= 0) return;

        for (int i = 0; i < len; i += 2)
        {
            int t = tailStart + i;
            short a = (short)(tail[t] | (tail[t + 1] << 8));
            short b = (short)(head[i] | (head[i + 1] << 8));
            int s = (a + b) / 2;
            tail[t] = (byte)(s & 0xFF);
            tail[t + 1] = (byte)((s >> 8) & 0xFF);
        }
    }

    private sealed class PreparedAudio : IDisposable
    {
        public PreparedAudio(MemoryStream wavStream)
        {
            WavStream = wavStream;
        }

        public MemoryStream WavStream { get; }

        public void Dispose()
        {
            try { WavStream.Dispose(); } catch { }
        }
    }

    private PreparedAudio? SynthesizeSegmentAudio(TextSegment segment, CancellationToken token)
    {
        if (token.IsCancellationRequested) return null;

        var voiceName = segment.Kind == TextKind.Chinese ? _chineseVoiceName : _englishVoiceName;
        var rate = segment.Kind == TextKind.Chinese ? _chineseVoiceRate : _englishVoiceRate;

        try
        {
            var ms = new MemoryStream();
            using var synth = new SpeechSynthesizer();
            var outFmt = new SpeechAudioFormatInfo(48000, AudioBitsPerSample.Sixteen, AudioChannel.Mono);
            synth.SetOutputToAudioStream(ms, outFmt);
            synth.Volume = _voice.Volume;
            synth.Rate = ClampRate(rate);
            if (!string.IsNullOrWhiteSpace(voiceName))
            {
                try { synth.SelectVoice(voiceName); }
                catch
                {
                    Log($"SynthesizeSegmentAudio SelectVoice failed voice={voiceName}");
                    return null;
                }
            }

            if (token.IsCancellationRequested) return null;
            synth.Speak(segment.Text);
            ms.Position = 0;
            return new PreparedAudio(ms);
        }
        catch
        {
            return null;
        }
    }

    private PromptBuilder? BuildMixedVoicePrompt(string text)
    {
        try
        {
            var segments = SplitByLanguage(text);
            segments = SmoothSegments(segments);
            segments = NormalizeMixedBoundaryWhitespace(segments);
            if (segments.Count == 0) return null;

            var prompt = new PromptBuilder();
            string activeVoice = "";
            int activeRate = 0;
            bool voiceOpen = false;
            bool styleOpen = false;

            foreach (var segment in segments)
            {
                var voiceName = segment.Kind == TextKind.Chinese ? _chineseVoiceName : _englishVoiceName;
                var rate = segment.Kind == TextKind.Chinese ? _chineseVoiceRate : _englishVoiceRate;
                if (string.IsNullOrWhiteSpace(voiceName))
                {
                    if (voiceOpen)
                    {
                        try { if (styleOpen) prompt.EndStyle(); } catch { }
                        try { prompt.EndVoice(); } catch { }
                        voiceOpen = false;
                        styleOpen = false;
                        activeVoice = "";
                        activeRate = 0;
                    }
                    AppendSegmentText(prompt, segment);
                    continue;
                }

                try
                {
                    if (!voiceOpen || !string.Equals(activeVoice, voiceName, StringComparison.OrdinalIgnoreCase) || activeRate != rate)
                    {
                        if (voiceOpen)
                        {
                            try { if (styleOpen) prompt.EndStyle(); } catch { }
                            try { prompt.EndVoice(); } catch { }
                            voiceOpen = false;
                            styleOpen = false;
                        }

                        prompt.StartVoice(voiceName);
                        voiceOpen = true;
                        activeVoice = voiceName;
                        activeRate = rate;

                        var style = new PromptStyle { Rate = MapPromptRate(rate) };
                        prompt.StartStyle(style);
                        styleOpen = true;
                    }
                    AppendSegmentText(prompt, segment);
                }
                catch
                {
                    if (voiceOpen)
                    {
                        try { if (styleOpen) prompt.EndStyle(); } catch { }
                        try { prompt.EndVoice(); } catch { }
                        voiceOpen = false;
                        styleOpen = false;
                        activeVoice = "";
                        activeRate = 0;
                    }
                    AppendSegmentText(prompt, segment);
                }
            }

            if (voiceOpen)
            {
                try { if (styleOpen) prompt.EndStyle(); } catch { }
                try { prompt.EndVoice(); } catch { }
            }
            return prompt;
        }
        catch
        {
            return null;
        }
    }

    private static bool TryExtractWav(MemoryStream wavStream, out WaveFormat fmt, out byte[] pcm)
    {
        fmt = default;
        pcm = Array.Empty<byte>();
        try
        {
            var data = wavStream.ToArray();
            if (data.Length == 0) return false;

            if (data.Length < 12 || data[0] != (byte)'R' || data[1] != (byte)'I' || data[2] != (byte)'F' || data[3] != (byte)'F')
            {
                const ushort rawFormatTag = 1;
                const ushort rawChannels = 1;
                const uint rawSamplesPerSec = 48000;
                const ushort rawBitsPerSample = 16;
                ushort rawBlockAlign = (ushort)(rawChannels * (rawBitsPerSample / 8));
                uint rawAvgBytesPerSec = rawSamplesPerSec * rawBlockAlign;
                fmt = new WaveFormat(rawFormatTag, rawChannels, rawSamplesPerSec, rawAvgBytesPerSec, rawBlockAlign, rawBitsPerSample);
                pcm = data;
                return pcm.Length > 0 && fmt.IsValid;
            }
            if (data.Length < 44) return false;
            int offset = 0;
            if (data[offset++] != (byte)'R' || data[offset++] != (byte)'I' || data[offset++] != (byte)'F' || data[offset++] != (byte)'F') return false;
            offset += 4;
            if (data[offset++] != (byte)'W' || data[offset++] != (byte)'A' || data[offset++] != (byte)'V' || data[offset++] != (byte)'E') return false;

            ushort wFormatTag = 0;
            ushort nChannels = 0;
            uint nSamplesPerSec = 0;
            uint nAvgBytesPerSec = 0;
            ushort nBlockAlign = 0;
            ushort wBitsPerSample = 0;
            int fmtFound = 0;
            int dataFound = 0;

            while (offset + 8 <= data.Length)
            {
                uint id = BitConverter.ToUInt32(data, offset);
                offset += 4;
                int size = BitConverter.ToInt32(data, offset);
                offset += 4;
                if (size < 0 || offset + size > data.Length) break;

                if (id == 0x20746D66)
                {
                    if (size < 16) return false;
                    wFormatTag = BitConverter.ToUInt16(data, offset + 0);
                    nChannels = BitConverter.ToUInt16(data, offset + 2);
                    nSamplesPerSec = BitConverter.ToUInt32(data, offset + 4);
                    nAvgBytesPerSec = BitConverter.ToUInt32(data, offset + 8);
                    nBlockAlign = BitConverter.ToUInt16(data, offset + 12);
                    wBitsPerSample = BitConverter.ToUInt16(data, offset + 14);
                    fmtFound = 1;
                }
                else if (id == 0x61746164)
                {
                    pcm = new byte[size];
                    Buffer.BlockCopy(data, offset, pcm, 0, size);
                    dataFound = 1;
                }

                offset += size;
                if ((size & 1) == 1) offset += 1;
                if (fmtFound == 1 && dataFound == 1) break;
            }

            if (fmtFound == 0 || dataFound == 0) return false;
            fmt = new WaveFormat(wFormatTag, nChannels, nSamplesPerSec, nAvgBytesPerSec, nBlockAlign, wBitsPerSample);
            return pcm.Length > 0 && fmt.IsValid;
        }
        catch
        {
            fmt = default;
            pcm = Array.Empty<byte>();
            return false;
        }
    }

    private static byte[] TrimSilence(byte[] pcm, WaveFormat fmt, int leadingPadMs, int trailingPadMs)
    {
        try
        {
            if (pcm.Length == 0) return pcm;
            if (!fmt.IsPcm16) return pcm;
            int frameSize = fmt.BlockAlign;
            if (frameSize <= 0) return pcm;
            int totalFrames = pcm.Length / frameSize;
            if (totalFrames <= 0) return Array.Empty<byte>();

            int threshold = 80;
            int first = 0;
            int last = totalFrames - 1;

            bool FrameHasSignal(int frameIndex)
            {
                int baseOff = frameIndex * frameSize;
                for (int ch = 0; ch < fmt.Channels; ch++)
                {
                    int off = baseOff + ch * 2;
                    short s = (short)(pcm[off] | (pcm[off + 1] << 8));
                    if (s < 0) s = (short)-s;
                    if (s > threshold) return true;
                }
                return false;
            }

            while (first < totalFrames && !FrameHasSignal(first)) first++;
            while (last >= first && !FrameHasSignal(last)) last--;
            if (last < first) return Array.Empty<byte>();

            int padLead = (int)((long)fmt.SamplesPerSec * Math.Max(0, leadingPadMs) / 1000L);
            int padTrail = (int)((long)fmt.SamplesPerSec * Math.Max(0, trailingPadMs) / 1000L);
            int start = Math.Max(0, first - padLead);
            int end = Math.Min(totalFrames - 1, last + padTrail);
            int frames = end - start + 1;
            if (frames <= 0) return Array.Empty<byte>();

            int bytes = frames * frameSize;
            var outBytes = new byte[bytes];
            Buffer.BlockCopy(pcm, start * frameSize, outBytes, 0, bytes);
            return outBytes;
        }
        catch
        {
            return pcm;
        }
    }

    private readonly struct WaveFormat
    {
        public WaveFormat(ushort formatTag, ushort channels, uint samplesPerSec, uint avgBytesPerSec, ushort blockAlign, ushort bitsPerSample)
        {
            FormatTag = formatTag;
            Channels = channels;
            SamplesPerSec = samplesPerSec;
            AvgBytesPerSec = avgBytesPerSec;
            BlockAlign = blockAlign;
            BitsPerSample = bitsPerSample;
        }

        public ushort FormatTag { get; }
        public ushort Channels { get; }
        public uint SamplesPerSec { get; }
        public uint AvgBytesPerSec { get; }
        public ushort BlockAlign { get; }
        public ushort BitsPerSample { get; }

        public bool IsValid => Channels > 0 && SamplesPerSec > 0 && BlockAlign > 0 && BitsPerSample > 0;
        public bool IsPcm16 => FormatTag == 1 && BitsPerSample == 16 && BlockAlign == Channels * 2;
    }

    private sealed class WaveOutPlayer : IDisposable
    {
        private const int WAVE_MAPPER = -1;
        private const int CALLBACK_EVENT = 0x00050000;
        private const uint WHDR_DONE = 0x00000001;

        [StructLayout(LayoutKind.Sequential)]
        private struct WAVEFORMATEX
        {
            public ushort wFormatTag;
            public ushort nChannels;
            public uint nSamplesPerSec;
            public uint nAvgBytesPerSec;
            public ushort nBlockAlign;
            public ushort wBitsPerSample;
            public ushort cbSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WAVEHDR
        {
            public IntPtr lpData;
            public uint dwBufferLength;
            public uint dwBytesRecorded;
            public IntPtr dwUser;
            public uint dwFlags;
            public uint dwLoops;
            public IntPtr lpNext;
            public IntPtr reserved;
        }

        [DllImport("winmm.dll")]
        private static extern int waveOutOpen(out IntPtr hWaveOut, int uDeviceID, ref WAVEFORMATEX lpFormat, IntPtr dwCallback, IntPtr dwInstance, int dwFlags);

        [DllImport("winmm.dll")]
        private static extern int waveOutPrepareHeader(IntPtr hWaveOut, IntPtr lpWaveOutHdr, int uSize);

        [DllImport("winmm.dll")]
        private static extern int waveOutUnprepareHeader(IntPtr hWaveOut, IntPtr lpWaveOutHdr, int uSize);

        [DllImport("winmm.dll")]
        private static extern int waveOutWrite(IntPtr hWaveOut, IntPtr lpWaveOutHdr, int uSize);

        [DllImport("winmm.dll")]
        private static extern int waveOutReset(IntPtr hWaveOut);

        [DllImport("winmm.dll")]
        private static extern int waveOutClose(IntPtr hWaveOut);

        private sealed class BufferState
        {
            public IntPtr HeaderPtr;
            public IntPtr DataPtr;
        }

        private readonly object _gate = new();
        private IntPtr _hwo;
        private WaveFormat _fmt;
        private int _disposed;
        private int _accepting = 1;
        private readonly AutoResetEvent _doneEvent = new(false);
        private readonly ManualResetEvent _stopEvent = new(false);
        private readonly Thread _reapThread;
        private readonly List<BufferState> _buffers = new();

        public WaveOutPlayer(WaveFormat fmt)
        {
            _fmt = fmt;
            var wf = new WAVEFORMATEX
            {
                wFormatTag = fmt.FormatTag,
                nChannels = fmt.Channels,
                nSamplesPerSec = fmt.SamplesPerSec,
                nAvgBytesPerSec = fmt.AvgBytesPerSec,
                nBlockAlign = fmt.BlockAlign,
                wBitsPerSample = fmt.BitsPerSample,
                cbSize = 0
            };
            var rc = waveOutOpen(out _hwo, WAVE_MAPPER, ref wf, _doneEvent.SafeWaitHandle.DangerousGetHandle(), IntPtr.Zero, CALLBACK_EVENT);
            if (rc != 0) _hwo = IntPtr.Zero;

            _reapThread = new Thread(ReapLoop)
            {
                IsBackground = true,
                Name = "WaveOutPlayer.ReapLoop"
            };
            _reapThread.Start();
        }

        public int InFlight => Volatile.Read(ref _inFlight);
        private int _inFlight;

        public bool IsSameFormat(WaveFormat fmt)
        {
            return fmt.FormatTag == _fmt.FormatTag
                && fmt.Channels == _fmt.Channels
                && fmt.SamplesPerSec == _fmt.SamplesPerSec
                && fmt.AvgBytesPerSec == _fmt.AvgBytesPerSec
                && fmt.BlockAlign == _fmt.BlockAlign
                && fmt.BitsPerSample == _fmt.BitsPerSample;
        }

        public void Enqueue(byte[] pcm)
        {
            if (pcm.Length == 0) return;
            if (Volatile.Read(ref _disposed) == 1) return;
            if (Volatile.Read(ref _accepting) == 0) return;

            lock (_gate)
            {
                if (_hwo == IntPtr.Zero) return;
                if (Volatile.Read(ref _accepting) == 0) return;

                var dataPtr = Marshal.AllocHGlobal(pcm.Length);
                Marshal.Copy(pcm, 0, dataPtr, pcm.Length);

                var state = new BufferState();
                state.DataPtr = dataPtr;

                var hdr = new WAVEHDR
                {
                    lpData = dataPtr,
                    dwBufferLength = (uint)pcm.Length,
                    dwBytesRecorded = 0,
                    dwUser = IntPtr.Zero,
                    dwFlags = 0,
                    dwLoops = 0,
                    lpNext = IntPtr.Zero,
                    reserved = IntPtr.Zero
                };

                int hdrSize = Marshal.SizeOf<WAVEHDR>();
                var hdrPtr = Marshal.AllocHGlobal(hdrSize);
                state.HeaderPtr = hdrPtr;
                Marshal.StructureToPtr(hdr, hdrPtr, false);

                var rcPrep = waveOutPrepareHeader(_hwo, hdrPtr, hdrSize);
                if (rcPrep != 0)
                {
                    try { Marshal.FreeHGlobal(hdrPtr); } catch { }
                    try { Marshal.FreeHGlobal(dataPtr); } catch { }
                    return;
                }

                Interlocked.Increment(ref _inFlight);
                var rcWrite = waveOutWrite(_hwo, hdrPtr, hdrSize);
                if (rcWrite != 0)
                {
                    Interlocked.Decrement(ref _inFlight);
                    try { waveOutUnprepareHeader(_hwo, hdrPtr, hdrSize); } catch { }
                    try { Marshal.FreeHGlobal(hdrPtr); } catch { }
                    try { Marshal.FreeHGlobal(dataPtr); } catch { }
                    return;
                }

                state.HeaderPtr = hdrPtr;
                _buffers.Add(state);
            }
        }

        public void Stop()
        {
            Volatile.Write(ref _accepting, 0);
            IntPtr h;
            lock (_gate)
            {
                h = _hwo;
                if (h != IntPtr.Zero)
                {
                    try { waveOutReset(h); } catch { }
                }
            }
        }

        private void ReapLoop()
        {
            try
            {
                var waits = new WaitHandle[] { _doneEvent, _stopEvent };
                while (true)
                {
                    int signaled = WaitHandle.WaitAny(waits);
                    if (signaled == 1) return;
                    DrainDoneBuffers();
                }
            }
            catch
            {
            }
        }

        private void DrainDoneBuffers()
        {
            List<BufferState>? done = null;
            IntPtr h;
            lock (_gate)
            {
                h = _hwo;
                if (h == IntPtr.Zero || _buffers.Count == 0) return;

                for (int i = _buffers.Count - 1; i >= 0; i--)
                {
                    var st = _buffers[i];
                    if (st.HeaderPtr == IntPtr.Zero) continue;
                    WAVEHDR hdr;
                    try { hdr = Marshal.PtrToStructure<WAVEHDR>(st.HeaderPtr); }
                    catch { continue; }
                    if ((hdr.dwFlags & WHDR_DONE) == 0) continue;

                    done ??= new List<BufferState>();
                    done.Add(st);
                    _buffers.RemoveAt(i);
                }
            }

            if (done == null) return;
            int hdrSize = Marshal.SizeOf<WAVEHDR>();
            foreach (var st in done)
            {
                try { if (h != IntPtr.Zero) waveOutUnprepareHeader(h, st.HeaderPtr, hdrSize); } catch { }
                try { if (st.HeaderPtr != IntPtr.Zero) Marshal.FreeHGlobal(st.HeaderPtr); } catch { }
                try { if (st.DataPtr != IntPtr.Zero) Marshal.FreeHGlobal(st.DataPtr); } catch { }
                int v = Interlocked.Decrement(ref _inFlight);
                if (v < 0) Interlocked.Exchange(ref _inFlight, 0);
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 1) return;

            Stop();
            try { _stopEvent.Set(); } catch { }
            try { _doneEvent.Set(); } catch { }
            try { if (_reapThread.IsAlive) _reapThread.Join(300); } catch { }

            try { DrainDoneBuffers(); } catch { }

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 800 && Volatile.Read(ref _inFlight) > 0)
            {
                Thread.Sleep(10);
            }

            IntPtr h;
            lock (_gate)
            {
                h = _hwo;
                _hwo = IntPtr.Zero;
                _buffers.Clear();
            }

            if (h != IntPtr.Zero)
            {
                try { waveOutClose(h); } catch { }
            }

            try { _doneEvent.Dispose(); } catch { }
            try { _stopEvent.Dispose(); } catch { }
        }
    }

    private static List<TextSegment> NormalizeMixedBoundaryWhitespace(List<TextSegment> segments)
    {
        if (segments.Count <= 1) return segments;

        static bool IsWs(char ch)
        {
            return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' || ch == '\u3000';
        }

        static string TrimEndWs(string s)
        {
            int end = s.Length - 1;
            while (end >= 0 && IsWs(s[end])) end--;
            if (end == s.Length - 1) return s;
            return end < 0 ? "" : s.Substring(0, end + 1);
        }

        static string TrimStartWs(string s)
        {
            int start = 0;
            while (start < s.Length && IsWs(s[start])) start++;
            if (start == 0) return s;
            return start >= s.Length ? "" : s.Substring(start);
        }

        var list = segments;
        for (int i = 0; i + 1 < list.Count; i++)
        {
            var a = list[i];
            var b = list[i + 1];
            if (a.Kind == b.Kind) continue;

            var at = TrimEndWs(a.Text);
            var bt = TrimStartWs(b.Text);
            if (!ReferenceEquals(at, a.Text) || !ReferenceEquals(bt, b.Text))
            {
                if (ReferenceEquals(list, segments)) list = new List<TextSegment>(segments);
                list[i] = new TextSegment(a.Kind, at);
                list[i + 1] = new TextSegment(b.Kind, bt);
            }
        }
        return list;
    }

    private void AppendSegmentText(PromptBuilder prompt, TextSegment segment)
    {
        if (!_mixedGranularOutput)
        {
            prompt.AppendText(segment.Text);
            return;
        }

        if (segment.Kind == TextKind.Chinese)
        {
            AppendChineseCharByChar(prompt, segment.Text);
            return;
        }

        AppendEnglishWordByWord(prompt, segment.Text);
    }

    private static void AppendChineseCharByChar(PromptBuilder prompt, string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        var neutral = new StringBuilder();
        foreach (var ch in text)
        {
            if (IsChineseChar(ch))
            {
                if (neutral.Length > 0)
                {
                    prompt.AppendText(neutral.ToString());
                    neutral.Clear();
                }
                prompt.AppendText(ch.ToString());
            }
            else
            {
                neutral.Append(ch);
            }
        }

        if (neutral.Length > 0)
        {
            prompt.AppendText(neutral.ToString());
        }
    }

    private static void AppendEnglishWordByWord(PromptBuilder prompt, string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        var word = new StringBuilder();
        var neutral = new StringBuilder();

        static bool IsWordChar(char ch)
        {
            return IsLatinLetter(ch) || char.IsDigit(ch) || ch == '\'' || ch == '-';
        }

        foreach (var ch in text)
        {
            if (IsWordChar(ch))
            {
                if (neutral.Length > 0)
                {
                    prompt.AppendText(neutral.ToString());
                    neutral.Clear();
                }
                word.Append(ch);
                continue;
            }

            if (word.Length > 0)
            {
                prompt.AppendText(word.ToString());
                word.Clear();
            }
            neutral.Append(ch);
        }

        if (word.Length > 0)
        {
            prompt.AppendText(word.ToString());
            word.Clear();
        }
        if (neutral.Length > 0)
        {
            prompt.AppendText(neutral.ToString());
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

    private static string VkToDisplay(uint vk)
    {
        if (vk >= 0x70 && vk <= 0x87)
        {
            return $"F{vk - 0x70 + 1}";
        }
        if ((vk >= 0x30 && vk <= 0x39) || (vk >= 0x41 && vk <= 0x5A))
        {
            return ((char)vk).ToString();
        }
        return $"VK_{vk}";
    }

    private bool TryParseNoModifierHotkey(object? value, out uint vk)
    {
        vk = 0;
        if (value == null) return false;

        try
        {
            if (value is JsonElement element)
            {
                if (element.ValueKind == JsonValueKind.String)
                {
                    var s = element.GetString();
                    return TryParseNoModifierHotkeyString(s, out vk);
                }
                if (element.ValueKind == JsonValueKind.Number)
                {
                    int idx = element.GetInt32();
                    return TryParseLegacyFunctionKeyIndex(idx, out vk);
                }
                return false;
            }

            if (value is string str)
            {
                return TryParseNoModifierHotkeyString(str, out vk);
            }

            if (value is int i)
            {
                return TryParseLegacyFunctionKeyIndex(i, out vk);
            }

            int converted = Convert.ToInt32(value);
            return TryParseLegacyFunctionKeyIndex(converted, out vk);
        }
        catch
        {
            return false;
        }
    }

    private static bool TryParseLegacyFunctionKeyIndex(int idx, out uint vk)
    {
        vk = 0;
        if (idx < 0 || idx >= 12) return false;
        vk = (uint)(0x70 + idx);
        return true;
    }

    private static bool TryParseNoModifierHotkeyString(string? s, out uint vk)
    {
        vk = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;
        var t = s.Trim();

        if (t.StartsWith("f", StringComparison.OrdinalIgnoreCase) && int.TryParse(t.Substring(1), out int fNum))
        {
            if (fNum >= 1 && fNum <= 24)
            {
                vk = (uint)(0x70 + fNum - 1);
                return true;
            }
            return false;
        }

        if (t.Length == 1)
        {
            char ch = char.ToUpperInvariant(t[0]);
            if ((ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9'))
            {
                vk = ch;
                return true;
            }
        }

        return false;
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

    private List<TextSegment> SplitLongSegments(List<TextSegment> segments, int maxChars)
    {
        if (segments.Count == 0) return segments;
        if (maxChars <= 0) return segments;

        var result = new List<TextSegment>();
        foreach (var seg in segments)
        {
            if (seg.Text.Length <= maxChars)
            {
                result.Add(seg);
                continue;
            }

            int start = 0;
            while (start < seg.Text.Length)
            {
                int len = Math.Min(maxChars, seg.Text.Length - start);
                result.Add(new TextSegment(seg.Kind, seg.Text.Substring(start, len)));
                start += len;
            }
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
    public object HotkeyIndex { get; set; } = 0; // 支持 string (例如 "F1") 或 int (0-11)
    public object StopHotkeyIndex { get; set; } = 1; // 支持 string (例如 "F3") 或 int (0-11)
    public object AutoReadHotkeyIndex { get; set; } = 0; // 支持 string (自定义) 或 int (F1-F12 索引)
    [JsonPropertyName("sapi-reader.log")]
    public bool? SapiReaderLog { get; set; }
    public string Version { get; set; } = "";
    public string VoiceName { get; set; } = "";
    public string ChineseVoiceName { get; set; } = "";
    public string EnglishVoiceName { get; set; } = "";
    public bool AutoSwitchVoice { get; set; } = true;
    public int Rate { get; set; }
    public int? ChineseVoiceRate { get; set; }
    public int? EnglishVoiceRate { get; set; }
    public int? MixedChineseMinChars { get; set; }
    public int? MixedEnglishMinLetters { get; set; }
    public bool? MixedGranularOutput { get; set; }
    public bool? WarmUpVoices { get; set; }
    [JsonPropertyName("Prevent the speaker from going into sleep mode")]
    public bool? PreventSpeakerFromGoingIntoSleepMode { get; set; }
    public bool? PrebufferMixedAudio { get; set; }
    public int? PrebufferMinTextLength { get; set; }
    public int? PrebufferMaxSegmentChars { get; set; }
    [JsonPropertyName("Pronunciation in advance")]
    public double? PronunciationInAdvance { get; set; }
    [JsonPropertyName("Pronunciation English in advance")]
    public double? PronunciationEnglishInAdvance { get; set; }
    [JsonPropertyName("Pronunciation Chinese in advance")]
    public double? PronunciationChineseInAdvance { get; set; }
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
