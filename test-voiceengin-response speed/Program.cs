using System.Diagnostics;
using System.Speech.Synthesis;

namespace VoiceEngineResponseSpeed;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        var options = Options.Parse(args);
        if (options.ShowHelp)
        {
            Console.WriteLine(Options.HelpText);
            return 0;
        }

        Console.WriteLine("Probing voices...");
        Console.Out.Flush();

        var installedResult = RunStaWithTimeout(
            () =>
            {
                using var probe = new SpeechSynthesizer();
                return probe.GetInstalledVoices().Where(v => v.Enabled).ToList();
            },
            options.ProbeTimeoutMs);

        if (installedResult.TimedOut)
        {
            Console.WriteLine($"Probe timeout after {options.ProbeTimeoutMs}ms.");
            return 4;
        }

        if (installedResult.Error != null)
        {
            Console.WriteLine($"Probe error: {installedResult.Error.GetType().Name}: {installedResult.Error.Message}");
            return 5;
        }

        var installed = installedResult.Value ?? new List<InstalledVoice>();

        if (installed.Count == 0)
        {
            Console.WriteLine("No enabled SAPI voices found.");
            return 2;
        }

        var selected = installed
            .Select(v => v.VoiceInfo.Name)
            .Where(name => options.VoiceContains == null || name.Contains(options.VoiceContains, StringComparison.OrdinalIgnoreCase))
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (selected.Count == 0)
        {
            Console.WriteLine("No voices matched filter.");
            return 3;
        }

        Console.WriteLine($"Voices: {selected.Count}");
        Console.WriteLine($"Iterations: {options.Iterations}");
        Console.WriteLine($"Timeout: {options.TimeoutMs}ms");
        Console.WriteLine($"Output: {(options.OutputNull ? "null" : "default")}");
        Console.WriteLine($"TextLength: {options.Text.Length}");
        Console.WriteLine();

        var rows = new List<ResultRow>();
        foreach (var voiceName in selected)
        {
            var rowResult = RunStaWithTimeout(() => BenchmarkVoice(voiceName, options), options.VoiceTimeoutMs);
            ResultRow row;
            if (rowResult.TimedOut)
            {
                row = new ResultRow(voiceName, new Stats(0, 0, 0, 0), new Stats(0, 0, 0, 0), Failures: options.Iterations, Total: options.Iterations);
            }
            else if (rowResult.Error != null)
            {
                row = new ResultRow(voiceName, new Stats(0, 0, 0, 0), new Stats(0, 0, 0, 0), Failures: options.Iterations, Total: options.Iterations);
            }
            else
            {
                row = rowResult.Value ?? new ResultRow(voiceName, new Stats(0, 0, 0, 0), new Stats(0, 0, 0, 0), Failures: options.Iterations, Total: options.Iterations);
            }

            rows.Add(row);
            PrintRow(row);
        }

        Console.WriteLine();
        PrintSummary(rows);
        return 0;
    }

    private static StaRunResult<T> RunStaWithTimeout<T>(Func<T> func, int timeoutMs)
    {
        var done = new ManualResetEventSlim(false);
        Exception? error = null;
        T? value = default;

        var thread = new Thread(() =>
        {
            try
            {
                value = func();
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                done.Set();
            }
        })
        {
            IsBackground = true
        };

        try { thread.SetApartmentState(ApartmentState.STA); } catch { }
        thread.Start();

        bool ok = done.Wait(timeoutMs);
        if (!ok) return new StaRunResult<T>(TimedOut: true, Value: default, Error: null);
        return new StaRunResult<T>(TimedOut: false, Value: value, Error: error);
    }

    private static ResultRow BenchmarkVoice(string voiceName, Options options)
    {
        var startLatencies = new List<double>(capacity: options.Iterations);
        var doneLatencies = new List<double>(capacity: options.Iterations);
        int failures = 0;

        using var synth = new SpeechSynthesizer();
        if (options.OutputNull) synth.SetOutputToNull();
        else synth.SetOutputToDefaultAudioDevice();

        try { synth.SelectVoice(voiceName); } catch { }
        try { synth.Rate = options.Rate; } catch { }
        try { synth.Volume = options.Volume; } catch { }

        if (options.WarmUp)
        {
            try { synth.Speak(" "); } catch { }
        }

        for (int i = 0; i < options.Iterations; i++)
        {
            var started = new ManualResetEventSlim(false);
            var completed = new ManualResetEventSlim(false);
            var sw = Stopwatch.StartNew();
            double startedMs = -1;
            double completedMs = -1;

            EventHandler<SpeakStartedEventArgs>? startedHandler = null;
            EventHandler<SpeakCompletedEventArgs>? completedHandler = null;

            startedHandler = (_, _) =>
            {
                startedMs = sw.Elapsed.TotalMilliseconds;
                started.Set();
            };

            completedHandler = (_, _) =>
            {
                completedMs = sw.Elapsed.TotalMilliseconds;
                completed.Set();
            };

            synth.SpeakStarted += startedHandler;
            synth.SpeakCompleted += completedHandler;

            try
            {
                synth.SpeakAsync(options.Text);

                bool okStart = started.Wait(options.TimeoutMs);
                bool okDone = completed.Wait(options.TimeoutMs);
                if (!okStart || !okDone || startedMs < 0 || completedMs < 0)
                {
                    failures++;
                }
                else
                {
                    startLatencies.Add(startedMs);
                    doneLatencies.Add(completedMs);
                }
            }
            catch
            {
                failures++;
            }
            finally
            {
                try { synth.SpeakAsyncCancelAll(); } catch { }
                synth.SpeakStarted -= startedHandler;
                synth.SpeakCompleted -= completedHandler;
                started.Dispose();
                completed.Dispose();
            }
        }

        return new ResultRow(
            VoiceName: voiceName,
            StartMs: Stats.From(startLatencies),
            DoneMs: Stats.From(doneLatencies),
            Failures: failures,
            Total: options.Iterations);
    }

    private static void PrintRow(ResultRow row)
    {
        Console.WriteLine(row.VoiceName);
        Console.WriteLine($"  start_ms  avg={row.StartMs.Avg:0.0}  p95={row.StartMs.P95:0.0}  min={row.StartMs.Min:0.0}  n={row.StartMs.Count}");
        Console.WriteLine($"  done_ms   avg={row.DoneMs.Avg:0.0}  p95={row.DoneMs.P95:0.0}  min={row.DoneMs.Min:0.0}  n={row.DoneMs.Count}");
        Console.WriteLine($"  failures  {row.Failures}/{row.Total}");
        Console.WriteLine();
    }

    private static void PrintSummary(List<ResultRow> rows)
    {
        Console.WriteLine("Top by start_ms avg:");
        foreach (var row in rows.OrderBy(r => r.StartMs.Avg).ThenBy(r => r.VoiceName, StringComparer.OrdinalIgnoreCase).Take(10))
        {
            Console.WriteLine($"  {row.StartMs.Avg,8:0.0} ms  {row.VoiceName}");
        }

        Console.WriteLine();
        Console.WriteLine("Top by done_ms avg:");
        foreach (var row in rows.OrderBy(r => r.DoneMs.Avg).ThenBy(r => r.VoiceName, StringComparer.OrdinalIgnoreCase).Take(10))
        {
            Console.WriteLine($"  {row.DoneMs.Avg,8:0.0} ms  {row.VoiceName}");
        }
    }

    private sealed record Options(
        bool ShowHelp,
        string Text,
        int Iterations,
        int TimeoutMs,
        int ProbeTimeoutMs,
        int VoiceTimeoutMs,
        bool OutputNull,
        bool WarmUp,
        int Rate,
        int Volume,
        string? VoiceContains)
    {
        public static string HelpText =>
            """
            Usage:
              dotnet run -c Release --project .\test-voiceengin-response speed\VoiceEngineResponseSpeed.csproj -- [options]

            Options:
              --text "<text>"            Text to speak (default: "你好，Hello world.")
              --iterations <n>           Iterations per voice (default: 5)
              --timeout-ms <n>           Timeout per iteration (default: 15000)
              --probe-timeout-ms <n>     Timeout for voice enumeration (default: 8000)
              --voice-timeout-ms <n>     Timeout for each voice benchmark (default: 60000)
              --output null|default      Output target (default: null)
              --warmup true|false        Warm up each voice once (default: true)
              --rate <n>                SAPI rate -10..10 (default: 0)
              --volume <n>              Volume 0..100 (default: 100)
              --voice-contains "<part>" Filter by voice name substring
              --help                     Show help
            """;

        public static Options Parse(string[] args)
        {
            string text = "你好，Hello world.";
            int iterations = 5;
            int timeoutMs = 15000;
            int probeTimeoutMs = 8000;
            int voiceTimeoutMs = 60000;
            bool outputNull = true;
            bool warmUp = true;
            int rate = 0;
            int volume = 100;
            string? voiceContains = null;
            bool help = false;

            for (int i = 0; i < args.Length; i++)
            {
                var a = args[i];
                if (string.Equals(a, "--help", StringComparison.OrdinalIgnoreCase) || string.Equals(a, "-h", StringComparison.OrdinalIgnoreCase))
                {
                    help = true;
                    continue;
                }

                if (string.Equals(a, "--text", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    text = args[++i];
                    continue;
                }

                if (string.Equals(a, "--iterations", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var it))
                {
                    iterations = Math.Max(1, it);
                    i++;
                    continue;
                }

                if (string.Equals(a, "--timeout-ms", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var to))
                {
                    timeoutMs = Math.Max(1000, to);
                    i++;
                    continue;
                }

                if (string.Equals(a, "--probe-timeout-ms", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var pto))
                {
                    probeTimeoutMs = Math.Max(1000, pto);
                    i++;
                    continue;
                }

                if (string.Equals(a, "--voice-timeout-ms", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var vto))
                {
                    voiceTimeoutMs = Math.Max(1000, vto);
                    i++;
                    continue;
                }

                if (string.Equals(a, "--output", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    var v = args[++i].Trim();
                    outputNull = !string.Equals(v, "default", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (string.Equals(a, "--warmup", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    var v = args[++i].Trim();
                    warmUp = !string.Equals(v, "false", StringComparison.OrdinalIgnoreCase) && !string.Equals(v, "0", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (string.Equals(a, "--rate", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var r))
                {
                    rate = Math.Clamp(r, -10, 10);
                    i++;
                    continue;
                }

                if (string.Equals(a, "--volume", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var vol))
                {
                    volume = Math.Clamp(vol, 0, 100);
                    i++;
                    continue;
                }

                if (string.Equals(a, "--voice-contains", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    voiceContains = args[++i];
                    continue;
                }
            }

            if (string.IsNullOrWhiteSpace(text)) text = " ";
            return new Options(help, text, iterations, timeoutMs, probeTimeoutMs, voiceTimeoutMs, outputNull, warmUp, rate, volume, voiceContains);
        }
    }

    private sealed record ResultRow(
        string VoiceName,
        Stats StartMs,
        Stats DoneMs,
        int Failures,
        int Total);

    private readonly record struct Stats(int Count, double Min, double Avg, double P95)
    {
        public static Stats From(List<double> values)
        {
            if (values.Count == 0) return new Stats(0, 0, 0, 0);

            values.Sort();
            double min = values[0];
            double avg = values.Average();
            double p95 = Percentile(values, 0.95);
            return new Stats(values.Count, min, avg, p95);
        }

        private static double Percentile(List<double> sortedValues, double p)
        {
            if (sortedValues.Count == 0) return 0;
            if (p <= 0) return sortedValues[0];
            if (p >= 1) return sortedValues[^1];

            double idx = (sortedValues.Count - 1) * p;
            int lo = (int)Math.Floor(idx);
            int hi = (int)Math.Ceiling(idx);
            if (lo == hi) return sortedValues[lo];
            double w = idx - lo;
            return sortedValues[lo] * (1 - w) + sortedValues[hi] * w;
        }
    }

    private sealed record StaRunResult<T>(bool TimedOut, T? Value, Exception? Error);
}
