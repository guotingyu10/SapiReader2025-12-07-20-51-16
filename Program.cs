namespace SapiReader;

static class Program
{
    [STAThread]
    static void Main()
    {
        // 设置未处理异常的捕获
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            LogException((Exception)e.ExceptionObject, "UnhandledException");
        };

        Application.ThreadException += (sender, e) =>
        {
            LogException(e.Exception, "ThreadException");
        };

        TaskScheduler.UnobservedTaskException += (sender, e) =>
        {
            LogException(e.Exception, "UnobservedTaskException");
            e.SetObserved();
        };

        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            LogException(ex, "Main Loop Exception");
        }
    }

    static void LogException(Exception ex, string source)
    {
        string message = $"[{DateTime.Now}] {source}: {ex.Message}\n{ex.StackTrace}";
        Console.WriteLine(message);
        try
        {
            File.AppendAllText("error.log", message + "\n\n");
            MessageBox.Show(message, "Error - " + source, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch
        {
            // Ignore logging errors if we can't write to file or show messagebox
        }
    }
}
