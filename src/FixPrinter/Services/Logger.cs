using System.IO;
using System.Reflection;
using System.Text;

namespace FixPrinter.Services;

/// <summary>
/// Пишет лог одновременно в файл (%LOCALAPPDATA%\FixPrinter\logs) и в окно приложения.
/// </summary>
public sealed class Logger
{
    public static Logger Instance { get; } = new();

    public string LogDirectory { get; }
    public string LogFilePath { get; }

    public event Action<string>? LineWritten;

    private readonly object _lock = new();
    private readonly StringBuilder _buffer = new();

    private Logger()
    {
        LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FixPrinter", "logs");
        Directory.CreateDirectory(LogDirectory);
        LogFilePath = Path.Combine(LogDirectory,
            $"fixprinter_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");

        var ver = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "?";
        Info($"FixPrinter v{ver}");
        Info($"ОС: {Environment.OSVersion.VersionString}, 64-бит: {Environment.Is64BitOperatingSystem}");
        Info($"Компьютер: {Environment.MachineName}, пользователь: {Environment.UserName}");
    }

    public void Info(string message) => WriteCore("INFO ", message);
    public void Warn(string message) => WriteCore("WARN ", message);
    public void Error(string message) => WriteCore("ERROR", message);

    private void WriteCore(string level, string message)
    {
        var line = $"{DateTime.Now:HH:mm:ss.fff} [{level}] {message}";
        lock (_lock)
        {
            _buffer.AppendLine(line);
            try { File.AppendAllText(LogFilePath, line + Environment.NewLine, Encoding.UTF8); }
            catch { /* лог не должен ронять приложение */ }
        }
        LineWritten?.Invoke(line);
    }

    public string GetFullLog()
    {
        lock (_lock) return _buffer.ToString();
    }

    /// <summary>Хвост лога, не длиннее maxChars (для тела GitHub issue).</summary>
    public string GetTail(int maxChars)
    {
        var all = GetFullLog();
        if (all.Length <= maxChars) return all;
        var tail = all[^maxChars..];
        var nl = tail.IndexOf('\n');
        return "…(начало обрезано, полный лог в файле)…\n" + (nl >= 0 ? tail[(nl + 1)..] : tail);
    }
}
