using System.Diagnostics;
using System.Net;
using System.Reflection;

namespace FixPrinter.Services;

/// <summary>
/// Отправка логов на GitHub: открывает в браузере форму нового issue
/// в репозитории levkastergit/fixprinter с уже заполненным текстом лога.
/// Пользователю остаётся нажать «Create» (нужен аккаунт GitHub).
/// </summary>
public static class GitHubReporter
{
    public const string Repo = "levkastergit/fixprinter";

    public static void SendLogs()
    {
        var log = Logger.Instance;
        var ver = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "?";
        var title = $"Логи FixPrinter v{ver} — {Environment.MachineName}, {DateTime.Now:yyyy-MM-dd HH:mm}";

        var body =
            "**Автоматический отчёт FixPrinter** — принтер не удалось починить с первой попытки.\n\n" +
            $"- Версия: v{ver}\n" +
            $"- ОС: {Environment.OSVersion.VersionString}\n\n" +
            "<details><summary>Лог (хвост)</summary>\n\n```\n" +
            log.GetTail(2600) +
            "\n```\n</details>\n\n" +
            $"Полный лог сохранён на компьютере: `{log.LogFilePath}` — " +
            "при необходимости перетащите файл в это окно, чтобы приложить целиком.";

        var url = $"https://github.com/{Repo}/issues/new" +
                  $"?title={WebUtility.UrlEncode(title)}" +
                  $"&labels={WebUtility.UrlEncode("logs,auto-report")}" +
                  $"&body={WebUtility.UrlEncode(body)}";

        log.Info("Открываю форму отправки логов на GitHub…");
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    public static void OpenLogFolder()
    {
        Process.Start(new ProcessStartInfo(Logger.Instance.LogDirectory) { UseShellExecute = true });
    }
}
