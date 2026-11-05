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

        // Лог НЕ кладём в URL: кириллица в URL-кодировке раздувает ссылку в 6 раз,
        // и GitHub отвечает «URL too long». Лог уже в буфере обмена — пользователь вставит его Ctrl+V.
        var body =
            "**Автоматический отчёт FixPrinter** — проблему не удалось устранить автоматически.\n\n" +
            $"- Версия: v{ver}\n" +
            $"- ОС: {Environment.OSVersion.VersionString}\n" +
            $"- Компьютер: {Environment.MachineName}\n\n" +
            "## Лог\n\n" +
            "Полный лог уже скопирован в буфер обмена — поставьте курсор на строку ниже и нажмите **Ctrl+V**. " +
            "Либо перетащите в это поле файл лога:\n" +
            $"`{log.LogFilePath}`\n\n" +
            "```\n(вставьте лог сюда)\n```";

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
