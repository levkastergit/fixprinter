using System.Diagnostics;
using System.IO;
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

    /// <summary>Готовит отправку логов: копия лога кладётся на рабочий стол (.txt),
    /// проводник открывается с выделенным файлом, в браузере — форма issue.
    /// Пользователю остаётся перетащить файл в поле и нажать Create.
    /// Возвращает путь к файлу-копии.</summary>
    public static string SendLogs()
    {
        var log = Logger.Instance;
        var ver = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "?";
        var title = $"Логи FixPrinter v{ver} — {Environment.MachineName}, {DateTime.Now:yyyy-MM-dd HH:mm}";

        // Копия лога на рабочем столе: .txt надёжно прикрепляется к issue перетаскиванием
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var txtPath = Path.Combine(desktop, "FixPrinter-лог.txt");
        try { File.Copy(log.LogFilePath, txtPath, overwrite: true); }
        catch (Exception ex)
        {
            log.Warn("Не удалось положить копию лога на рабочий стол: " + ex.Message);
            txtPath = log.LogFilePath;
        }

        var body =
            "**Автоматический отчёт FixPrinter** — проблему не удалось устранить автоматически.\n\n" +
            $"- Версия: v{ver}\n" +
            $"- ОС: {Environment.OSVersion.VersionString}\n" +
            $"- Компьютер: {Environment.MachineName}\n\n" +
            "## Лог\n\n" +
            "🖱️ Рядом открылось окно проводника с выделенным файлом **FixPrinter-лог.txt** " +
            "(он лежит на рабочем столе). **Перетащите этот файл мышкой прямо в это поле** — " +
            "он прикрепится автоматически. Потом нажмите зелёную кнопку **Create**.\n\n";

        var url = $"https://github.com/{Repo}/issues/new" +
                  $"?title={WebUtility.UrlEncode(title)}" +
                  $"&labels={WebUtility.UrlEncode("logs,auto-report")}" +
                  $"&body={WebUtility.UrlEncode(body)}";

        log.Info($"Копия лога: {txtPath}. Открываю проводник и форму GitHub…");
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{txtPath}\"")
                { UseShellExecute = true });
        }
        catch (Exception ex) { log.Warn("Проводник: " + ex.Message); }
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        return txtPath;
    }

    public static void OpenLogFolder()
    {
        Process.Start(new ProcessStartInfo(Logger.Instance.LogDirectory) { UseShellExecute = true });
    }
}
