using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using FixPrinter.Services;

namespace FixPrinter;

public partial class MainWindow : Window
{
    private readonly Logger _log = Logger.Instance;
    private readonly PrinterDoctor _doctor = new();

    public MainWindow()
    {
        InitializeComponent();
        VersionText.Text = "v" + (Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "?") +
                           "  ·  github.com/" + GitHubReporter.Repo;
        _log.LineWritten += line => Dispatcher.BeginInvoke(() =>
        {
            LogBox.AppendText(line + Environment.NewLine);
            LogScroll.ScrollToEnd();
        });
        LogBox.Text = _log.GetFullLog();
        SourceInitialized += (_, _) => EnableDarkTitleBar();
    }

    private async void FixButton_Click(object sender, RoutedEventArgs e)
    {
        SetBusy(true);
        SummaryCard.Visibility = Visibility.Collapsed;
        SendLogsBigButton.Visibility = Visibility.Collapsed;
        ResultsList.ItemsSource = null;

        try
        {
            _log.Info("======= Диагностика запущена пользователем =======");
            var issues = await Task.Run(_doctor.Diagnose);
            ShowResults(issues);

            var problems = issues.Where(i => i.IsProblem).ToList();
            if (problems.Count == 0)
            {
                ShowSummary(ok: true,
                    "Проблем не обнаружено — подсистема печати в порядке.",
                    "Если принтер всё равно не печатает, нажмите «Отправить логи» внизу — " +
                    "разработчик изучит журнал и выпустит обновление.");
                return;
            }

            _log.Info($"Найдено проблем: {problems.Count}. Пытаюсь починить…");
            await Task.Run(() => _doctor.Fix(issues));

            _log.Info("Повторная проверка после починки…");
            var after = await Task.Run(_doctor.Diagnose);
            ShowResults(after);

            var remaining = after.Where(i => i.Severity == Severity.Error).ToList();
            if (remaining.Count == 0)
            {
                ShowSummary(ok: true,
                    "Готово! Найденные проблемы исправлены ✨",
                    "Попробуйте напечатать пробную страницу. Если что-то осталось — отправьте логи.");
            }
            else
            {
                ShowSummary(ok: false,
                    $"С первой попытки починить не удалось: осталось проблем — {remaining.Count}.",
                    "Отправьте логи на GitHub — разработчик посмотрит их и выпустит новую версию с исправлением " +
                    "именно вашей проблемы. Это займёт меньше минуты.");
                SendLogsBigButton.Visibility = Visibility.Visible;
            }
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void ShowResults(List<CheckResult> results)
    {
        ResultsList.ItemsSource = results
            .OrderByDescending(r => r.Severity)
            .Select(r => new ResultRow(r))
            .ToList();
    }

    private void ShowSummary(bool ok, string text, string hint)
    {
        SummaryCard.Visibility = Visibility.Visible;
        SummaryText.Text = (ok ? "✅  " : "⚠️  ") + text;
        SummaryText.Foreground = new SolidColorBrush(ok ? Color.FromRgb(0x4A, 0xDE, 0x80) : Color.FromRgb(0xFB, 0xBF, 0x24));
        SummaryHint.Text = hint;
    }

    private void SetBusy(bool busy)
    {
        FixButton.IsEnabled = !busy;
        FixButton.Content = busy ? "⏳  Работаю…" : "🔧  Найти проблемы и починить";
        Progress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
    }

    private void SendLogs_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(_log.GetFullLog());
            _log.Info("Полный лог скопирован в буфер обмена.");
        }
        catch { /* буфер может быть занят */ }
        GitHubReporter.SendLogs();
        MessageBox.Show(
            "В браузере откроется GitHub с уже заполненным отчётом — просто нажмите «Create».\n\n" +
            "Полный лог также скопирован в буфер обмена (Ctrl+V, чтобы вставить).\n" +
            "Файл лога: " + _log.LogFilePath,
            "Отправка логов", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OpenLogs_Click(object sender, RoutedEventArgs e) => GitHubReporter.OpenLogFolder();

    private void OpenManager_Click(object sender, RoutedEventArgs e)
    {
        new PrinterManagerWindow(_doctor) { Owner = this }.ShowDialog();
    }

    private void EnableDarkTitleBar()
    {
        try
        {
            var handle = new WindowInteropHelper(this).Handle;
            var on = 1;
            DwmSetWindowAttribute(handle, 20 /* DWMWA_USE_IMMERSIVE_DARK_MODE */, ref on, sizeof(int));
        }
        catch { /* не критично */ }
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    private sealed class ResultRow
    {
        public string Icon { get; }
        public string Title { get; }
        public string Details { get; }
        public Brush TitleBrush { get; }
        public Visibility DetailsVisibility { get; }

        public ResultRow(CheckResult r)
        {
            Icon = r.Severity switch
            {
                Severity.Ok => "✅",
                Severity.Warning => "⚠️",
                _ => "❌"
            };
            Title = r.Title;
            Details = r.Details;
            TitleBrush = new SolidColorBrush(r.Severity switch
            {
                Severity.Ok => Color.FromRgb(0xCB, 0xD5, 0xE1),
                Severity.Warning => Color.FromRgb(0xFB, 0xBF, 0x24),
                _ => Color.FromRgb(0xF8, 0x71, 0x71)
            });
            DetailsVisibility = string.IsNullOrEmpty(r.Details) ? Visibility.Collapsed : Visibility.Visible;
        }
    }
}
