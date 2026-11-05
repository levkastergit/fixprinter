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
        var txtPath = GitHubReporter.SendLogs();
        MessageBox.Show(
            "Открылись два окна:\n" +
            "1) Проводник — в нём выделен файл «FixPrinter-лог.txt» (лежит на рабочем столе).\n" +
            "2) Страница GitHub с формой отчёта.\n\n" +
            "ПЕРЕТАЩИТЕ файл мышкой из проводника в большое текстовое поле на странице GitHub " +
            "и нажмите зелёную кнопку Create.\n\n" +
            "Можно и без GitHub: просто отправьте этот файл в чат (Telegram и т.п.):\n" + txtPath,
            "Отправка логов", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OpenLogs_Click(object sender, RoutedEventArgs e) => GitHubReporter.OpenLogFolder();

    private void OpenManager_Click(object sender, RoutedEventArgs e)
    {
        new PrinterManagerWindow(_doctor) { Owner = this }.ShowDialog();
    }

    private async void InstallDriver_Click(object sender, RoutedEventArgs e)
    {
        var downloads = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Выберите .inf из распакованного драйвера (у Kyocera — OEMSETUP.INF в папке 64bit)",
            Filter = "Файл драйвера (*.inf)|*.inf",
            InitialDirectory = System.IO.Directory.Exists(downloads) ? downloads : ""
        };
        if (dialog.ShowDialog(this) != true) return;

        var infPath = dialog.FileName;
        string planError = "";
        var plan = await Task.Run(() => _doctor.PlanDriverInstall(infPath, out planError));
        if (plan == null)
        {
            ShowSummary(false, planError, "Выберите основной .inf драйвера (у Kyocera — OEMSETUP.INF).");
            return;
        }

        var picker = new ModelPickerWindow(plan) { Owner = this };
        if (picker.ShowDialog() != true || picker.SelectedModel == null)
        {
            _log.Info("Выбор модели отменён пользователем.");
            return;
        }
        var model = picker.SelectedModel;
        _log.Info($"Пользователь выбрал модель: «{model}»");

        SetBusy(true);
        SummaryCard.Visibility = Visibility.Collapsed;
        try
        {
            var (ok, summary) = await Task.Run(() => _doctor.InstallDriver(infPath, model));
            ShowSummary(ok, summary, ok
                ? "Напечатайте пробную страницу (Параметры → Принтеры → Печать пробной страницы) и проверьте Chrome."
                : "Нажмите «Отправить логи» — разберу, что пошло не так.");
        }
        finally
        {
            SetBusy(false);
        }
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
