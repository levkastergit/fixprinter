using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using FixPrinter.Services;

namespace FixPrinter;

public partial class PrintWindow : Window
{
    private readonly PrintService _print = new();
    private readonly Logger _log = Logger.Instance;
    private readonly ObservableCollection<FileItem> _files = new();
    private readonly Dictionary<string, List<byte[]>> _rendered = new(StringComparer.OrdinalIgnoreCase);
    private const int RenderDpi = 200;
    private int _previewGen;

    private static readonly string[] Supported = { ".pdf", ".doc", ".docx", ".rtf", ".odt", ".txt" };

    public PrintWindow()
    {
        InitializeComponent();
        FileList.ItemsSource = _files;
        _files.CollectionChanged += (_, _) =>
            EmptyHint.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        ScaleCombo.ItemsSource = new[] { "Вписать в лист", "100%", "75%", "50%", "150%", "200%" };
        ScaleCombo.SelectedIndex = 0;
        LoadPrinters();
    }

    // --- Починка drag&drop для программы, запущенной от администратора (обход блокировки UIPI) ---
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try
        {
            var h = new WindowInteropHelper(this).Handle;
            foreach (uint msg in new uint[] { 0x0233 /*WM_DROPFILES*/, 0x004A /*WM_COPYDATA*/, 0x0049 /*WM_COPYGLOBALDATA*/ })
                ChangeWindowMessageFilterEx(h, msg, 1 /*MSGFLT_ALLOW*/, IntPtr.Zero);
        }
        catch (Exception ex) { _log.Warn("Не удалось разрешить перетаскивание: " + ex.Message); }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool ChangeWindowMessageFilterEx(IntPtr hWnd, uint msg, uint action, IntPtr pChangeFilterStruct);

    private void LoadPrinters()
    {
        try
        {
            var printers = PrintService.ListRealPrinters();
            PrinterCombo.ItemsSource = printers;
            PrinterCombo.SelectedItem = printers.FirstOrDefault(p => !IsVirtual(p)) ?? printers.FirstOrDefault();
        }
        catch (Exception ex) { _log.Error("Список принтеров: " + ex.Message); }
    }

    private static bool IsVirtual(string name) =>
        new[] { "PDF", "XPS", "OneNote", "Fax" }.Any(v => name.Contains(v, StringComparison.OrdinalIgnoreCase));

    // ---------- Файлы ----------

    private void AddFiles_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Выберите PDF или Word файлы",
            Multiselect = true,
            Filter = "Документы (PDF, Word)|*.pdf;*.doc;*.docx;*.rtf;*.odt;*.txt|Все файлы|*.*"
        };
        if (dlg.ShowDialog(this) == true)
            foreach (var f in dlg.FileNames) AddFile(f);
    }

    private void AddFile(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (!Supported.Contains(ext)) { StatusText.Text = $"Пропущен (не PDF/Word): {Path.GetFileName(path)}"; return; }
        if (_files.Any(f => f.Path.Equals(path, StringComparison.OrdinalIgnoreCase))) return;
        var item = new FileItem { Path = path, Name = Path.GetFileName(path), Info = "…" };
        _files.Add(item);
        if (FileList.SelectedItem == null) FileList.SelectedItem = item;
    }

    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string path)
        {
            var item = _files.FirstOrDefault(f => f.Path == path);
            if (item != null) { _files.Remove(item); _rendered.Remove(path); }
        }
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
            foreach (var p in paths) AddFile(p);
    }

    // ---------- Рендер + предпросмотр ----------

    private async Task<List<byte[]>?> EnsureRenderedAsync(FileItem item)
    {
        if (_rendered.TryGetValue(item.Path, out var cached)) return cached;
        try
        {
            var pages = await Task.Run(() => _print.RenderFile(item.Path, RenderDpi,
                msg => Dispatcher.BeginInvoke(() => StatusText.Text = msg)));
            _rendered[item.Path] = pages;
            item.Info = $"{pages.Count} стр.";
            return pages;
        }
        catch (Exception ex)
        {
            _log.Error($"Рендер «{item.Name}»: {ex.Message}");
            item.Info = "ошибка";
            StatusText.Text = $"Не удалось открыть «{item.Name}»: {ex.Message}";
            return null;
        }
    }

    private void FileList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) => _ = RefreshPreview();
    private void Option_Changed(object sender, RoutedEventArgs e) => _ = RefreshPreview();

    private void PagesBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) _ = RefreshPreview();
    }

    private async Task RefreshPreview()
    {
        if (!IsLoaded) return;
        var gen = ++_previewGen;
        if (FileList.SelectedItem is not FileItem item)
        {
            PreviewList.ItemsSource = null;
            PreviewHint.Visibility = Visibility.Visible;
            return;
        }

        PreviewHint.Visibility = Visibility.Collapsed;
        PreviewBusy.Visibility = Visibility.Visible;

        var pages = await EnsureRenderedAsync(item);
        if (gen != _previewGen) return;
        if (pages == null || pages.Count == 0) { PreviewBusy.Visibility = Visibility.Collapsed; return; }

        var opt = BuildOptions();
        var indices = PrintService.ParsePages(PagesBox.Text, pages.Count);

        var composed = await Task.Run(() =>
            indices.Select(i => (Index: i, Png: _print.ComposeSheet(pages[i], opt))).ToList());
        if (gen != _previewGen) return;

        PreviewList.ItemsSource = composed.Select(c => new PreviewItem
        {
            Image = ToBitmap(c.Png),
            Caption = $"Стр. {c.Index + 1}"
        }).ToList();
        PreviewBusy.Visibility = Visibility.Collapsed;
        StatusText.Text = $"«{item.Name}»: к печати {indices.Count} из {pages.Count} стр.";
    }

    private static BitmapImage ToBitmap(byte[] png)
    {
        var bi = new BitmapImage();
        using var ms = new MemoryStream(png);
        bi.BeginInit();
        bi.CacheOption = BitmapCacheOption.OnLoad;
        bi.StreamSource = ms;
        bi.EndInit();
        bi.Freeze();
        return bi;
    }

    // ---------- Параметры и печать ----------

    private PrintService.Options BuildOptions()
    {
        int.TryParse(CopiesBox.Text, out var copies);
        var (scale, percent) = ParseScale();
        return new PrintService.Options
        {
            PrinterName = PrinterCombo.SelectedItem as string ?? "",
            Paper = PaperA3.IsChecked == true ? PrintService.Paper.A3 : PrintService.Paper.A4,
            Copies = Math.Clamp(copies <= 0 ? 1 : copies, 1, 99),
            Grayscale = GrayscaleBox.IsChecked == true,
            AutoRotate = AutoRotateBox.IsChecked == true,
            Scale = scale,
            ScalePercent = percent,
            RenderDpi = RenderDpi
        };
    }

    private (PrintService.ScaleMode, int) ParseScale()
    {
        var s = ScaleCombo.SelectedItem as string ?? "Вписать в лист";
        if (s.StartsWith("Впис")) return (PrintService.ScaleMode.Fit, 100);
        var num = new string(s.Where(char.IsDigit).ToArray());
        return (PrintService.ScaleMode.Percent, int.TryParse(num, out var p) && p > 0 ? p : 100);
    }

    private async void Print_Click(object sender, RoutedEventArgs e)
    {
        if (_files.Count == 0)
        {
            MessageBox.Show(this, "Добавьте хотя бы один файл.", "FixPrinter", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        var opt = BuildOptions();
        if (string.IsNullOrEmpty(opt.PrinterName))
        {
            MessageBox.Show(this, "Выберите принтер.", "FixPrinter", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        PrintButton.IsEnabled = false;
        var ok = 0;
        var failed = new List<string>();
        try
        {
            foreach (var item in _files.ToList())
            {
                var pages = await EnsureRenderedAsync(item);
                if (pages == null || pages.Count == 0) { failed.Add($"{item.Name}: не удалось открыть"); continue; }
                var indices = PrintService.ParsePages(PagesBox.Text, pages.Count);
                try
                {
                    StatusText.Text = $"Печатаю «{item.Name}»…";
                    await Task.Run(() => _print.PrintPages(pages, indices, opt, item.Name));
                    ok++;
                }
                catch (Exception ex)
                {
                    _log.Error($"Печать «{item.Name}»: {ex.Message}");
                    failed.Add($"{item.Name}: {ex.Message}");
                }
            }
        }
        finally { PrintButton.IsEnabled = true; }

        if (failed.Count == 0)
        {
            StatusText.Text = $"Готово. Отправлено на принтер: {ok} док.";
            MessageBox.Show(this,
                $"Отправлено на принтер: {ok}.\n\nЕсли лист не вышел — проверьте, что принтер включён и в лотке " +
                "есть бумага нужного формата (A3/A4).",
                "FixPrinter — печать", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
            StatusText.Text = $"Отправлено: {ok}, с ошибкой: {failed.Count}.";
            MessageBox.Show(this,
                $"Успешно: {ok}.\nНе удалось:\n• " + string.Join("\n• ", failed) +
                "\n\nНажмите «Отправить логи» в главном окне — разберу причину.",
                "FixPrinter — печать", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private sealed class FileItem : System.ComponentModel.INotifyPropertyChanged
    {
        public string Path { get; init; } = "";
        public string Name { get; init; } = "";
        private string _info = "";
        public string Info
        {
            get => _info;
            set { _info = value; PropertyChanged?.Invoke(this, new(nameof(Info))); }
        }
        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    }

    private sealed class PreviewItem
    {
        public BitmapImage? Image { get; init; }
        public string Caption { get; init; } = "";
    }
}
