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
        OrientCombo.ItemsSource = new[] { "Авто", "Книжная", "Альбомная" };
        OrientCombo.SelectedIndex = 0;
        LoadPrinters();
        _zoomTimer.Tick += (_, _) => { _zoomTimer.Stop(); _ = RefreshPreview(); };
        UpdateZoomText();
    }

    // ---------- Перетаскивание файлов ----------
    // WPF принимает файлы через OLE drag&drop, а Windows (UIPI) не пропускает OLE-перетаскивание из
    // обычного проводника в программу, запущенную от администратора. Поэтому используем классический
    // механизм WM_DROPFILES (DragAcceptFiles): его сообщения разрешаем фильтром UIPI, OLE-приёмник снимаем.
    private const int WM_DROPFILES = 0x0233;
    private const uint WM_COPYDATA = 0x004A;
    private const uint WM_COPYGLOBALDATA = 0x0049;
    private const uint MSGFLT_ALLOW = 1;

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            var source = HwndSource.FromHwnd(hwnd);
            source?.AddHook(WndProc);

            RevokeDragDrop(hwnd); // снимаем OLE-приёмник WPF, иначе проводник выберет OLE и получит отказ
            foreach (var msg in new[] { (uint)WM_DROPFILES, WM_COPYDATA, WM_COPYGLOBALDATA })
                ChangeWindowMessageFilterEx(hwnd, msg, MSGFLT_ALLOW, IntPtr.Zero);
            DragAcceptFiles(hwnd, true);
            _log.Info("Перетаскивание файлов включено (WM_DROPFILES).");
        }
        catch (Exception ex) { _log.Warn("Не удалось включить перетаскивание: " + ex.Message); }
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != WM_DROPFILES) return IntPtr.Zero;
        try
        {
            var count = DragQueryFile(wParam, 0xFFFFFFFF, null, 0);
            for (uint i = 0; i < count; i++)
            {
                var len = DragQueryFile(wParam, i, null, 0);
                var sb = new System.Text.StringBuilder((int)len + 1);
                DragQueryFile(wParam, i, sb, (uint)sb.Capacity);
                AddPath(sb.ToString());
            }
        }
        catch (Exception ex) { _log.Warn("Приём перетащенных файлов: " + ex.Message); }
        finally { DragFinish(wParam); }
        handled = true;
        return IntPtr.Zero;
    }

    /// <summary>Ctrl+V: вставка файлов, скопированных в проводнике (работает при любых правах).</summary>
    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.V || (Keyboard.Modifiers & ModifierKeys.Control) == 0) return;
        if (e.OriginalSource is System.Windows.Controls.TextBox) return; // в полях ввода — обычная вставка текста
        try
        {
            if (!Clipboard.ContainsFileDropList()) { StatusText.Text = "В буфере обмена нет файлов. Скопируйте их в проводнике (Ctrl+C)."; return; }
            foreach (var p in Clipboard.GetFileDropList()) if (p != null) AddPath(p);
            e.Handled = true;
        }
        catch (Exception ex) { StatusText.Text = "Не удалось вставить файлы: " + ex.Message; }
    }

    /// <summary>Добавляет файл, а для папки — все PDF/Word внутри неё.</summary>
    private void AddPath(string path)
    {
        if (Directory.Exists(path))
        {
            foreach (var f in Directory.EnumerateFiles(path).Where(f => Supported.Contains(Path.GetExtension(f).ToLowerInvariant())).OrderBy(f => f))
                AddFile(f);
            return;
        }
        AddFile(path);
    }

    [DllImport("shell32.dll")]
    private static extern void DragAcceptFiles(IntPtr hWnd, bool fAccept);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint DragQueryFile(IntPtr hDrop, uint iFile, System.Text.StringBuilder? lpszFile, uint cch);

    [DllImport("shell32.dll")]
    private static extern void DragFinish(IntPtr hDrop);

    [DllImport("ole32.dll")]
    private static extern int RevokeDragDrop(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool ChangeWindowMessageFilterEx(IntPtr hWnd, uint msg, uint action, IntPtr pChangeFilterStruct);

    // ---------- Масштаб предпросмотра ----------

    public static readonly DependencyProperty ThumbWidthProperty =
        DependencyProperty.Register(nameof(ThumbWidth), typeof(double), typeof(PrintWindow), new PropertyMetadata(DefaultThumb));

    public double ThumbWidth
    {
        get => (double)GetValue(ThumbWidthProperty);
        set => SetValue(ThumbWidthProperty, value);
    }

    private const double DefaultThumb = 150, MinThumb = 80, MaxThumb = 1600;
    private readonly System.Windows.Threading.DispatcherTimer _zoomTimer = new() { Interval = TimeSpan.FromMilliseconds(300) };
    private int _composedLongSide;
    private Point? _panStart;
    private double _panH, _panV;

    /// <summary>Разрешение сборки листа под текущий масштаб (с запасом на масштаб экрана Windows).</summary>
    private int DesiredLongSide()
    {
        var dpiScale = System.Windows.Media.VisualTreeHelper.GetDpi(this).DpiScaleY;
        var want = (int)(ThumbWidth * 1.4142 * Math.Max(1.0, dpiScale) * 1.1);
        var cap = (PreviewList.Items.Count > 30) ? 1400 : 2600; // ограничение памяти для больших документов
        return Math.Clamp(want, 360, cap);
    }

    private void SetZoom(double newWidth, Point? anchor = null)
    {
        newWidth = Math.Clamp(newWidth, MinThumb, MaxThumb);
        if (Math.Abs(newWidth - ThumbWidth) < 0.5) return;

        // Приближаем к точке под курсором: запоминаем относительную позицию и восстанавливаем после перерасчёта
        var a = anchor ?? new Point(PreviewScroll.ViewportWidth / 2, PreviewScroll.ViewportHeight / 2);
        var relX = PreviewScroll.ExtentWidth > 0 ? (PreviewScroll.HorizontalOffset + a.X) / PreviewScroll.ExtentWidth : 0;
        var relY = PreviewScroll.ExtentHeight > 0 ? (PreviewScroll.VerticalOffset + a.Y) / PreviewScroll.ExtentHeight : 0;

        ThumbWidth = newWidth;
        PreviewScroll.UpdateLayout();
        PreviewScroll.ScrollToHorizontalOffset(Math.Max(0, relX * PreviewScroll.ExtentWidth - a.X));
        PreviewScroll.ScrollToVerticalOffset(Math.Max(0, relY * PreviewScroll.ExtentHeight - a.Y));
        UpdateZoomText();

        // Перерисовать листы в большем разрешении, когда колёсико остановится
        if (DesiredLongSide() > _composedLongSide * 1.15) { _zoomTimer.Stop(); _zoomTimer.Start(); }
    }

    private void UpdateZoomText() => ZoomText.Text = $"{Math.Round(ThumbWidth / DefaultThumb * 100)}%";

    private void Preview_MouseWheel(object sender, MouseWheelEventArgs e)
    {
        var factor = e.Delta > 0 ? 1.15 : 1 / 1.15;
        SetZoom(ThumbWidth * factor, e.GetPosition(PreviewScroll));
        e.Handled = true; // колёсико масштабирует, а не прокручивает
    }

    private void ZoomIn_Click(object sender, RoutedEventArgs e) => SetZoom(ThumbWidth * 1.25);
    private void ZoomOut_Click(object sender, RoutedEventArgs e) => SetZoom(ThumbWidth / 1.25);
    private void ZoomReset_Click(object sender, RoutedEventArgs e) => SetZoom(DefaultThumb);
    private void ZoomFit_Click(object sender, RoutedEventArgs e) =>
        SetZoom(Math.Max(MinThumb, PreviewScroll.ViewportWidth - 40), new Point(0, 0));

    // Перетаскивание листа мышью — прокрутка (колёсико занято масштабом)
    private void Preview_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is System.Windows.Controls.Primitives.ScrollBar ||
            FindParent<System.Windows.Controls.Primitives.ScrollBar>(e.OriginalSource as DependencyObject) != null) return;
        _panStart = e.GetPosition(PreviewScroll);
        _panH = PreviewScroll.HorizontalOffset;
        _panV = PreviewScroll.VerticalOffset;
        PreviewScroll.CaptureMouse();
        PreviewScroll.Cursor = Cursors.SizeAll;
    }

    private void Preview_MouseMove(object sender, MouseEventArgs e)
    {
        if (_panStart is not Point start || e.LeftButton != MouseButtonState.Pressed) return;
        var p = e.GetPosition(PreviewScroll);
        PreviewScroll.ScrollToHorizontalOffset(_panH - (p.X - start.X));
        PreviewScroll.ScrollToVerticalOffset(_panV - (p.Y - start.Y));
    }

    private void Preview_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_panStart == null) return;
        _panStart = null;
        PreviewScroll.ReleaseMouseCapture();
        PreviewScroll.Cursor = Cursors.Hand;
    }

    private static T? FindParent<T>(DependencyObject? d) where T : DependencyObject
    {
        while (d != null)
        {
            if (d is T t) return t;
            d = d is System.Windows.Media.Visual or System.Windows.Media.Media3D.Visual3D
                ? System.Windows.Media.VisualTreeHelper.GetParent(d)
                : LogicalTreeHelper.GetParent(d); // Run и другие не-визуальные элементы
        }
        return null;
    }

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

    private async void Printer_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (PrinterCombo.SelectedItem is not string printer) return;
        // Опрос драйвера медленный — делаем в фоне, чтобы окно не подвисало
        TrayCombo.ItemsSource = new[] { "Авто" };
        TrayCombo.SelectedIndex = 0;
        DuplexCombo.ItemsSource = new[] { "Нет" };
        DuplexCombo.SelectedIndex = 0;
        DuplexCombo.IsEnabled = false;
        try
        {
            var (sources, canDuplex) = await Task.Run(() =>
                (PrintService.GetPaperSources(printer), PrintService.CanDuplex(printer)));
            if (PrinterCombo.SelectedItem as string != printer) return; // пока ждали, выбрали другой принтер

            TrayCombo.ItemsSource = new[] { "Авто" }.Concat(sources).ToList();
            TrayCombo.SelectedIndex = 0;
            var duplex = new List<string> { "Нет" };
            if (canDuplex) { duplex.Add("По длинной стороне"); duplex.Add("По короткой стороне"); }
            DuplexCombo.ItemsSource = duplex;
            DuplexCombo.SelectedIndex = 0;
            DuplexCombo.IsEnabled = duplex.Count > 1;
        }
        catch (Exception ex) { _log.Warn("Возможности принтера: " + ex.Message); }
        _ = RefreshPreview();
    }

    // Поля печати по принтеру и формату: опрос драйвера медленный, поэтому кэшируем
    private readonly Dictionary<(string, PrintService.Paper), (double L, double T, double R, double B)> _insetsCache = new();

    private async Task<(double L, double T, double R, double B)> GetInsetsAsync(PrintService.Options opt)
    {
        var key = (opt.PrinterName, opt.Paper);
        if (_insetsCache.TryGetValue(key, out var cached)) return cached;
        var insets = await Task.Run(() => _print.PrintableInsets(opt));
        _insetsCache[key] = insets;
        return insets;
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
    private void Option_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e) => _ = RefreshPreview();

    private void FitToSheet_Click(object sender, RoutedEventArgs e)
    {
        ScaleCombo.SelectedIndex = 0; // «Вписать в лист» (вызовет пересбор предпросмотра)
    }

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
        var insets = await GetInsetsAsync(opt);
        if (gen != _previewGen) return;
        var longSide = Math.Clamp((int)(ThumbWidth * 1.4142 * Math.Max(1.0, System.Windows.Media.VisualTreeHelper.GetDpi(this).DpiScaleY) * 1.1),
                                  360, indices.Count > 30 ? 1400 : 2600);

        var composed = await Task.Run(() =>
            indices.Select(i => (Index: i, Png: _print.ComposeSheet(pages[i], opt, insets, longSide))).ToList());
        if (gen != _previewGen) return;

        // Сохраняем прокрутку, чтобы при перерисовке в большем разрешении вид не прыгал
        var h = PreviewScroll.HorizontalOffset;
        var v = PreviewScroll.VerticalOffset;
        PreviewList.ItemsSource = composed.Select(c => new PreviewItem
        {
            Image = ToBitmap(c.Png),
            Caption = $"Стр. {c.Index + 1}"
        }).ToList();
        _composedLongSide = longSide;
        PreviewScroll.UpdateLayout();
        PreviewScroll.ScrollToHorizontalOffset(h);
        PreviewScroll.ScrollToVerticalOffset(v);
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
        var orient = (OrientCombo.SelectedItem as string) switch
        {
            "Книжная" => PrintService.Orient.Portrait,
            "Альбомная" => PrintService.Orient.Landscape,
            _ => PrintService.Orient.Auto
        };
        var duplex = (DuplexCombo.SelectedItem as string) switch
        {
            "По длинной стороне" => PrintService.DuplexMode.Long,
            "По короткой стороне" => PrintService.DuplexMode.Short,
            _ => PrintService.DuplexMode.None
        };
        var tray = TrayCombo.SelectedItem as string;
        return new PrintService.Options
        {
            PrinterName = PrinterCombo.SelectedItem as string ?? "",
            Paper = PaperA3.IsChecked == true ? PrintService.Paper.A3 : PrintService.Paper.A4,
            Copies = Math.Clamp(copies <= 0 ? 1 : copies, 1, 99),
            Collate = CollateBox.IsChecked == true,
            Grayscale = GrayscaleBox.IsChecked == true,
            Orientation = orient,
            Scale = scale,
            ScalePercent = percent,
            Duplex = duplex,
            PaperSource = (tray == null || tray == "Авто") ? null : tray,
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
