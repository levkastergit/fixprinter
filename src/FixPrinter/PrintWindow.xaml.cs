using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using FixPrinter.Services;

namespace FixPrinter;

public partial class PrintWindow : Window
{
    private readonly PrintService _print = new();
    private readonly Logger _log = Logger.Instance;
    private readonly ObservableCollection<FileItem> _files = new();

    private static readonly string[] Supported = { ".pdf", ".doc", ".docx", ".rtf", ".odt", ".txt" };

    public PrintWindow()
    {
        InitializeComponent();
        FileList.ItemsSource = _files;
        _files.CollectionChanged += (_, _) =>
            EmptyHint.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        LoadPrinters();
    }

    private void LoadPrinters()
    {
        try
        {
            var printers = PrintService.ListRealPrinters();
            PrinterCombo.ItemsSource = printers;
            // По умолчанию — реальный принтер (не PDF/XPS/Fax), иначе первый
            var preferred = printers.FirstOrDefault(p => !IsVirtual(p)) ?? printers.FirstOrDefault();
            PrinterCombo.SelectedItem = preferred;
        }
        catch (Exception ex)
        {
            _log.Error("Список принтеров для печати: " + ex.Message);
        }
    }

    private static bool IsVirtual(string name) =>
        new[] { "PDF", "XPS", "OneNote", "Fax" }.Any(v => name.Contains(v, StringComparison.OrdinalIgnoreCase));

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
        if (!Supported.Contains(ext))
        {
            StatusText.Text = $"Пропущен (не PDF/Word): {Path.GetFileName(path)}";
            return;
        }
        if (_files.Any(f => f.Path.Equals(path, StringComparison.OrdinalIgnoreCase))) return;
        _files.Add(new FileItem { Path = path, Name = Path.GetFileName(path) });
    }

    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string path)
        {
            var item = _files.FirstOrDefault(f => f.Path == path);
            if (item != null) _files.Remove(item);
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

    private async void Print_Click(object sender, RoutedEventArgs e)
    {
        if (_files.Count == 0)
        {
            MessageBox.Show(this, "Добавьте хотя бы один файл.", "FixPrinter",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (PrinterCombo.SelectedItem is not string printer)
        {
            MessageBox.Show(this, "Выберите принтер.", "FixPrinter",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        int.TryParse(CopiesBox.Text, out var copies);
        var opt = new PrintService.Options
        {
            PrinterName = printer,
            Paper = PaperA3.IsChecked == true ? PrintService.Paper.A3 : PrintService.Paper.A4,
            Copies = Math.Clamp(copies <= 0 ? 1 : copies, 1, 99),
            Grayscale = GrayscaleBox.IsChecked == true,
            AutoRotate = AutoRotateBox.IsChecked == true
        };

        PrintButton.IsEnabled = false;
        var files = _files.Select(f => f.Path).ToList();
        var ok = 0;
        var failed = new List<string>();
        try
        {
            foreach (var file in files)
            {
                try
                {
                    await Task.Run(() => _print.PrintDocument(file, opt,
                        msg => Dispatcher.BeginInvoke(() => StatusText.Text = msg)));
                    ok++;
                }
                catch (Exception ex)
                {
                    _log.Error($"Печать «{Path.GetFileName(file)}»: {ex.Message}");
                    failed.Add($"{Path.GetFileName(file)}: {ex.Message}");
                }
            }
        }
        finally
        {
            PrintButton.IsEnabled = true;
        }

        if (failed.Count == 0)
        {
            StatusText.Text = $"Готово. Отправлено на принтер: {ok} док. Проверьте лоток принтера.";
            MessageBox.Show(this,
                $"Отправлено на принтер: {ok}.\n\nЕсли страница не вышла — проверьте, что принтер включён и " +
                "заправлен, а в лотке есть бумага нужного формата (A3/A4).",
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

    private sealed class FileItem
    {
        public string Path { get; init; } = "";
        public string Name { get; init; } = "";
    }
}
