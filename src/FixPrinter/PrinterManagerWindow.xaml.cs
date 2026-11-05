using System.Windows;
using FixPrinter.Services;

namespace FixPrinter;

public partial class PrinterManagerWindow : Window
{
    private readonly PrinterDoctor _doctor;
    private readonly Logger _log = Logger.Instance;

    public PrinterManagerWindow(PrinterDoctor doctor)
    {
        _doctor = doctor;
        InitializeComponent();
        LoadPrinters();
    }

    private void LoadPrinters()
    {
        try
        {
            var items = _doctor.GetPrinterOverview()
                .Select(p => new Row
                {
                    Name = p.Name,
                    Badge = p.Default ? "· по умолчанию" : p.Virtual ? "· виртуальный" : "",
                    Info = $"Драйвер: {p.Driver}   Порт: {p.Port}"
                })
                .ToList();
            PrinterList.ItemsSource = items;
            StatusText.Text = $"Очередей печати: {items.Count}";
        }
        catch (Exception ex)
        {
            _log.Error("Список принтеров для менеджера: " + ex.Message);
            StatusText.Text = "Не удалось получить список принтеров";
        }
    }

    private async void Delete_Click(object sender, RoutedEventArgs e)
    {
        var selected = (PrinterList.ItemsSource as List<Row>)?
            .Where(r => r.IsSelected).Select(r => r.Name).ToList() ?? new List<string>();
        if (selected.Count == 0)
        {
            MessageBox.Show("Отметьте галочками принтеры, которые нужно удалить.",
                "FixPrinter", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var answer = MessageBox.Show(
            "Будут удалены (очередь + драйвер + порт):\n\n" +
            string.Join("\n", selected.Select(s => "• " + s)) +
            "\n\nСами устройства не пострадают: подключённый принтер всегда можно добавить заново. Продолжить?",
            "FixPrinter — подтверждение удаления",
            MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (answer != MessageBoxResult.Yes) return;

        DeleteButton.IsEnabled = false;
        StatusText.Text = "Удаляю…";
        var summary = await Task.Run(() => _doctor.RemovePrinters(selected));
        DeleteButton.IsEnabled = true;
        StatusText.Text = summary;
        LoadPrinters();
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => LoadPrinters();

    private sealed class Row
    {
        public string Name { get; init; } = "";
        public string Badge { get; init; } = "";
        public string Info { get; init; } = "";
        public bool IsSelected { get; set; }
    }
}
