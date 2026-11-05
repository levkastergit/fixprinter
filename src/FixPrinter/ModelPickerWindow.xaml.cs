using System.Windows;
using System.Windows.Input;
using FixPrinter.Services;

namespace FixPrinter;

public partial class ModelPickerWindow : Window
{
    public string? SelectedModel { get; private set; }

    public ModelPickerWindow(PrinterDoctor.DriverInstallPlan plan)
    {
        InitializeComponent();
        ModelList.ItemsSource = plan.Models;
        ModelList.SelectedItem = plan.Suggested;
        ModelList.ScrollIntoView(plan.Suggested);
        UsbText.Text = plan.UsbDeviceIds.Count > 0
            ? "Windows видит по USB: " + string.Join(", ", plan.UsbDeviceIds)
            : "Принтер по USB сейчас не определяется — выберите модель вручную.";
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (ModelList.SelectedItem is not string model)
        {
            MessageBox.Show(this, "Выберите модель в списке.", "FixPrinter",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        SelectedModel = model;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

    private void ModelList_MouseDoubleClick(object sender, MouseButtonEventArgs e) => Ok_Click(sender, e);
}
