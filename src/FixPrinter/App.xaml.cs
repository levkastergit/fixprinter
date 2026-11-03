using System.Windows;
using FixPrinter.Services;

namespace FixPrinter;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, args) =>
        {
            Logger.Instance.Error("Необработанная ошибка: " + args.Exception);
            MessageBox.Show("Произошла ошибка: " + args.Exception.Message +
                            "\n\nПодробности сохранены в лог. Отправьте логи на GitHub через кнопку в приложении.",
                "FixPrinter", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };
    }
}
