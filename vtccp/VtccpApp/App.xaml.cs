namespace VtccpApp;

using System.Windows;
using VtccpApp.ViewModels;

public partial class App : Application
{
    private MainViewModel? _main;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        DispatcherUnhandledException += (s, args) =>
        {
            var ex = args.Exception;
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP] DispatcherUnhandledException: {ex.GetType().Name}: {ex.Message}");
            MessageBox.Show(
                $"Unhandled runtime error:\n\n{ex.GetType().Name}\n{ex.Message}" +
                $"\n\nSource: {ex.Source}\n\n--- Stack (top) ---\n" +
                (ex.StackTrace?.Length > 0
                    ? ex.StackTrace[..Math.Min(600, ex.StackTrace.Length)]
                    : "(none)"),
                "VTCCP — Unhandled Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            args.Handled = true;
        };

        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        try
        {
            _main = new MainViewModel();
            MainWindow = new MainWindow { DataContext = _main };
            MainWindow.Show();
        }
        catch (Exception ex)
        {
            var inner = ex.InnerException is not null
                ? $"\n\nInner: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}"
                : string.Empty;

            MessageBox.Show(
                $"Startup failed:\n\n{ex.GetType().Name}\n{ex.Message}{inner}" +
                $"\n\nSource: {ex.Source}\n\n--- Stack (top) ---\n" +
                (ex.StackTrace?.Length > 0
                    ? ex.StackTrace[..Math.Min(800, ex.StackTrace.Length)]
                    : "(none)"),
                "VTCCP — Startup Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        if (_main is not null) await _main.SaveConfigAsync();
        base.OnExit(e);
    }
}
