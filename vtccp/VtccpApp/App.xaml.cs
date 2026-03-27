namespace VtccpApp;

using System.Windows;
using VtccpApp.ViewModels;

public partial class App : Application
{
    private MainViewModel? _main;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        _main = new MainViewModel();
        MainWindow = new MainWindow { DataContext = _main };
        MainWindow.Show();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        if (_main is not null) await _main.SaveConfigAsync();
        base.OnExit(e);
    }
}
