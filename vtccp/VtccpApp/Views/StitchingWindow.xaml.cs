namespace VtccpApp.Views;

using System.Windows;
using VtccpApp.ViewModels;

public partial class StitchingWindow : Window
{
    public StitchingWindow(StitchingViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        Closed += (_, _) => viewModel.Dispose();
    }
}
