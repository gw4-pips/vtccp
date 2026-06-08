namespace VtccpApp.Views;

using System.Windows;
using VtccpApp.ViewModels;

public partial class LiveFeedWindow : Window
{
    public LiveFeedWindow(LiveFeedViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        Closed += (_, _) => viewModel.Dispose();
    }
}
