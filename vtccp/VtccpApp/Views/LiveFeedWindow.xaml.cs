namespace VtccpApp.Views;

using System.Windows;
using System.Windows.Input;
using VtccpApp.ViewModels;

public partial class LiveFeedWindow : Window
{
    public LiveFeedWindow(LiveFeedViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        Closed += (_, _) => viewModel.Dispose();
    }

    private void OnImagePanelClick(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is LiveFeedViewModel vm)
            vm.DismissRoi();
    }
}
