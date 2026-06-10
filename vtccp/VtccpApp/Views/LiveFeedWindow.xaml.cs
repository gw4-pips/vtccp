namespace VtccpApp.Views;

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using VtccpApp.ViewModels;

/// <summary>
/// Code-behind for the Live View window.
///
/// ROI rubber-band interaction
/// ───────────────────────────
/// Left-click + drag anywhere in <see cref="ImagePanel"/> draws a free-form ROI
/// rectangle.  The rubber-band rectangle (<see cref="RoiBand"/>) is rendered on
/// <see cref="RoiCanvas"/> (IsHitTestVisible=False) and positioned by setting
/// Canvas.Left / Canvas.Top / Width / Height directly.
///
/// On mouse-up the rectangle's normalised position (each component in [0, 1]
/// relative to the panel's render size) is stored in the view-model via
/// <see cref="LiveFeedViewModel.SetRoi"/>.
///
/// Right-clicking anywhere clears the current ROI (visual + VM).
///
/// The normalised Rect is the authoritative value; the code-behind pixel
/// coordinates become stale on window resize but are only used for display —
/// the VM value is what downstream code uses.
/// </summary>
public partial class LiveFeedWindow : Window
{
    // ── Constructor ───────────────────────────────────────────────────────────

    public LiveFeedWindow(LiveFeedViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        Closed += (_, _) => viewModel.Dispose();
    }

    // ── ROI rubber-band state ─────────────────────────────────────────────────

    private Point _dragStart;
    private bool  _isDragging;

    // ── Mouse handlers ────────────────────────────────────────────────────────

    private void OnPanelLeftDown(object sender, MouseButtonEventArgs e)
    {
        _dragStart   = e.GetPosition(ImagePanel);
        _isDragging  = true;

        // Start with a collapsed rubber-band at the click point.
        Canvas.SetLeft(RoiBand, _dragStart.X);
        Canvas.SetTop (RoiBand, _dragStart.Y);
        RoiBand.Width      = 0;
        RoiBand.Height     = 0;
        RoiBand.Visibility = Visibility.Visible;

        ImagePanel.CaptureMouse();
        e.Handled = true;
    }

    private void OnPanelMouseMove(object sender, MouseEventArgs e)
    {
        if (!_isDragging) return;

        var pos  = e.GetPosition(ImagePanel);
        var x    = Math.Min(pos.X, _dragStart.X);
        var y    = Math.Min(pos.Y, _dragStart.Y);
        var w    = Math.Abs(pos.X - _dragStart.X);
        var h    = Math.Abs(pos.Y - _dragStart.Y);

        Canvas.SetLeft(RoiBand, x);
        Canvas.SetTop (RoiBand, y);
        RoiBand.Width  = w;
        RoiBand.Height = h;
    }

    private void OnPanelLeftUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isDragging) return;
        _isDragging = false;
        ImagePanel.ReleaseMouseCapture();

        // Commit to view-model in normalised [0,1] coordinates.
        var panelW = ImagePanel.ActualWidth;
        var panelH = ImagePanel.ActualHeight;

        if (panelW > 0 && panelH > 0)
        {
            var normRect = new Rect(
                Canvas.GetLeft(RoiBand) / panelW,
                Canvas.GetTop (RoiBand) / panelH,
                RoiBand.Width            / panelW,
                RoiBand.Height           / panelH);

            if (DataContext is LiveFeedViewModel vm)
            {
                vm.SetRoi(normRect);

                // If the rectangle was too small, clear the band too.
                if (!vm.HasRoi)
                    RoiBand.Visibility = Visibility.Collapsed;
            }
        }

        e.Handled = true;
    }

    private void OnPanelMouseLeave(object sender, MouseEventArgs e)
    {
        // Cancel an in-progress drag when the cursor leaves the panel.
        if (!_isDragging) return;
        _isDragging = false;
        ImagePanel.ReleaseMouseCapture();
        RoiBand.Visibility = Visibility.Collapsed;
    }

    private void OnPanelRightUp(object sender, MouseButtonEventArgs e)
    {
        RoiBand.Visibility = Visibility.Collapsed;

        if (DataContext is LiveFeedViewModel vm)
            vm.ClearRoi();

        // Flash a brief "ROI cleared" hint centred in the panel.
        ShowRoiClearedHint();
        e.Handled = true;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Briefly shows the "ROI cleared" label near the centre of the panel,
    /// then fades it out over 1 s.
    /// </summary>
    private void ShowRoiClearedHint()
    {
        // Centre the hint.
        Canvas.SetLeft(RoiClearedHint, (ImagePanel.ActualWidth  - 90) / 2);
        Canvas.SetTop (RoiClearedHint, (ImagePanel.ActualHeight - 28) / 2);
        RoiClearedHint.Visibility = Visibility.Visible;
        RoiClearedHint.Opacity    = 1.0;

        var fade = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(1.2))
        {
            BeginTime = TimeSpan.FromSeconds(0.3)
        };
        fade.Completed += (_, _) => RoiClearedHint.Visibility = Visibility.Collapsed;
        RoiClearedHint.BeginAnimation(OpacityProperty, fade);
    }
}
