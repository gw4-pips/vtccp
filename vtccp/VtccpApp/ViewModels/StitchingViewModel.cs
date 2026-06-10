namespace VtccpApp.ViewModels;

using System.IO;
using System.Windows.Media.Imaging;
using DeviceInterface;
using DeviceInterface.Dmcc;
using DeviceInterface.Imaging;
using ExcelEngine.Models;
using VtccpApp.Commands;

/// <summary>
/// Drives the Symbol Stitching window (STITCH-1 Phase 1).
///
/// Workflow
/// --------
///   Idle
///     → [Capture Left]  → CapturingLeft   (GetFreshFrameAsync fires TRIGGER ON + IMAGE.SEND)
///     → LeftCaptured    left JPEG stored, Capture Right enabled
///     → [Capture Right] → CapturingRight
///     → BothCaptured    right JPEG stored, Preview Stitch enabled
///     → [Preview Stitch] composite rendered at current SeamPosition
///     → [Verify]        composite saved to temp file → IMAGE.LOAD + IMAGE.REPLAY
///     → Verifying
///     → Result          grade shown, Reset available
///
/// Seam slider
/// -----------
/// SeamPosition is a pixel column index in [0, LeftImage.PixelWidth].
/// LeftSeam = SeamPosition (cut point in left image).
/// RightSeam = SeamPosition mirrored into right image assuming 50% overlap.
///
/// OpticsSource
/// ------------
/// Composite results are tagged "StitchedImage" — distinct from LiveScan /
/// LoadedImage — so the report can carry the stitching disclaimer.
///
/// Note: algorithm parameters may need tuning once C128 FX label test images
/// are received.  See backlog STITCH-1.
/// </summary>
public sealed class StitchingViewModel : ViewModelBase, IDisposable
{
    // ── State ─────────────────────────────────────────────────────────────────

    private enum StitchState
    {
        Idle,
        CapturingLeft,
        LeftCaptured,
        CapturingRight,
        BothCaptured,
        Previewing,
        Verifying,
        Result
    }

    // ── Device coordinates ────────────────────────────────────────────────────

    private readonly string          _host;
    private readonly DeviceSession?  _session;

    // ── Frame buffers ─────────────────────────────────────────────────────────

    private byte[]? _leftJpeg;
    private byte[]? _rightJpeg;
    private byte[]? _leftCorrected;
    private byte[]? _rightCorrected;

    // ── Runtime ───────────────────────────────────────────────────────────────

    private StitchState _state = StitchState.Idle;

    // ── Bindable properties ───────────────────────────────────────────────────

    private BitmapImage? _leftImage;
    private BitmapImage? _rightImage;
    private BitmapImage? _compositeImage;
    private int          _seamPosition;
    private int          _seamMax = 100;
    private string       _statusText = "Capture the left half first, then the right half.";
    private string?      _resultText;

    public BitmapImage? LeftImage
    {
        get => _leftImage;
        private set => Set(ref _leftImage, value);
    }

    public BitmapImage? RightImage
    {
        get => _rightImage;
        private set => Set(ref _rightImage, value);
    }

    public BitmapImage? CompositeImage
    {
        get => _compositeImage;
        private set => Set(ref _compositeImage, value);
    }

    public int SeamPosition
    {
        get => _seamPosition;
        set
        {
            if (Set(ref _seamPosition, value))
                RefreshComposite();
        }
    }

    public int SeamMax
    {
        get => _seamMax;
        private set => Set(ref _seamMax, value);
    }

    public string StatusText
    {
        get => _statusText;
        private set => Set(ref _statusText, value);
    }

    public string? ResultText
    {
        get => _resultText;
        private set
        {
            if (Set(ref _resultText, value))
                OnPropertyChanged(nameof(HasResult));
        }
    }

    /// <summary>True when a verification result is available — drives result-row visibility.</summary>
    public bool HasResult => _resultText is not null;

    public bool SeamSliderVisible =>
        _state is StitchState.BothCaptured or StitchState.Previewing;

    public bool CanVerify =>
        _state == StitchState.Previewing && _session is not null;

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand CaptureLeftCommand  { get; }
    public RelayCommand CaptureRightCommand { get; }
    public RelayCommand PreviewCommand      { get; }
    public RelayCommand VerifyCommand       { get; }
    public RelayCommand ResetCommand        { get; }

    // ── Construction ─────────────────────────────────────────────────────────

    /// <param name="host">Device IP — used for GetFreshFrameAsync frame capture.</param>
    /// <param name="session">
    ///   Connected DeviceSession for IMAGE.LOAD + IMAGE.REPLAY verification.
    ///   May be null — capture and preview still work; Verify is disabled.
    /// </param>
    public StitchingViewModel(string host, DeviceSession? session = null)
    {
        _host    = host;
        _session = session;

        CaptureLeftCommand  = new RelayCommand(
            async () => await OnCaptureAsync(side: 0),
            () => _state is StitchState.Idle or StitchState.LeftCaptured or StitchState.BothCaptured
                       or StitchState.Previewing or StitchState.Result);

        CaptureRightCommand = new RelayCommand(
            async () => await OnCaptureAsync(side: 1),
            () => _state is StitchState.LeftCaptured or StitchState.BothCaptured
                       or StitchState.Previewing or StitchState.Result);

        PreviewCommand = new RelayCommand(
            OnPreview,
            () => _state == StitchState.BothCaptured);

        VerifyCommand = new RelayCommand(
            async () => await OnVerifyAsync(),
            () => CanVerify);

        ResetCommand = new RelayCommand(OnReset);
    }

    // ── Command handlers ──────────────────────────────────────────────────────

    /// <param name="side">0 = left, 1 = right.</param>
    private async Task OnCaptureAsync(int side)
    {
        _state = side == 0 ? StitchState.CapturingLeft : StitchState.CapturingRight;
        StatusText = side == 0
            ? "Capturing left half — hold the symbol steady…"
            : "Capturing right half — hold the symbol steady…";
        NotifyStateChanged();

        try
        {
            byte[]? jpeg = await LiveFeedClient.GetFreshFrameAsync(_host);
            if (jpeg is null)
            {
                StatusText = "Capture failed — no frame received. Try again.";
                _state = side == 0 ? StitchState.Idle : StitchState.LeftCaptured;
                NotifyStateChanged();
                return;
            }

            byte[] corrected = await Task.Run(() => StitchingEngine.CorrectSkew(jpeg));

            if (side == 0)
            {
                _leftJpeg      = jpeg;
                _leftCorrected = corrected;
                LeftImage      = BytesToBitmapImage(corrected);

                int suggested  = StitchingEngine.EstimateSeam(corrected);
                SeamMax        = corrected.Length > 0 ? LeftImage!.PixelWidth : 100;
                _seamPosition  = suggested;
                OnPropertyChanged(nameof(SeamPosition));

                _state     = StitchState.LeftCaptured;
                StatusText = "Left half captured — now capture the right half (overlap ~25%).";
            }
            else
            {
                _rightJpeg      = jpeg;
                _rightCorrected = corrected;
                RightImage      = BytesToBitmapImage(corrected);

                _state     = StitchState.BothCaptured;
                StatusText = "Both halves captured — press Preview Stitch, then adjust the seam.";
            }

            NotifyStateChanged();
        }
        catch (Exception ex)
        {
            StatusText = $"Capture error: {ex.Message}";
            _state = side == 0 ? StitchState.Idle : StitchState.LeftCaptured;
            NotifyStateChanged();
        }
    }

    private void OnPreview()
    {
        RefreshComposite();
        _state = StitchState.Previewing;
        StatusText = "Adjust the seam slider until the composite looks correct, then press Verify.";
        NotifyStateChanged();
    }

    private void RefreshComposite()
    {
        if (_leftCorrected is null || _rightCorrected is null) return;
        try
        {
            int leftSeam  = _seamPosition;
            int rightSeam = _seamPosition;

            byte[] composite = StitchingEngine.Composite(
                _leftCorrected, _rightCorrected, leftSeam, rightSeam);
            CompositeImage = BytesToBitmapImage(composite);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[STITCH] Composite refresh: {ex.Message}");
        }
    }

    private async Task OnVerifyAsync()
    {
        if (_session is null || _leftCorrected is null || _rightCorrected is null) return;

        _state     = StitchState.Verifying;
        StatusText = "Uploading composite to device — verifying…";
        NotifyStateChanged();

        try
        {
            byte[] composite = StitchingEngine.Composite(
                _leftCorrected, _rightCorrected, _seamPosition, _seamPosition);

            string tmpPath = Path.Combine(
                Path.GetTempPath(),
                $"vtccp_stitch_{DateTime.Now:yyyyMMdd_HHmmss}.jpg");

            await File.WriteAllBytesAsync(tmpPath, composite);

            var record = await _session.LoadImageAndVerifyAsync(tmpPath);

            File.Delete(tmpPath);

            if (record is not null)
            {
                string grade  = record.OverallGrade?.LetterGradeString ?? "?";
                string formal = record.FormalGrade ?? grade;
                ResultText = $"Grade {grade}  —  {formal}";
                StatusText = $"✓ Stitched verification complete: Grade {grade}  ({formal})";
            }
            else
            {
                StatusText = "Verification timed out — no result received. Check device.";
            }

            _state = StitchState.Result;
            NotifyStateChanged();
        }
        catch (Exception ex)
        {
            StatusText = $"Verification failed: {ex.Message}";
            _state = StitchState.Previewing;
            NotifyStateChanged();
        }
    }

    private void OnReset()
    {
        _leftJpeg       = null;
        _rightJpeg      = null;
        _leftCorrected  = null;
        _rightCorrected = null;
        LeftImage       = null;
        RightImage      = null;
        CompositeImage  = null;
        ResultText      = null;
        _seamPosition   = 0;
        OnPropertyChanged(nameof(SeamPosition));

        _state     = StitchState.Idle;
        StatusText = "Capture the left half first, then the right half.";
        NotifyStateChanged();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static BitmapImage BytesToBitmapImage(byte[] jpeg)
    {
        var bmp = new BitmapImage();
        using var ms = new MemoryStream(jpeg);
        bmp.BeginInit();
        bmp.CacheOption  = BitmapCacheOption.OnLoad;
        bmp.StreamSource = ms;
        bmp.EndInit();
        bmp.Freeze();
        return bmp;
    }

    private void NotifyStateChanged()
    {
        OnPropertyChanged(nameof(SeamSliderVisible));
        OnPropertyChanged(nameof(CanVerify));
        RelayCommand.Refresh();
    }

    // ── Disposal ──────────────────────────────────────────────────────────────

    public void Dispose() { }
}
