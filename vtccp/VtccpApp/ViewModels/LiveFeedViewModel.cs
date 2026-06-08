namespace VtccpApp.ViewModels;

using System.IO;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using DeviceInterface.Dmcc;
using DeviceInterface.Dmst;
using ExcelEngine.Models;
using VtccpApp.Commands;

/// <summary>
/// Drives the Live View window (Phase I — image feed + trigger only).
///
/// State machine
/// ─────────────
///   Idle   → [Go Live]  → Live    IMAGE.SEND polling starts, HTTP subscriber opens
///   Live   → [Verify]   → Frozen  TRIGGER ON fired, polling stops, last frame held
///   Frozen → [Go Live]  → Live    polling restarts, subscriber reopened
///   Live   → [Cancel]   → Idle    polling stops, subscriber closed
///   Frozen → [Cancel]   → Idle    subscriber closed, last frame held
///
/// Verification result (HTTP subscriber)
/// ──────────────────────────────────────
/// When the device fires a TruCheck scan (origin="common" in codes.xml), the
/// <c>JpegImageBase64</c> barcode-crop JPEG replaces the frozen IMAGE.SEND frame
/// and the status bar shows the grade summary.  If the subscriber is not reachable
/// (e.g. no SDK port open), the last IMAGE.SEND frame simply stays frozen.
/// </summary>
public sealed class LiveFeedViewModel : ViewModelBase, IDisposable
{
    // ── State ─────────────────────────────────────────────────────────────────

    private enum FeedState { Idle, Live, Frozen }

    // ── Device coordinates ────────────────────────────────────────────────────

    private readonly string _host;
    private readonly int    _sdkPort;   // port 44444 for HttpEventSubscriber

    // ── Runtime ───────────────────────────────────────────────────────────────

    private FeedState            _state           = FeedState.Idle;
    private bool                 _fetchInProgress;
    private DispatcherTimer?     _timer;
    private HttpEventSubscriber? _subscriber;
    private readonly VerificationXmlMap _xmlMap = new();

    // ── Bindable properties ───────────────────────────────────────────────────

    private BitmapImage? _liveImage;
    private string       _statusText = "Press Go Live to start the camera feed.";

    public BitmapImage? LiveImage
    {
        get => _liveImage;
        private set => Set(ref _liveImage, value);
    }

    public string StatusText
    {
        get => _statusText;
        private set => Set(ref _statusText, value);
    }

    /// <summary>Label for the primary action button: "Go Live" or "Verify".</summary>
    public string GoLiveVerifyLabel => _state == FeedState.Live ? "Verify" : "Go Live";

    public bool CanCancel => _state != FeedState.Idle;

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand GoLiveVerifyCommand { get; }
    public RelayCommand CancelCommand       { get; }

    // ── Construction ─────────────────────────────────────────────────────────

    /// <param name="host">Device host / IP address.</param>
    /// <param name="sdkPort">Port 44444 — used by <see cref="HttpEventSubscriber"/>.</param>
    public LiveFeedViewModel(string host, int sdkPort = DmccCommand.SdkHttpPort)
    {
        _host    = host;
        _sdkPort = sdkPort;

        GoLiveVerifyCommand = new RelayCommand(OnGoLiveVerify);
        CancelCommand       = new RelayCommand(OnCancel, () => CanCancel);
    }

    // ── Command handlers ──────────────────────────────────────────────────────

    private void OnGoLiveVerify()
    {
        if (_state == FeedState.Live)
            OnVerify();
        else
            OnGoLive();
    }

    private void OnGoLive()
    {
        StopSubscriber();
        StartSubscriber();
        StartTimer();

        _state = FeedState.Live;
        StatusText = "Live feed active — press Verify to trigger a scan.";
        NotifyStateChanged();
    }

    private void OnVerify()
    {
        StopTimer();

        // Fire trigger on a background thread; UI state transitions to Frozen immediately.
        _ = SendTriggerAsync();

        _state = FeedState.Frozen;
        StatusText = "Scan triggered — waiting for result…";
        NotifyStateChanged();
    }

    private void OnCancel()
    {
        StopTimer();
        StopSubscriber();

        _state = FeedState.Idle;
        StatusText = LiveImage is null
            ? "Feed stopped. Press Go Live to restart."
            : "Feed stopped — last frame held. Press Go Live to restart.";
        NotifyStateChanged();
    }

    // ── IMAGE.SEND polling ────────────────────────────────────────────────────

    private void StartTimer()
    {
        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(333) };
        _timer.Tick += OnTimerTick;
        _timer.Start();
    }

    private void StopTimer()
    {
        _timer?.Stop();
        _timer = null;
    }

    private async void OnTimerTick(object? sender, EventArgs e)
    {
        if (_fetchInProgress) return;
        _fetchInProgress = true;
        try
        {
            byte[]? jpeg = await LiveFeedClient.GetLiveImageAsync(_host);
            if (jpeg is not null)
                LiveImage = BytesToBitmapImage(jpeg);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] Tick error: {ex.Message}");
        }
        finally
        {
            _fetchInProgress = false;
        }
    }

    // ── Software trigger ──────────────────────────────────────────────────────

    private async Task SendTriggerAsync()
    {
        try
        {
            using var totalCts = new CancellationTokenSource(4_000);
            using var tcp      = new System.Net.Sockets.TcpClient();
            await tcp.ConnectAsync(_host, DmccCommand.RawDmccPort, totalCts.Token);
            using var stream = tcp.GetStream();

            // Drain welcome banner.
            try
            {
                using var bc = new CancellationTokenSource(400);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            // Switch to Extended ACK mode so TRIGGER ON gets a ||[0]\r\n response.
            byte[] modeCmd = System.Text.Encoding.ASCII.GetBytes(
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n");
            await stream.WriteAsync(modeCmd, totalCts.Token);
            try
            {
                using var ac = new CancellationTokenSource(600);
                await stream.ReadAsync(new byte[64], ac.Token);
            }
            catch { }

            // TRIGGER ON.
            byte[] trigCmd = System.Text.Encoding.ASCII.GetBytes(
                $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n");
            await stream.WriteAsync(trigCmd, totalCts.Token);

            // Read ACK.
            try
            {
                using var tc = new CancellationTokenSource(1_500);
                byte[] ackBuf = new byte[64];
                int n = await stream.ReadAsync(ackBuf, tc.Token);
                if (n > 0)
                    System.Diagnostics.Debug.WriteLine(
                        "[VTCCP-LIVEFEED] TRIGGER ACK: " +
                        System.Text.Encoding.ASCII.GetString(ackBuf, 0, n).Trim());
            }
            catch { }

            System.Diagnostics.Debug.WriteLine("[VTCCP-LIVEFEED] TRIGGER ON sent via raw TCP.");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] Trigger failed: {ex.GetType().Name}: {ex.Message}");

            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                if (_state == FeedState.Frozen)
                    StatusText = $"Trigger failed: {ex.Message}";
            });
        }
    }

    // ── HTTP result subscriber ────────────────────────────────────────────────

    private void StartSubscriber()
    {
        var ctx = new VerificationRecord { Symbology = string.Empty, DeviceName = _host };
        _subscriber = new HttpEventSubscriber(_host, _sdkPort, _xmlMap, ctx, OnResultReceived);

        _ = _subscriber.StartAsync().ContinueWith(t =>
        {
            if (t.IsFaulted)
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] HTTP subscriber failed: " +
                    t.Exception?.GetBaseException().Message);
        });
    }

    private void StopSubscriber()
    {
        var sub = _subscriber;
        _subscriber = null;
        if (sub is not null) _ = sub.StopAsync();
    }

    /// <summary>
    /// Called from the HttpEventSubscriber thread-pool thread when a codes.xml
    /// origin="common" result arrives.  Dispatches UI updates to the UI thread.
    /// </summary>
    private void OnResultReceived(VerificationRecord record)
    {
        System.Windows.Application.Current?.Dispatcher.Invoke(() =>
        {
            // Replace the frozen IMAGE.SEND frame with the barcode-crop JPEG from the result.
            if (!string.IsNullOrEmpty(record.JpegImageBase64))
            {
                try
                {
                    byte[] jpeg = Convert.FromBase64String(record.JpegImageBase64);
                    LiveImage = BytesToBitmapImage(jpeg);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-LIVEFEED] Result JPEG decode failed: {ex.Message}");
                }
            }

            string grade = record.OverallGrade?.LetterGradeString is { Length: > 0 } g
                           ? g : "?";
            string symb  = record.Symbology is { Length: > 0 } s ? s : "?";
            StatusText   = $"Result: Grade {grade} — {symb} — press Go Live to continue.";

            _state = FeedState.Frozen;
            NotifyStateChanged();
        });
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
        bmp.Freeze();   // Required for safe hand-off to the WPF render thread.
        return bmp;
    }

    private void NotifyStateChanged()
    {
        OnPropertyChanged(nameof(GoLiveVerifyLabel));
        OnPropertyChanged(nameof(CanCancel));
        RelayCommand.Refresh();
    }

    // ── Disposal ──────────────────────────────────────────────────────────────

    public void Dispose()
    {
        StopTimer();
        StopSubscriber();
    }
}
