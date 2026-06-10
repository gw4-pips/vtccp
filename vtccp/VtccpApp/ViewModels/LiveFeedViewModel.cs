namespace VtccpApp.ViewModels;

using System.IO;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using DeviceInterface.Dmcc;
using DeviceInterface.Dmst;
using ExcelEngine.Models;
using VtccpApp.Commands;

/// <summary>
/// Drives the Live View window (Phase I — full-frame image feed + trigger).
///
/// State machine
/// ─────────────
///   Idle   → [Go Live]  → Live    TRIGGER ON polling starts every 400 ms,
///                                 HTTP subscriber opens, live frames appear.
///   Live   → [Verify]   → Frozen  Polling stops, Verify TRIGGER ON fired,
///                                 last full frame held until result arrives.
///   Frozen → [Go Live]  → Live    Polling restarts, subscriber reopened.
///   Live   → [Cancel]   → Idle    Polling stops, subscriber closed.
///   Frozen → [Cancel]   → Idle    Subscriber closed, last frame held.
///
/// TRIGGER.TYPE invariant
/// ──────────────────────
/// TRIGGER.TYPE is NEVER changed — it stays at 0 (Single/External) throughout
/// all states, exactly as DMST does.  The polling loop is purely client-side.
///
/// Image source
/// ────────────
/// Live frames come from IMAGE.SEND — the full camera scene (not barcode-
/// cropped) at the current IMAGE.SIZE resolution.  At DMST default
/// IMAGE.SIZE=1 (1/4 area) the output is 1224×1024 JPEG.
/// The L1 barcode-crop from push XML JpegImageBase64 is not used here.
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

    // ── ROI overlay ───────────────────────────────────────────────────────────

    private bool _roiVisible = true;

    /// <summary>
    /// True while the ROI guide rectangle is shown over the live image.
    /// Set to false when the operator clicks anywhere in the image panel.
    /// Resets to true each time Go Live is pressed.
    /// </summary>
    public bool RoiVisible
    {
        get => _roiVisible;
        private set => Set(ref _roiVisible, value);
    }

    /// <summary>Called from code-behind when the operator clicks in the image panel.</summary>
    public void DismissRoi() => RoiVisible = false;

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

        // Start client-side polling loop: TRIGGER ON + IMAGE.SEND every 400 ms.
        // TRIGGER.TYPE stays 0 — no device configuration is changed.
        StartTimer();

        RoiVisible = true;
        _state = FeedState.Live;
        StatusText = "Live feed active — press Verify to trigger a scan.";
        NotifyStateChanged();
    }

    private void OnVerify()
    {
        StopTimer();

        // Fire one verification trigger on a background thread.
        // A 400 ms delay inside SendTriggerAsync lets any in-flight poll
        // complete before the verification trigger is sent.
        _ = SendTriggerAsync();

        _state = FeedState.Frozen;
        StatusText = "Scan triggered — waiting for result…";
        NotifyStateChanged();
    }

    private void OnCancel()
    {
        StopTimer();
        StopSubscriber();

        // No device command needed — TRIGGER.TYPE is already 0 and the reader
        // returns to idle as soon as polling stops.
        _state = FeedState.Idle;
        StatusText = LiveImage is null
            ? "Feed stopped. Press Go Live to restart."
            : "Feed stopped — last frame held. Press Go Live to restart.";
        NotifyStateChanged();
    }

    // ── IMAGE.SEND polling ────────────────────────────────────────────────────

    private void StartTimer()
    {
        // 100 ms tick — fires frequently, but _fetchInProgress guard ensures only
        // one GetFreshFrameAsync runs at a time.  Effective rate is determined by
        // scan time (~400–650 ms), not by this interval.
        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
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
            byte[]? jpeg = await LiveFeedClient.GetFreshFrameAsync(_host);
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

    // ── Software trigger (Verify) ─────────────────────────────────────────────

    private async Task SendTriggerAsync()
    {
        try
        {
            // Wait for any in-flight poll cycle to finish before sending the
            // verification trigger (350 ms > one 300 ms poll interval).
            await Task.Delay(350);

            using var totalCts = new CancellationTokenSource(4_000);
            using var tcp      = new System.Net.Sockets.TcpClient();
            await tcp.ConnectAsync(_host, DmccCommand.RawDmccPort, totalCts.Token);
            using var stream = tcp.GetStream();

            // Drain welcome banner.
            try
            {
                using var bc = new CancellationTokenSource(300);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            // Extended ACK mode.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n",
                300, totalCts.Token);

            // TRIGGER ON — fires the TruCheck verification scan.
            // TRIGGER.TYPE is already 0; no mode change required.
            byte[] trigCmd = System.Text.Encoding.ASCII.GetBytes(
                $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n");
            await stream.WriteAsync(trigCmd, totalCts.Token);

            // Read ACK.
            try
            {
                using var tc  = new CancellationTokenSource(1_500);
                byte[] ackBuf = new byte[64];
                int n = await stream.ReadAsync(ackBuf, tc.Token);
                if (n > 0)
                    System.Diagnostics.Debug.WriteLine(
                        "[VTCCP-LIVEFEED] TRIGGER ACK: " +
                        System.Text.Encoding.ASCII.GetString(ackBuf, 0, n).Trim());
            }
            catch { }

            System.Diagnostics.Debug.WriteLine("[VTCCP-LIVEFEED] Verify TRIGGER ON sent.");
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

    /// <summary>
    /// Writes an ASCII command to <paramref name="stream"/> then reads (and
    /// discards) the ACK within <paramref name="drainMs"/> ms.
    /// </summary>
    private static async Task WriteAndDrainAsync(
        System.Net.Sockets.NetworkStream stream,
        string                           command,
        int                              drainMs,
        CancellationToken                ct)
    {
        await stream.WriteAsync(
            System.Text.Encoding.ASCII.GetBytes(command), ct);
        try
        {
            using var drain = CancellationTokenSource.CreateLinkedTokenSource(ct);
            drain.CancelAfter(drainMs);
            await stream.ReadAsync(new byte[64], drain.Token);
        }
        catch { }
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
    /// Called when a codes.xml origin="common" result arrives.
    /// Updates the status bar with grade + symbology.  The live IMAGE.SEND
    /// frame is kept — the L1 barcode-crop from JpegImageBase64 is not shown.
    /// </summary>
    private void OnResultReceived(VerificationRecord record)
    {
        // Only process the result that was deliberately requested via Verify.
        // The subscriber may also fire on background monitor scans; ignore those.
        if (_state != FeedState.Frozen) return;

        System.Windows.Application.Current?.Dispatcher.Invoke(() =>
        {
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
        bmp.Freeze();
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
        // TRIGGER.TYPE was never changed — no restore command needed on close.
    }
}
