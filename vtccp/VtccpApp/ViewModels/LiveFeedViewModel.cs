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

        // Put the device in Continuous trigger mode so the image buffer is refreshed on
        // every scan cycle — IMAGE.SEND then returns a live frame on each timer tick.
        // The timer starts after the SET fires (fire-and-forget; first tick may be blank
        // if the command hasn't completed yet, which is harmless).
        _ = SetTriggerTypeAsync(5);

        StartTimer();

        _state = FeedState.Live;
        StatusText = "Live feed active — press Verify to trigger a scan.";
        NotifyStateChanged();
    }

    private void OnVerify()
    {
        StopTimer();

        // Restore Single trigger mode and fire TRIGGER ON in one TCP session so there
        // is no race window where a stale Continuous result arrives between the two commands.
        _ = SendTriggerAsync();

        _state = FeedState.Frozen;
        StatusText = "Scan triggered — waiting for result…";
        NotifyStateChanged();
    }

    private void OnCancel()
    {
        StopTimer();
        StopSubscriber();

        // Restore Single trigger mode before leaving.
        _ = SetTriggerTypeAsync(0);

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

            // Extended ACK mode.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n",
                600, totalCts.Token);

            // Restore Single trigger mode (stops Continuous scanning) before firing
            // the verification trigger so no stale Continuous result can race in.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}SET TRIGGER.TYPE 0\r\n",
                800, totalCts.Token);

            // TRIGGER ON — fires the TruCheck verification scan.
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

    /// <summary>
    /// Sends <c>SET TRIGGER.TYPE <paramref name="triggerType"/></c> via a short-lived
    /// raw TCP connection on port 23.  Used to switch between Continuous (5) for live
    /// polling and Single (0) for idle / post-verify restore.
    /// </summary>
    private async Task SetTriggerTypeAsync(int triggerType)
    {
        try
        {
            using var cts = new CancellationTokenSource(3_000);
            using var tcp = new System.Net.Sockets.TcpClient();
            await tcp.ConnectAsync(_host, DmccCommand.RawDmccPort, cts.Token);
            using var stream = tcp.GetStream();

            try
            {
                using var bc = new CancellationTokenSource(300);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n",
                600, cts.Token);

            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}SET TRIGGER.TYPE {triggerType}\r\n",
                800, cts.Token);

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] TRIGGER.TYPE set to {triggerType}.");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] SetTriggerType({triggerType}) failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Writes an ASCII command to <paramref name="stream"/> then reads (and discards)
    /// the ACK within <paramref name="drainMs"/> milliseconds.  Swallows all errors
    /// so a missed ACK never blocks the caller.
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
    /// Called from the HttpEventSubscriber thread-pool thread when a codes.xml
    /// origin="common" result arrives.  Dispatches UI updates to the UI thread.
    /// </summary>
    private void OnResultReceived(VerificationRecord record)
    {
        // In Continuous trigger mode (Live state) the subscriber fires on every scan cycle.
        // Only process the result that was deliberately requested via Verify.
        if (_state != FeedState.Frozen) return;

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
        // Always restore Single trigger mode so the device doesn't stay in Continuous
        // after the window is closed (fire-and-forget — best effort on teardown).
        _ = SetTriggerTypeAsync(0);
    }
}
