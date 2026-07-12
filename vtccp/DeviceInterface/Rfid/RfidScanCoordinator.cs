using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// Coordinates the RFID scan lifecycle with the barcode verification trigger.
///
/// Workflow:
///   1. Barcode verifier triggers a scan → caller notifies via <see cref="OnBarcodeScannedAsync"/>.
///   2. Coordinator opens the RFID scan window (configurable duration).
///   3. After the window closes, cross-validation is performed via <see cref="RfidValidator"/>.
///   4. Result is raised through <see cref="ValidationCompleted"/> for UI and Excel writing.
///
/// Thread safety: <see cref="OnBarcodeScannedAsync"/> is safe to call from any thread,
/// but only one scan cycle runs at a time — concurrent calls queue up.
/// </summary>
public sealed class RfidScanCoordinator : IAsyncDisposable
{
    private readonly IEpcReader _reader;
    private readonly RfidValidator _validator;
    private readonly RfidScanCoordinatorSettings _settings;
    private readonly SemaphoreSlim _scanLock = new(1, 1);
    private bool _disposed;

    /// <summary>
    /// Raised when a scan+validation cycle completes.
    /// The second argument is the <see cref="VerificationRecord"/> that triggered the cycle.
    /// Subscribers should marshal to the UI thread if needed.
    /// </summary>
    public event AsyncEventHandler<(RfidValidationResult Result, VerificationRecord BarcodeRecord)>?
        ValidationCompleted;

    public RfidScanCoordinator(
        IEpcReader reader,
        RfidValidator validator,
        RfidScanCoordinatorSettings settings)
    {
        _reader    = reader    ?? throw new ArgumentNullException(nameof(reader));
        _validator = validator ?? throw new ArgumentNullException(nameof(validator));
        _settings  = settings  ?? throw new ArgumentNullException(nameof(settings));
    }

    /// <summary>
    /// Called when a barcode scan completes. Opens the RFID scan window,
    /// waits for tags, validates, then fires <see cref="ValidationCompleted"/>.
    /// Returns immediately if the coordinator is disabled or already in a scan cycle.
    /// </summary>
    public async Task OnBarcodeScannedAsync(
        VerificationRecord barcodeRecord,
        CancellationToken ct = default)
    {
        if (!_settings.Enabled) return;

        // Skip if already scanning (non-blocking trylock)
        if (!await _scanLock.WaitAsync(0, ct).ConfigureAwait(false))
            return;

        try
        {
            var timeout = TimeSpan.FromMilliseconds(_settings.ScanWindowMs);
            var sw = System.Diagnostics.Stopwatch.StartNew();

            IReadOnlyList<EpcReadResult> reads;
            try
            {
                reads = await _reader.TriggerInventoryAsync(timeout, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex)
            {
                reads = [];
                // Log non-fatal read failure; validation result will be NoTag
                System.Diagnostics.Debug.WriteLine($"[RfidScanCoordinator] Read error: {ex.Message}");
            }

            int elapsed = (int)sw.ElapsedMilliseconds;
            var result  = _validator.Validate(reads, barcodeRecord, elapsed);

            if (ValidationCompleted is { } handler)
                await handler(this, (result, barcodeRecord)).ConfigureAwait(false);
        }
        finally
        {
            _scanLock.Release();
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        _scanLock.Dispose();
        await _reader.DisposeAsync().ConfigureAwait(false);
    }
}

/// <summary>
/// Settings governing coordinator behaviour. Typically persisted via AppSettings.
/// </summary>
public sealed class RfidScanCoordinatorSettings
{
    /// <summary>Whether RFID scanning is active. False = coordinator skips all scans.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// How long (ms) to hold the RFID scan window open after a barcode trigger.
    /// Default: 3000 ms. Range: 500–10000 ms.
    /// </summary>
    public int ScanWindowMs { get; set; } = 3000;

    /// <summary>
    /// When true, a failed RFID cross-validation (GTIN or serial mismatch) will
    /// be reflected as a soft failure flag in the VerificationRecord written to Excel.
    /// The barcode grade itself is never altered.
    /// </summary>
    public bool FlagMismatchInReport { get; set; } = true;
}

/// <summary>Async event handler delegate compatible with C# events.</summary>
public delegate Task AsyncEventHandler<TArgs>(object sender, TArgs args);
