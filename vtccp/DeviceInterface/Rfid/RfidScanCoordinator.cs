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
    private readonly bool _ownsReader;
    private bool _disposed;

    /// <summary>
    /// Maximum time (ms) to wait for a TID read response after inventory.
    /// 2 000 ms is sufficient — the DLL callback fires within a few hundred ms on FW 1.8.0.
    /// </summary>
    private const int TidReadTimeoutMs = 2000;

    /// <summary>
    /// Overall budget (ms) for the CheckTagStatus (lock check) step.
    /// The vendor SDK call runs on a worker and returns the lock status directly.
    /// A timeout yields "Unknown" and never blocks the scan result.
    /// </summary>
    private const int LockCheckTimeoutMs = 5000;

    /// <summary>
    /// Raised when a scan+validation cycle completes.
    /// The second argument is the <see cref="VerificationRecord"/> that triggered the cycle.
    /// Subscribers should marshal to the UI thread if needed.
    /// </summary>
    public event AsyncEventHandler<(RfidValidationResult Result, VerificationRecord BarcodeRecord)>?
        ValidationCompleted;

    /// <param name="ownsReader">
    /// When true (default) the coordinator disposes the reader in
    /// <see cref="DisposeAsync"/>. Pass false when the reader's lifetime is
    /// managed externally (e.g. a UI-level Connect/Disconnect button that keeps
    /// the reader connected across sessions).
    /// </param>
    public RfidScanCoordinator(
        IEpcReader reader,
        RfidValidator validator,
        RfidScanCoordinatorSettings settings,
        bool ownsReader = true)
    {
        _reader     = reader    ?? throw new ArgumentNullException(nameof(reader));
        _validator  = validator ?? throw new ArgumentNullException(nameof(validator));
        _settings   = settings  ?? throw new ArgumentNullException(nameof(settings));
        _ownsReader = ownsReader;
    }

    /// <summary>
    /// Called when a barcode scan completes. Opens the RFID scan window,
    /// waits for tags, validates, then fires <see cref="ValidationCompleted"/>.
    /// Returns the <see cref="RfidValidationResult"/> so callers can embed it
    /// directly in the <see cref="VerificationRecord"/> before writing to Excel.
    /// Returns null if the coordinator is disabled or already in a scan cycle.
    /// </summary>
    public async Task<RfidValidationResult?> OnBarcodeScannedAsync(
        VerificationRecord barcodeRecord,
        CancellationToken ct = default)
    {
        if (!_settings.Enabled) return null;

        // Skip if already scanning (non-blocking trylock)
        if (!await _scanLock.WaitAsync(0, ct).ConfigureAwait(false))
            return null;

        try
        {
            var timeout = TimeSpan.FromMilliseconds(_settings.ScanWindowMs);
            var sw = System.Diagnostics.Stopwatch.StartNew();

            IReadOnlyList<EpcReadResult> reads;
            try
            {
                reads = await _reader.TriggerInventoryAsync(timeout, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { return null; }
            catch (Exception ex)
            {
                reads = [];
                // Log non-fatal read failure; validation result will be NoTag
                System.Diagnostics.Debug.WriteLine($"[RfidScanCoordinator] Read error: {ex.Message}");
            }

            // ── TID read (FW 1.8.0 workaround: called after inventory, not during) ──
            // Attempt TID read for the first/selected tag.  Non-fatal: a null TID
            // means the field is omitted from the report rather than blocking validation.
            if (reads.Count > 0 && reads[0].Tid is null)
            {
                try
                {
                    string? tid = await _reader
                        .ReadTidAsync(reads[0].EpcBytes, TimeSpan.FromMilliseconds(TidReadTimeoutMs), ct)
                        .ConfigureAwait(false);

                    if (tid is not null)
                    {
                        // EpcReadResult is an immutable record — rebuild with TID populated.
                        var updated = new List<EpcReadResult>(reads.Count);
                        updated.Add(reads[0] with { Tid = tid });
                        for (int i = 1; i < reads.Count; i++)
                            updated.Add(reads[i]);
                        reads = updated.AsReadOnly();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[RfidScanCoordinator] TID read error: {ex.Message}");
                }
            }

            // ── Lock status check (CheckTagStatus, after the TID read) ─────────────
            // Non-fatal: a null/Unknown lock status is rendered as "Unknown" in the
            // report rather than blocking validation.
            if (reads.Count > 0 && reads[0].LockStatus is null)
            {
                try
                {
                    string? lockStatus = await _reader
                        .ReadLockStatusAsync(reads[0].EpcBytes,
                            TimeSpan.FromMilliseconds(LockCheckTimeoutMs), ct)
                        .ConfigureAwait(false);

                    if (lockStatus is not null)
                    {
                        var updated = new List<EpcReadResult>(reads.Count);
                        updated.Add(reads[0] with { LockStatus = lockStatus });
                        for (int i = 1; i < reads.Count; i++)
                            updated.Add(reads[i]);
                        reads = updated.AsReadOnly();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[RfidScanCoordinator] Lock check error: {ex.Message}");
                }
            }

            int elapsed = (int)sw.ElapsedMilliseconds;
            var result  = _validator.Validate(reads, barcodeRecord, elapsed);

            if (ValidationCompleted is { } handler)
                await handler(this, (result, barcodeRecord)).ConfigureAwait(false);

            return result;
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
        if (_ownsReader)
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
