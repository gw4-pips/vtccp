namespace DeviceInterface.Dmst;

using ExcelEngine.Models;

/// <summary>
/// Watches a directory for DMST TruCheck HTML reports, scrapes each one as it
/// arrives, cross-validates it against the push XML result, and deletes the file.
///
/// Lifecycle ("set before scan, clear after" pattern):
///   1. Call <see cref="Start"/> once at session start.
///   2. Before each scan trigger: call <see cref="ArmForNextScan"/>.
///      This sets the DMST auto-save path to the watch directory so DMST writes
///      the HTML report there automatically — no manual user setup required.
///   3. Trigger scan (via DeviceSession).
///   4. When push XML VerificationRecord is ready: call <see cref="TryMergeAsync"/>.
///      This waits for the HTML file, scrapes it, cross-validates, and merges.
///   5. After merge: call <see cref="DisarmAfterScan"/> to clear the DMST path.
///
/// Threading: FileSystemWatcher callbacks arrive on thread-pool threads.
/// Pending reports are held in a lock-protected list and matched to push records
/// on demand in TryMergeAsync.
///
/// TODO (pending first live HTML sample and registry key investigation):
///   - Implement <see cref="ParseHtml"/> with actual CSS/XPath selectors.
///   - Implement DmstAutoSaveConfig to write/clear the DMST auto-save registry key.
///     Check HKCU\Software\Cognex\DataMan Setup Tool\ or
///     %AppData%\Cognex\DataMan Setup Tool\settings.xml on a live DMST install.
///   - Wire ArmForNextScan / DisarmAfterScan into DeviceSession.TriggerAndCaptureAsync.
/// </summary>
public sealed class DmstHtmlScraper : IDisposable
{
    /// <summary>
    /// Correlation tolerance: push XML DateTime vs HTML report DateTime.
    /// The firmware timestamp in both should agree within 1 second; 2 seconds
    /// absorbs any clock or write-delay jitter.
    /// </summary>
    public static readonly TimeSpan CorrelationWindow = TimeSpan.FromSeconds(2);

    /// <summary>
    /// How long to wait for an HTML file to appear after a scan before giving up.
    /// DMST typically writes the file within 200–500 ms of the scan completing.
    /// </summary>
    public static readonly TimeSpan FileArrivalTimeout = TimeSpan.FromSeconds(4);

    private readonly string                   _watchDirectory;
    private FileSystemWatcher?                _watcher;
    private readonly List<PendingHtmlReport>  _pending = [];
    private readonly object                   _lock    = new();

    /// <summary>
    /// Initialises the scraper to watch <paramref name="watchDirectory"/>.
    /// The directory is created if it does not exist.
    /// </summary>
    public DmstHtmlScraper(string watchDirectory)
    {
        _watchDirectory = watchDirectory;
        Directory.CreateDirectory(watchDirectory);
    }

    // ── Lifecycle ─────────────────────────────────────────────────────────────

    /// <summary>Starts the FileSystemWatcher. Idempotent.</summary>
    public void Start()
    {
        if (_watcher is not null) return;

        _watcher = new FileSystemWatcher(_watchDirectory, "*.html")
        {
            NotifyFilter          = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents   = true,
            IncludeSubdirectories = false,
        };

        _watcher.Created += OnFileCreated;
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] Watching for DMST HTML reports in: {_watchDirectory}");
    }

    /// <summary>Stops the FileSystemWatcher and clears the pending queue.</summary>
    public void Stop()
    {
        if (_watcher is null) return;
        _watcher.EnableRaisingEvents = false;
        _watcher.Created -= OnFileCreated;
        _watcher.Dispose();
        _watcher = null;
        lock (_lock) { _pending.Clear(); }
        System.Diagnostics.Debug.WriteLine("[VTCCP-SCRAPER] Stopped.");
    }

    // ── Per-scan arming ───────────────────────────────────────────────────────

    /// <summary>
    /// Sets the DMST auto-save path to the watch directory immediately before
    /// firing a scan trigger. Call this before every DeviceSession.TriggerAsync.
    ///
    /// This is the "set before scan" half of the pattern — the user never has to
    /// manually configure the DMST auto-save path; VTCCP sets it transiently per
    /// scan and clears it in DisarmAfterScan.
    ///
    /// TODO: implement via DmstAutoSaveConfig once registry key is located.
    /// Check HKCU\Software\Cognex\DataMan Setup Tool\ on the device PC.
    /// </summary>
    public void ArmForNextScan()
    {
        // TODO: DmstAutoSaveConfig.SetSavePath(_watchDirectory);
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] ArmForNextScan: would set DMST save path to '{_watchDirectory}'. " +
            "DmstAutoSaveConfig not yet implemented — awaiting registry key.");
    }

    /// <summary>
    /// Clears the DMST auto-save path after the HTML has been scraped and merged.
    /// Call this after TryMergeAsync returns, whether or not a report was found.
    ///
    /// TODO: implement via DmstAutoSaveConfig.
    /// </summary>
    public void DisarmAfterScan()
    {
        // TODO: DmstAutoSaveConfig.ClearSavePath();
        System.Diagnostics.Debug.WriteLine(
            "[VTCCP-SCRAPER] DisarmAfterScan: would clear DMST save path. " +
            "DmstAutoSaveConfig not yet implemented.");
    }

    // ── Merge entry point ─────────────────────────────────────────────────────

    /// <summary>
    /// Waits up to <see cref="FileArrivalTimeout"/> for an HTML report correlated
    /// to <paramref name="record"/> by <see cref="VerificationRecord.VerificationDateTime"/>,
    /// then runs <see cref="DmstReportValidator.MergeAndValidate"/> and returns the
    /// enriched record.
    ///
    /// Cross-validation runs unconditionally on every scan where an HTML report
    /// arrives — even when SYMBOL.RESULT FULL provides some supplemental fields,
    /// we compare all overlapping push-XML values against the HTML report for
    /// integrity assurance and Cognex discrepancy detection.
    ///
    /// Returns the original record unmodified if no correlated HTML arrives in time.
    /// </summary>
    public async Task<VerificationRecord> TryMergeAsync(
        VerificationRecord record,
        CancellationToken  ct = default)
    {
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(FileArrivalTimeout);

        while (!timeoutCts.Token.IsCancellationRequested)
        {
            PendingHtmlReport? match = null;

            lock (_lock)
            {
                match = _pending.FirstOrDefault(p =>
                    p.Report.ParseSucceeded &&
                    p.Report.ScanDateTime.HasValue &&
                    Math.Abs((p.Report.ScanDateTime.Value - record.VerificationDateTime).TotalSeconds)
                        <= CorrelationWindow.TotalSeconds);

                if (match is not null)
                    _pending.Remove(match);
            }

            if (match is not null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Correlated HTML report to scan at " +
                    $"{record.VerificationDateTime:HH:mm:ss}. Running merge+validate.");
                return DmstReportValidator.MergeAndValidate(record, match.Report);
            }

            try { await Task.Delay(50, timeoutCts.Token); }
            catch (OperationCanceledException) { break; }
        }

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] No HTML report arrived within {FileArrivalTimeout.TotalSeconds}s " +
            $"for scan at {record.VerificationDateTime:HH:mm:ss} — emitting record without HTML data.");

        return record;
    }

    // ── FileSystemWatcher callback ────────────────────────────────────────────

    private void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        _ = Task.Run(async () =>
        {
            // Brief settle — DMST may not have finished writing.
            await Task.Delay(150);

            try
            {
                string html   = await File.ReadAllTextAsync(e.FullPath);
                var    report = ParseHtml(html, e.FullPath);

                lock (_lock) { _pending.Add(new PendingHtmlReport(report)); }

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Parsed HTML report '{Path.GetFileName(e.FullPath)}': " +
                    $"succeeded={report.ParseSucceeded}, dt={report.ScanDateTime?.ToString("HH:mm:ss") ?? "null"}");

                // Delete immediately — data is in memory, file is transient.
                File.Delete(e.FullPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Error reading/parsing '{e.FullPath}': {ex.Message}");
            }
        });
    }

    // ── HTML parser (stub) ────────────────────────────────────────────────────

    /// <summary>
    /// Parses a DMST TruCheck HTML report string into a <see cref="DmstHtmlReport"/>.
    ///
    /// TODO: Implement after first live HTML sample is captured and inspected.
    ///
    /// Implementation plan:
    ///   - Use HtmlAgilityPack (lightweight, no headless browser dependency) to
    ///     parse the HTML DOM. NuGet: HtmlAgilityPack.
    ///   - Locate field values by table row labels, e.g.:
    ///       //tr[td[contains(., 'Error Correction Level')]]/td[2]
    ///   - Parse ScanDateTime from the report header (format TBD from sample).
    ///   - Parse numeric fields with decimal InvariantCulture.
    ///   - Return ParseSucceeded=false with ParseError set if structure not recognised.
    ///
    /// Key unknowns to resolve from first live HTML:
    ///   - HTML structure: tables vs divs vs CSS classes
    ///   - Exact field label strings (may vary by DMST version or locale)
    ///   - DateTime format in the header
    ///   - Whether ECI appears as "ECI", "ECI Assignment", or similar
    ///   - Whether image polarity appears as "Normal"/"Inverted" or other values
    /// </summary>
    private static DmstHtmlReport ParseHtml(string htmlContent, string sourcePath)
    {
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] ParseHtml stub called — {htmlContent.Length} chars from " +
            $"'{Path.GetFileName(sourcePath)}'. Real parser pending first live HTML sample.");

        return new DmstHtmlReport
        {
            SourceFilePath = sourcePath,
            ParseSucceeded = false,
            ParseError     = "ParseHtml not yet implemented — awaiting first live HTML sample.",
        };
    }

    // ── IDisposable ───────────────────────────────────────────────────────────

    public void Dispose() => Stop();

    // ── Internal types ────────────────────────────────────────────────────────

    private sealed record PendingHtmlReport(DmstHtmlReport Report);
}
