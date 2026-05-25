namespace DeviceInterface.Dmst;

using ExcelEngine.Models;

/// <summary>
/// Watches the DMST TruCheck quality-report directory for HTML files, scrapes
/// each report as it arrives, cross-validates it against the push XML result,
/// and deletes the file.
///
/// ── Path derivation (no registry, no user config) ────────────────────────────
///
/// DMST always writes quality reports to:
///   {Documents}\{DeviceName}\CodeQuality\
/// regardless of what the "Path" field shows in DMST Options → Data Logging →
/// Reporting. The Options UI displays only the base path; DMST appends the
/// device name and "CodeQuality" subdirectory automatically.
///
/// VTCCP constructs this path from two already-known values at ConnectAsync time:
///   - Environment.SpecialFolder.MyDocuments  (Windows API — no registry)
///   - DeviceInfo.Name                        (read from device on connect)
///
/// Call BuildReportPath(deviceInfo.Name) to get the watch directory.
///
/// ── One-time prerequisite ────────────────────────────────────────────────────
///
/// In DMST Options → Data Logging → Reporting:
///   "Preferred Quality Report File Extension" must be set to ".html"
///   (default is ".pdf"; change it once and it persists).
///
/// When set to .html, DMST writes one HTML file per scan to the CodeQuality
/// directory. VTCCP picks it up within ~200 ms, scrapes it, cross-validates,
/// merges supplemental fields into the VerificationRecord, and deletes the file.
/// No per-scan setup is required; the watcher runs continuously for the session.
///
/// ── Lifecycle ────────────────────────────────────────────────────────────────
///
///   1. string watchPath = DmstHtmlScraper.BuildReportPath(deviceInfo.Name);
///   2. var scraper = new DmstHtmlScraper(watchPath);
///   3. scraper.Start();
///   4. [per scan] var enrichedRecord = await scraper.TryMergeAsync(pushRecord);
///   5. scraper.Stop() / scraper.Dispose() at session end.
///
/// Threading: FileSystemWatcher callbacks arrive on thread-pool threads.
/// Pending reports are held in a lock-protected list and matched to push records
/// on demand in TryMergeAsync.
/// </summary>
public sealed class DmstHtmlScraper : IDisposable
{
    // ── Constants ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Fixed DMST report subfolder name. DMST always creates this under
    /// {Documents}\{DeviceName}\ regardless of DMST version or configuration.
    /// Confirmed on DM475-63530E-PIPS-Verif-Lab (2026-05-25).
    /// </summary>
    public const string DmstReportSubfolder = "CodeQuality";

    /// <summary>
    /// Correlation tolerance: push XML DateTime vs HTML report DateTime.
    /// The firmware timestamp in both should agree within 1 second; 2 seconds
    /// absorbs any clock or file-write jitter.
    /// </summary>
    public static readonly TimeSpan CorrelationWindow = TimeSpan.FromSeconds(2);

    /// <summary>
    /// How long to wait for an HTML file to appear after a scan before giving up.
    /// DMST typically writes the file within 200–500 ms of the scan completing.
    /// </summary>
    public static readonly TimeSpan FileArrivalTimeout = TimeSpan.FromSeconds(4);

    // ── Fields ────────────────────────────────────────────────────────────────

    private readonly string                  _watchDirectory;
    private FileSystemWatcher?               _watcher;
    private readonly List<PendingHtmlReport> _pending = [];
    private readonly object                  _lock    = new();

    // ── Construction ──────────────────────────────────────────────────────────

    /// <summary>
    /// Initialises the scraper to watch <paramref name="watchDirectory"/>.
    /// Use <see cref="BuildReportPath"/> to obtain the correct path from the
    /// device name. The directory is created if it does not exist.
    /// </summary>
    public DmstHtmlScraper(string watchDirectory)
    {
        _watchDirectory = watchDirectory;
        Directory.CreateDirectory(watchDirectory);
    }

    // ── Path construction ─────────────────────────────────────────────────────

    /// <summary>
    /// Constructs the DMST report watch path from the device name.
    ///
    /// <paramref name="deviceName"/> is the name returned by DeviceInfo.Name at
    /// connect time — e.g. "DM475-63530E-PIPS-Verif-Lab".
    ///
    /// Result: C:\Users\{user}\Documents\DM475-63530E-PIPS-Verif-Lab\CodeQuality
    ///
    /// No registry access, no user configuration — both inputs are known to VTCCP
    /// at connect time via Windows API and the DMCC device identity response.
    /// </summary>
    public static string BuildReportPath(string deviceName)
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            deviceName,
            DmstReportSubfolder);

    // ── Lifecycle ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Starts watching the CodeQuality directory for new HTML report files.
    /// Idempotent — safe to call if already started.
    ///
    /// Prerequisite: DMST Options → Data Logging → Reporting →
    ///   "Preferred Quality Report File Extension" must be ".html".
    /// </summary>
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
            $"[VTCCP-SCRAPER] Watching '{_watchDirectory}' for DMST HTML reports.");
    }

    /// <summary>Stops the FileSystemWatcher and clears the pending list.</summary>
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

    // ── Per-scan merge ────────────────────────────────────────────────────────

    /// <summary>
    /// Waits up to <see cref="FileArrivalTimeout"/> for an HTML report that
    /// correlates to <paramref name="record"/> by
    /// <see cref="VerificationRecord.VerificationDateTime"/> (±<see cref="CorrelationWindow"/>).
    ///
    /// When a match is found, runs <see cref="DmstReportValidator.MergeAndValidate"/>:
    ///   - Supplemental fields (QR_ECLevel, QR_MaskPattern, QR_ECI, ImagePolarity)
    ///     are merged into the record and tagged in DataSourceExceptions.
    ///   - All overlapping push-XML fields are cross-validated against the HTML
    ///     values; any discrepancy is recorded in ValidationDiscrepancies.
    ///
    /// Cross-validation runs unconditionally on every scan that has an HTML report —
    /// even when SYMBOL.RESULT FULL provides some supplemental fields.
    ///
    /// Returns the original record unmodified if no correlated HTML arrives in time
    /// (e.g. DMST extension is still .pdf, or DMST is closed).
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
            $"[VTCCP-SCRAPER] No HTML report within {FileArrivalTimeout.TotalSeconds}s for " +
            $"scan at {record.VerificationDateTime:HH:mm:ss}. " +
            "Check DMST Options → Reporting → File Extension is set to .html.");

        return record;
    }

    // ── FileSystemWatcher callback ────────────────────────────────────────────

    private void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        _ = Task.Run(async () =>
        {
            // Brief settle — DMST may not have finished flushing the file.
            await Task.Delay(150);

            try
            {
                string html   = await File.ReadAllTextAsync(e.FullPath);
                var    report = ParseHtml(html, e.FullPath);

                lock (_lock) { _pending.Add(new PendingHtmlReport(report)); }

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Parsed '{Path.GetFileName(e.FullPath)}': " +
                    $"ok={report.ParseSucceeded}, dt={report.ScanDateTime?.ToString("HH:mm:ss") ?? "null"}");

                // Delete immediately — data is held in memory; file is transient.
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
    /// Parses a DMST TruCheck HTML report into a <see cref="DmstHtmlReport"/>.
    ///
    /// TODO: implement after first live HTML sample is captured.
    ///
    /// To capture the first sample:
    ///   1. In DMST Options → Data Logging → Reporting, set File Extension to .html.
    ///   2. Run one QR scan.
    ///   3. Copy the .html file from {Documents}\{DeviceName}\CodeQuality\ before it
    ///      is deleted (or add a File.Copy before the File.Delete above temporarily).
    ///   4. Inspect the file — identify field label strings, table/div structure,
    ///      DateTime format, and whether ECI/ECLevel appear in the report.
    ///
    /// Implementation strategy once format is known:
    ///   - NuGet: HtmlAgilityPack (lightweight DOM parser, no headless browser).
    ///   - XPath: //tr[td[contains(., 'Error Correction Level')]]/td[2]
    ///   - Parse ScanDateTime from report header.
    ///   - Parse numeric fields with decimal InvariantCulture.
    ///   - Return ParseSucceeded=false + ParseError if structure not recognised
    ///     (new DMST version may change layout).
    /// </summary>
    private static DmstHtmlReport ParseHtml(string htmlContent, string sourcePath)
    {
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] ParseHtml stub — {htmlContent.Length} chars from " +
            $"'{Path.GetFileName(sourcePath)}'. Implement after first live HTML sample.");

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
