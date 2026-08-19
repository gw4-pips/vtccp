namespace DeviceInterface.Dmst;

using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using ExcelEngine.Models;

/// <summary>
/// Watches the DMST TruCheck quality-report directory for HTML files, scrapes
/// each report as it arrives, cross-validates it against the push XML result,
/// and preserves the original file by default.
///
/// ── Explicit report directory ────────────────────────────────────────────────
///
/// The report directory is installation-specific and cannot be safely inferred
/// from the device name. It is explicitly configured here until the application
/// exposes a user-editable setting.
///
/// ── One-time prerequisite ────────────────────────────────────────────────────
///
/// In DMST Options → Data Logging → Reporting:
///   "Preferred Quality Report File Extension" must be set to ".html"
///   (default is ".pdf"; change it once and it persists).
///
/// When set to .html, DMST writes one HTML file per scan to the CodeQuality
/// directory. VTCCP picks it up within ~200 ms, scrapes it, cross-validates,
/// and merges supplemental fields into the VerificationRecord without removing
/// the original report.
/// No per-scan setup is required; the watcher runs continuously for the session.
///
/// ── Lifecycle ────────────────────────────────────────────────────────────────
///
///   1. string watchPath = DmstHtmlScraper.ConfiguredReportDirectory;
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
    /// Explicit DMST report directory for this installation.
    /// This must be changed when the operator's DMST reporting path changes.
    /// </summary>
    public const string ConfiguredReportDirectory =
        @"C:\Users\Administrator\Documents\DM Reports & Decoded Images\DM475-866D76\CodeQuality";

    /// <summary>
    /// Linear (1D) symbology names that trigger multi-mode detection when found
    /// as a plain cell value in the Webscan TruCheck HTML report.
    /// </summary>
    private static readonly string[] KnownLinearSymbologies =
        ["EAN-13", "EAN-8", "UPC-A", "UPC-E"];

    /// <summary>
    /// How long to wait for an HTML file to appear after a scan before giving up.
    /// DMST typically writes the file within 200–500 ms of the scan completing.
    /// </summary>
    public static readonly TimeSpan FileArrivalTimeout = TimeSpan.FromSeconds(4);

    // ── Fields ────────────────────────────────────────────────────────────────

    private readonly string                  _watchDirectory;
    private FileSystemWatcher?               _watcher;
    private readonly List<PendingHtmlReport> _pending   = [];
    private readonly HashSet<string>         _ownedPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string>         _processingPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string>         _consumedPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<Task>           _processingTasks = [];
    private CancellationTokenSource?         _processingCts;
    private int                              _processingGeneration;
    private readonly object                  _lock      = new();

    // ── Source-path tracking (Replace mode support) ───────────────────────────

    /// <summary>
    /// The <see cref="DmstHtmlReport.SourceFilePath"/> of the most recently matched
    /// pending report, set each time <see cref="TryMergeAsync"/> finds a correlation.
    ///
    /// Used by <c>SessionViewModel</c> in Replace mode to write the hybrid HTML back
    /// to the same path (same folder, same filename) as the original Webscan report.
    ///
    /// Thread-safety note: written under <see cref="_lock"/> inside TryMergeAsync;
    /// callers should capture the value immediately after TryMergeAsync returns to
    /// avoid a race with the next incoming scan.
    /// </summary>
    public string? LastMatchedSourcePath { get; private set; }

    // ── Construction ──────────────────────────────────────────────────────────

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

        lock (_lock)
        {
            _processingGeneration++;
            _processingCts = new CancellationTokenSource();
        }

        _watcher = new FileSystemWatcher(_watchDirectory, "*.html")
        {
            NotifyFilter          = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents   = false,
            IncludeSubdirectories = false,
        };

        _watcher.Created += OnFileCreated;
        _watcher.Renamed += OnFileRenamed;
        _watcher.Changed += OnFileChanged;
        _watcher.EnableRaisingEvents = true;
        ScanDirectoryForReports();

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] Watching '{_watchDirectory}' for DMST HTML reports.");
    }

    /// <summary>Stops the FileSystemWatcher and clears the pending list.</summary>
    public void Stop()
    {
        if (_watcher is null) return;
        _watcher.EnableRaisingEvents = false;
        _watcher.Created -= OnFileCreated;
        _watcher.Renamed -= OnFileRenamed;
        _watcher.Changed -= OnFileChanged;
        _watcher.Dispose();
        _watcher = null;

        CancellationTokenSource? processingCts;
        Task[] processingTasks;
        lock (_lock)
        {
            _processingGeneration++;
            processingCts = _processingCts;
            _processingCts = null;
            processingTasks = _processingTasks.ToArray();
            foreach (string path in _processingPaths)
                _consumedPaths.Add(path);
        }

        processingCts?.Cancel();
        try
        {
            Task.WhenAll(processingTasks).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException)
        {
            // Expected: Stop invalidates all generation-scoped file work.
        }
        catch (AggregateException ex) when (
            ex.InnerExceptions.All(e => e is OperationCanceledException))
        {
            // Expected when multiple queued reads observe cancellation.
        }
        finally
        {
            processingCts?.Dispose();
        }

        lock (_lock)
        {
            _pending.Clear();
            _ownedPaths.Clear();
            _processingPaths.Clear();
            _processingTasks.Clear();
        }
        System.Diagnostics.Debug.WriteLine("[VTCCP-SCRAPER] Stopped.");
    }

    // ── Per-scan merge ────────────────────────────────────────────────────────

    /// <summary>
    /// Waits up to <see cref="FileArrivalTimeout"/> for an HTML report that
    /// correlates to <paramref name="record"/> by exact HTML "Verified:" text.
    /// Harmless case/whitespace differences are normalized second.  There is no
    /// timestamp or filename fallback: an uncertain association is left unmerged.
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
    ///
    /// The <c>SourcePath</c> element of the returned tuple is the
    /// <see cref="DmstHtmlReport.SourceFilePath"/> of the matched report, or <c>null</c>
    /// when no match was found.  It is returned as part of the per-call result so that
    /// concurrent callers each get their own atomic path without reading shared state.
    /// </summary>
    public async Task<(VerificationRecord Record, string? SourcePath)> TryMergeAsync(
        VerificationRecord record,
        CancellationToken  ct = default)
    {
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(FileArrivalTimeout);
        int pollsUntilRescan = 0;

        while (!timeoutCts.Token.IsCancellationRequested)
        {
            // FileSystemWatcher can miss temp-file-to-HTML renames or events that
            // happen while the watcher is starting. Periodic idempotent rescans
            // ensure an on-disk DMST report is still discovered.
            if (pollsUntilRescan-- <= 0)
            {
                ScanDirectoryForReports();
                pollsUntilRescan = 5;
            }

            PendingHtmlReport? match = null;
            string?            sourcePath;

            lock (_lock)
            {
                match = FindBestCorrelation(record);

                if (match is not null)
                {
                    _pending.Remove(match);
                    sourcePath = match.Report.SourceFilePath;
                }
                else
                {
                    sourcePath = null;
                }
            }

            if (match is not null)
            {
                lock (_lock) { LastMatchedSourcePath = match.Report.SourceFilePath; }
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Correlated HTML report to scan at " +
                    $"{record.VerificationDateTime:HH:mm:ss}; source='" +
                    $"{Path.GetFileName(match.Report.SourceFilePath)}'. Running merge+validate.");
                return (DmstReportValidator.MergeAndValidate(record, match.Report), sourcePath);
            }

            try { await Task.Delay(50, timeoutCts.Token); }
            catch (OperationCanceledException) { break; }
        }

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SCRAPER] No HTML report within {FileArrivalTimeout.TotalSeconds}s for " +
            $"scan at {record.VerificationDateTime:HH:mm:ss}. " +
            "Check DMST Options → Reporting → File Extension is set to .html.");

        return (record, null);
    }

    /// <summary>
    /// Compares two raw TruCheck "Verified:" values without applying any timezone
    /// conversion. Only harmless presentation differences are normalized.
    /// </summary>
    internal static bool VerifiedStringsEquivalent(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            return false;

        return string.Equals(
            NormalizeVerifiedForCorrelation(left),
            NormalizeVerifiedForCorrelation(right),
            StringComparison.Ordinal);
    }

    private static string NormalizeVerifiedForCorrelation(string value)
    {
        string normalized = WebUtility.HtmlDecode(value)
            .Replace('\u00a0', ' ')
            .Trim();
        normalized = Regex.Replace(normalized, @"\s+", " ");
        normalized = Regex.Replace(normalized, @"\s*\(\s*(\d+)\s*ms\s*\)\s*", "($1ms) ",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return normalized.Trim().ToUpperInvariant();
    }

    /// <summary>
    /// Matches only a verifier-supplied <c>Verified:</c> identity.  A filename
    /// timestamp is never evidence that a report belongs to a transport record.
    /// </summary>
    private PendingHtmlReport? FindBestCorrelation(VerificationRecord record)
    {
        IEnumerable<PendingHtmlReport> candidates =
            _pending.Where(p => p.Report.ParseSucceeded);

        if (string.IsNullOrWhiteSpace(record.HtmlVerifiedString))
            return null;

        PendingHtmlReport? exact = candidates.FirstOrDefault(p =>
            !string.IsNullOrWhiteSpace(p.Report.HtmlVerifiedString) &&
            string.Equals(record.HtmlVerifiedString, p.Report.HtmlVerifiedString,
                StringComparison.Ordinal));
        if (exact is not null) return exact;

        return candidates.FirstOrDefault(p =>
            VerifiedStringsEquivalent(
                record.HtmlVerifiedString, p.Report.HtmlVerifiedString));
    }

    // ── Diagnostic capture ────────────────────────────────────────────────────

    /// <summary>
    /// When true, the first HTML report received is copied to
    /// <see cref="DiagnosticCapturePath"/>.
    /// Set to true temporarily to capture an HTML sample for parser diagnostics.
    /// ParseHtml() is fully implemented and validated against the 2026-05-25 live sample.
    /// </summary>
    /// <summary>
    public bool DiagnosticCaptureEnabled { get; set; } = false;

    /// <summary>
    /// Path where the first captured HTML report is saved for inspection.
    /// Default: {Documents}\VTCCP-Diagnostic\dmst_report_sample.html
    /// Change before calling Start() if a different location is preferred.
    /// </summary>
    public string DiagnosticCapturePath { get; set; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "VTCCP-Diagnostic",
        "dmst_report_sample.html");

    private bool _diagnosticCaptured;

    // ── FileSystemWatcher callback ────────────────────────────────────────────

    /// <summary>
    /// Registers <paramref name="path"/> as a file that VTCCP itself will write (Replace mode).
    /// When the <see cref="FileSystemWatcher"/> fires for this path, <see cref="OnFileCreated"/>
    /// skips it instead of parsing and deleting it, so the hybrid report written by
    /// <c>HybridReportGenerator.SaveToPathAsync</c> is preserved in the CodeQuality folder.
    ///
    /// Call this immediately before writing the hybrid file.  The path is removed from the
    /// set on first match so it is re-processed normally on any subsequent appearance.
    /// </summary>
    public void RegisterOwnedPath(string path)
    {
        lock (_lock) { _ownedPaths.Add(path); }
    }

    private void OnFileCreated(object sender, FileSystemEventArgs e)
        => QueueFileForProcessing(e.FullPath);

    private void OnFileRenamed(object sender, RenamedEventArgs e)
        => QueueFileForProcessing(e.FullPath);

    private void OnFileChanged(object sender, FileSystemEventArgs e)
        => QueueFileForProcessing(e.FullPath);

    private void ScanDirectoryForReports()
    {
        try
        {
            foreach (string path in Directory.EnumerateFiles(
                _watchDirectory, "*.html", SearchOption.TopDirectoryOnly))
                QueueFileForProcessing(path);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] Directory rescan failed: {ex.Message}");
        }
    }

    private void QueueFileForProcessing(string path)
    {
        if (!string.Equals(Path.GetExtension(path), ".html",
                StringComparison.OrdinalIgnoreCase))
            return;

        int generation;
        CancellationToken token;
        lock (_lock)
        {
            if (_processingCts is null)
                return;

            // RegisterOwnedPath is called before SaveToPathAsync. Owned hybrid
            // output must be suppressed before the consumed-path check because Replace
            // mode intentionally reuses the original DMST path.
            if (_ownedPaths.Remove(path))
            {
                _consumedPaths.Add(path);
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Skipping owned hybrid path: '{Path.GetFileName(path)}'");
                return;
            }

            if (_consumedPaths.Contains(path) || !_processingPaths.Add(path))
                return;

            generation = _processingGeneration;
            token = _processingCts.Token;
        }

        Task task = Task.Run(
            () => ProcessFileAsync(path, generation, token), token);
        lock (_lock)
        {
            if (generation == _processingGeneration)
                _processingTasks.Add(task);
        }
        _ = task.ContinueWith(
            _ =>
            {
                lock (_lock)
                {
                    if (generation == _processingGeneration)
                        _processingTasks.Remove(task);
                }
            },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private async Task ProcessFileAsync(
        string path,
        int generation,
        CancellationToken token)
    {
        try
        {
            string html   = await ReadStableHtmlAsync(path, token);
            var    report = ParseHtml(html, path);

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] Parsed '{Path.GetFileName(path)}': " +
                $"ok={report.ParseSucceeded}, dt={report.ScanDateTime?.ToString("HH:mm:ss") ?? "null"}");

            // ── Diagnostic capture (first sample only) ──────────────────
            if (DiagnosticCaptureEnabled && !_diagnosticCaptured)
            {
                _diagnosticCaptured = true;
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(DiagnosticCapturePath)!);
                    File.Copy(path, DiagnosticCapturePath, overwrite: true);
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SCRAPER] Diagnostic copy saved → '{DiagnosticCapturePath}'");
                }
                catch (Exception copyEx)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SCRAPER] Diagnostic copy failed: {copyEx.Message}");
                }
            }

            token.ThrowIfCancellationRequested();

            if (!report.ParseSucceeded)
            {
                lock (_lock)
                {
                    if (generation == _processingGeneration)
                        _consumedPaths.Add(path);
                }
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Ignoring complete but unparseable HTML report: " +
                    $"'{Path.GetFileName(path)}'.");
                return;
            }

            // Make the parsed report available to TryMergeAsync while preserving
            // the original DMST HTML as the durable provenance artifact.
            lock (_lock)
            {
                if (generation == _processingGeneration &&
                    !token.IsCancellationRequested)
                {
                    _consumedPaths.Add(path);
                    _pending.Add(new PendingHtmlReport(report));
                }
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            // Session stopped or restarted. Generation checks prevent stale reports
            // from entering the next session's pending queue.
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] Error reading/parsing '{path}': {ex.Message}");
        }
        finally
        {
            lock (_lock)
            {
                if (generation == _processingGeneration)
                    _processingPaths.Remove(path);
            }
        }
    }

    private static async Task<string> ReadStableHtmlAsync(
        string path,
        CancellationToken token)
    {
        Exception? lastError = null;

        // A filename can become visible before DMST finishes flushing it. Require
        // stable metadata, a stable read, and a closing HTML tag before consuming
        // or deleting the report. Changed events and rescans remain retryable.
        for (int attempt = 0; attempt < 25; attempt++)
        {
            token.ThrowIfCancellationRequested();

            try
            {
                var before = new FileInfo(path);
                if (!before.Exists || before.Length == 0)
                {
                    await Task.Delay(100, token);
                    continue;
                }

                long lengthBefore = before.Length;
                DateTime writeBefore = before.LastWriteTimeUtc;

                await Task.Delay(100, token);

                var settled = new FileInfo(path);
                if (!settled.Exists ||
                    settled.Length != lengthBefore ||
                    settled.LastWriteTimeUtc != writeBefore)
                    continue;

                string html;
                await using (var stream = new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete,
                    bufferSize: 4096,
                    useAsync: true))
                using (var reader = new StreamReader(stream))
                    html = await reader.ReadToEndAsync(token);

                var after = new FileInfo(path);
                if (!after.Exists ||
                    after.Length != lengthBefore ||
                    after.LastWriteTimeUtc != writeBefore)
                    continue;

                if (!Regex.IsMatch(html, @"</html\s*>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    await Task.Delay(100, token);
                    continue;
                }

                return html;
            }
            catch (Exception ex) when (
                ex is IOException or UnauthorizedAccessException)
            {
                lastError = ex;
                await Task.Delay(100, token);
            }
        }

        throw new IOException(
            $"HTML report did not become complete and stable: '{path}'.",
            lastError);
    }

    // ── HTML parser ───────────────────────────────────────────────────────────

    /// <summary>
    /// Parses a DMST TruCheck HTML report into a <see cref="DmstHtmlReport"/>.
    ///
    /// Format confirmed 2026-05-25 from live scan on fw 6.1.16_sr4 (QR GUID, Grade A).
    /// File: 2026-05-24_23-03-58-752_1779678267324.html
    ///
    /// ── Structure ────────────────────────────────────────────────────────────
    ///
    /// The DMST HTML report is a single minified line containing two tables:
    ///
    ///   1. Header table (cells 0–30 approx):
    ///      Labels and values are in SEPARATE rows (multi-column layout).
    ///      Do NOT use consecutive-pair extraction here.
    ///      OverallGrade is extracted by searching for the "D.D (L)" pattern.
    ///
    ///   2. Simple characteristics table (cells 31–60 approx):
    ///      Consecutive &lt;td&gt;Label&lt;/td&gt;&lt;td&gt;Value&lt;/td&gt; pairs.
    ///      All 4 target fields live here.
    ///      Indexed positions (confirmed, QR Grade A scan):
    ///        [31/32]="QR Size"/"29x29"
    ///        [33/34]="Horizontal BWG"/"-3%"
    ///        [35/36]="Vertical BWG"/"-4%"
    ///        [37/38]="Encoded characters"/"36"   ← HTML authoritative (push XML: 39 = WRONG)
    ///        [39/40]="Total Codewords"/"70"
    ///        [41/42]="Data Codewords"/"44"        ← empty in push XML; HTML authoritative
    ///        [43/44]="Error Correction Budget"/"26" ← empty in push XML; HTML authoritative
    ///        [45/46]="Errors Corrected"/"0"
    ///        [47/48]="Error Capacity Used"/"0"
    ///        [49/50]="Error Correction Level"/"M"  ← PRIMARY TARGET
    ///        [51/52]="Data Mask Pattern"/"2"        ← PRIMARY TARGET
    ///        [53/54]="Image"/"Black on white"       ← PRIMARY TARGET (ImagePolarity)
    ///        [55/56]="Nominal X Dim"/"12.6 mil"
    ///        [57/58]="Pixels per Module"/"9.75"
    ///        [59/60]="ECI"/"000003"                 ← PRIMARY TARGET
    ///
    ///   3. Grade parameters table (cells 61+ approx):
    ///      6-cell rows: [label][secondary][pct%][numeric][letter][PASS/FAIL]
    ///      UEC row: [61]="1. Unused Error Correction (UEC)"  [62]=""  [63]="100.0%"  ...
    ///      SC row:  [67]="2. Symbol Contrast (SC)"           [68]="Rl/Rd (87/6)"  [69]="nan%"  ...
    ///      ANU row: [85]="4. Axial Nonuniformity (ANU)"      [86]=""  [87]="0.8%"  ...
    ///      GNU row: [91]="5. Grid Nonuniformity (GNU)"        [92]=""  [93]="0.0%"  ...
    ///      Note: SC shows "nan%" for IMAGE.LOAD scans (loaded-image scan, no live illumination).
    ///
    /// ── DateTime ─────────────────────────────────────────────────────────────
    ///
    /// The in-page DateTime header is CORRUPT ("31-Dec-1970 07:00:00" = Unix epoch).
    /// ScanDateTime is parsed from the filename prefix: "yyyy-MM-dd_HH-mm-ss-mmm_..."
    ///
    /// ── No external library required ─────────────────────────────────────────
    ///
    /// The cell extraction uses <c>Regex.Matches</c> on &lt;td&gt; elements.
    /// The minified single-line HTML makes regex extraction reliable and fast.
    /// HtmlAgilityPack is not needed and has not been added as a dependency.
    /// </summary>
    internal static DmstHtmlReport ParseHtml(
        string htmlContent,
        string sourcePath,
        bool hasSyntheticSourcePath = false)
    {
        try
        {
            // ── Step 1: extract all <td> text in document order ──────────────
            //
            // Empty cells ARE included — they matter for grade-row offset arithmetic.
            // WebUtility.HtmlDecode converts &amp; &lt; &gt; etc. in cell text.
            var cells = Regex.Matches(htmlContent, @"<td[^>]*>(.*?)</td>", RegexOptions.Singleline)
                .Select(m => WebUtility.HtmlDecode(
                    Regex.Replace(m.Groups[1].Value, "<[^>]+>", "").Trim()))
                .ToList();

            // ── Step 2: build consecutive-pair label→value lookup ─────────────
            //
            // Covers the simple characteristics table cleanly.
            // The header section's wrong pairs ("Standard"→"Grade" etc.) use generic
            // labels that are never looked up by the code below — safe to include.
            var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < cells.Count - 1; i++)
                if (!string.IsNullOrEmpty(cells[i]))
                    lookup.TryAdd(cells[i], cells[i + 1]);

            string? Get(string label)
            {
                if (!lookup.TryGetValue(label, out var v)) return null;
                return string.IsNullOrEmpty(v) ? null : v;
            }

            string? GetAny(params string[] labels)
            {
                foreach (string label in labels)
                {
                    string? value = Get(label);
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
                return null;
            }

            int? GetInt(string label)
            {
                var s = Get(label);
                return int.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;
            }

            decimal? GetDecimal(string label, bool stripPercent = false)
            {
                var s = Get(label);
                if (stripPercent) s = s?.TrimEnd('%').Trim();
                return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;
            }

            // ── Step 2b: DM matrix size normalisation ─────────────────────────
            //
            // QR HTML uses label "QR Size"     → value "29x29" (clean, no suffix).
            // DM  HTML uses label "Matrix Size" → value "16x36 (Data: 14x34)"
            //   The " (Data: RxC)" suffix lists the data-only region (without the
            //   finder/clock border) and must be stripped so the value matches the
            //   format used by the push XML parser and ECC200 lookup table.
            static string? StripDataSuffix(string? raw)
            {
                if (raw is null) return null;
                var idx = raw.IndexOf(" (Data:", StringComparison.OrdinalIgnoreCase);
                return idx >= 0 ? raw[..idx].Trim() : raw;
            }

            // ── Step 3: grade parameter rows ──────────────────────────────────
            //
            // Row structure: [label][secondary][pct%][numeric][letter][PASS/FAIL]
            // Scan ahead up to 5 cells for the first value ending in "%" that is
            // not "nan%" (nan% = IMAGE.LOAD scan with no live illumination → null).
            decimal? GetGradePct(string label)
            {
                int idx = cells.FindIndex(
                    c => string.Equals(c, label, StringComparison.OrdinalIgnoreCase));
                if (idx < 0) return null;
                for (int i = idx + 1; i < Math.Min(idx + 5, cells.Count); i++)
                {
                    var cell = cells[i];
                    if (cell.EndsWith('%') && cell != "nan%")
                    {
                        var num = cell.TrimEnd('%').Trim();
                        if (decimal.TryParse(num, NumberStyles.Any,
                                             CultureInfo.InvariantCulture, out var pct))
                            return pct;
                    }
                }
                return null;
            }

            // ── Step 4: overall grade letter(s) from "D.D (L)" header pattern ──
            //
            // Single-mode: one occurrence near cell 19.
            // Multi-mode:  two occurrences — linear (EAN/UPC) section first, then 2D.
            //
            // Strategy: collect ALL cells matching "D.D (L)" (grade-parameter rows
            // use separate numeric and letter cells, so this pattern is unique to the
            // header summary area).  Assign first → linear (if multi-mode) or 2D (if
            // single-mode); assign second → 2D (multi-mode only).
            var allGradeDisplays = cells
                .Where(c => Regex.IsMatch(c, @"^\d+\.\d+\s*\([A-Fa-f]\)$"))
                .ToList();

            string? overallGrade              = null;
            string? linearOverallGrade        = null;
            decimal? linearOverallGradeNumeric = null;

            // Detect multi-mode: requires BOTH a known linear symbology cell AND at
            // least two "D.D (L)" grade patterns.
            //
            // A standalone EAN/UPC report has a linear symbology cell but only one
            // grade display → isMultiMode = false, so Linear* fields remain null.
            // Only when two grade patterns are present (one per symbol) does the
            // linear-first ordering assignment apply.
            int linearSymbIdx = cells.FindIndex(
                c => KnownLinearSymbologies.Contains(c, StringComparer.OrdinalIgnoreCase));
            bool isMultiMode = linearSymbIdx >= 0 && allGradeDisplays.Count >= 2;

            // Shared helper: parse letter and numeric from "D.D (L)" string.
            static (string Letter, decimal Numeric) ParseGradeDisplay(string display)
            {
                var m = Regex.Match(display, @"(\d+\.\d+)\s*\(([A-Fa-f])\)");
                if (!m.Success) return (string.Empty, 0m);
                string letter = m.Groups[2].Value.ToUpperInvariant();
                decimal.TryParse(m.Groups[1].Value, NumberStyles.Any,
                                 CultureInfo.InvariantCulture, out decimal numeric);
                return (letter, numeric);
            }

            if (isMultiMode)
            {
                // First "D.D (L)" = linear (1D) symbol grade; second = 2D symbol grade.
                if (allGradeDisplays.Count >= 1)
                {
                    var (lLetter, lNumeric) = ParseGradeDisplay(allGradeDisplays[0]);
                    if (lLetter.Length > 0)
                    {
                        linearOverallGrade        = lLetter;
                        linearOverallGradeNumeric = lNumeric;
                    }
                }
                if (allGradeDisplays.Count >= 2)
                {
                    var (dLetter, _) = ParseGradeDisplay(allGradeDisplays[1]);
                    if (dLetter.Length > 0) overallGrade = dLetter;
                }
            }
            else
            {
                // Single-mode: first (only) grade is for whatever symbol was scanned.
                if (allGradeDisplays.Count >= 1)
                {
                    var (letter, _) = ParseGradeDisplay(allGradeDisplays[0]);
                    if (letter.Length > 0) overallGrade = letter;
                }
            }

            // ── Step 4b: multi-mode linear symbol extraction ──────────────────
            //
            // When IsMultiMode is true, extract the EAN/UPC symbol's characteristics
            // from the cells near the linear symbology marker.
            //
            // Assumptions (validated against Webscan TruCheck multi-mode layout):
            //   • LinearSymbology cell is followed within ~10 cells by the digit string
            //     (the decoded EAN/UPC data — always all-numeric, 8–13 digits).
            //   • The linear formal grade appears in the cells as "X/D+/D+[/text]"
            //     where X is a letter grade (A–F).  This distinguishes it from the
            //     2D formal grade which starts with a decimal number (e.g., "4.0/10/…").
            //   • Aperture, wavelength, and lighting are parsed from the formal grade.
            //   • Standard is fixed as "ISO/IEC 15416" for all 1D EAN/UPC symbols.

            string? linearSymbology   = isMultiMode ? cells[linearSymbIdx] : null;
            string? linearDecodedData = null;
            string? linearFormalGrade = null;
            int?    linearAperture    = null;
            int?    linearWavelength  = null;
            string? linearLighting    = null;
            string? linearStandard    = isMultiMode ? "ISO/IEC 15416" : null;

            if (isMultiMode)
            {
                // Decoded data: first all-digit string (8–14 chars) after the symbology cell.
                for (int i = linearSymbIdx + 1;
                     i < Math.Min(linearSymbIdx + 10, cells.Count); i++)
                {
                    var c = cells[i];
                    if (c.Length >= 8 && c.Length <= 14 && c.All(char.IsDigit))
                    {
                        linearDecodedData = c;
                        break;
                    }
                }

                // Formal grade: "Letter/ApertureDigits/WavelengthDigits[/Lighting]"
                // Search from the start of the linear section up to 80 cells ahead.
                // Pattern: grade letter then two slash-separated numeric groups, optional lighting.
                var formalGradeRx = new Regex(
                    @"^([A-Fa-f])/(\d{1,3})/(\d{3,4})(?:/(.+))?$",
                    RegexOptions.Compiled);

                for (int i = linearSymbIdx;
                     i < Math.Min(linearSymbIdx + 80, cells.Count); i++)
                {
                    var fm = formalGradeRx.Match(cells[i]);
                    if (!fm.Success) continue;

                    linearFormalGrade = cells[i];
                    if (int.TryParse(fm.Groups[2].Value,
                            NumberStyles.Any, CultureInfo.InvariantCulture, out int ap))
                        linearAperture = ap;
                    if (int.TryParse(fm.Groups[3].Value,
                            NumberStyles.Any, CultureInfo.InvariantCulture, out int wl))
                        linearWavelength = wl;
                    if (fm.Groups[4].Success && !string.IsNullOrWhiteSpace(fm.Groups[4].Value))
                        linearLighting = fm.Groups[4].Value.Trim();
                    break;
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Multi-mode linear: symb={linearSymbology} " +
                    $"data={linearDecodedData ?? "null"} " +
                    $"grade={linearOverallGrade ?? "null"} " +
                    $"formal={linearFormalGrade ?? "null"} " +
                    $"ap={linearAperture?.ToString() ?? "null"} " +
                    $"wl={linearWavelength?.ToString() ?? "null"} " +
                    $"lighting={linearLighting ?? "null"}");
            }

            // ── Step 4c: Data Format Check table (DM TC HTML) ─────────────────
            //
            // The DM TC HTML contains a dedicated DFC table identified by a <th>
            // reading "Data Format Check".  The device has already validated the
            // GS1 AIs — scraping this is more reliable than re-parsing the push
            // XML decoded data string (which BarcodeDataFormatter may have altered).
            //
            // Table structure (minified, single line):
            //   <th>Data Format Check</th>
            //   <th>GS1 Application Data Format: PASS</th>
            //   <tr><td><strong>Name</strong></td>…</tr>   ← skip (header)
            //   <tr><td>GS1 Header</td><td>&lt;F1&gt;</td><td>PASS</td></tr>
            //   …
            DataFormatCheckResult? scrapedDfc = null;
            string? htmlApplicationStandard = null;
            {
                // Split on </table> to isolate the DFC table without nested-table risk.
                string[] tableParts = htmlContent.Split("</table>", StringSplitOptions.None);
                foreach (string part in tableParts)
                {
                    if (!part.Contains("Data Format Check", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Overall verdict from the second <th>: "…: PASS" or "…: FAIL".
                    string? dfcStandard = null;
                    var dfcOverall = OverallPassFail.Pass;
                    foreach (Match thm in Regex.Matches(part,
                        @"<th[^>]*>(.*?)</th>", RegexOptions.Singleline | RegexOptions.IgnoreCase))
                    {
                        string th = WebUtility.HtmlDecode(
                            Regex.Replace(thm.Groups[1].Value, "<[^>]+>", "").Trim());
                        if (th.Contains("Data Format Check", StringComparison.OrdinalIgnoreCase))
                            continue;
                        int colon = th.LastIndexOf(':');
                        if (colon > 0)
                        {
                            dfcStandard = th[..colon].Trim();
                            htmlApplicationStandard = dfcStandard;
                            if (th[(colon + 1)..].Trim()
                                    .Equals("FAIL", StringComparison.OrdinalIgnoreCase))
                                dfcOverall = OverallPassFail.Fail;
                        }
                        break;
                    }

                    // Data rows: <tr> containing exactly 3 <td> cells, no <th> or <strong>.
                    var dfcRows = new List<DataFormatCheckRow>();
                    foreach (Match trm in Regex.Matches(part,
                        @"<tr>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase))
                    {
                        string row = trm.Groups[1].Value;
                        if (row.Contains("<th",     StringComparison.OrdinalIgnoreCase)) continue;
                        if (row.Contains("<strong", StringComparison.OrdinalIgnoreCase)) continue;
                        var tds = Regex.Matches(row,
                            @"<td[^>]*>(.*?)</td>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                        if (tds.Count < 3) continue;
                        string rName  = WebUtility.HtmlDecode(
                            Regex.Replace(tds[0].Groups[1].Value, "<[^>]+>", "").Trim());
                        string rData  = WebUtility.HtmlDecode(
                            Regex.Replace(tds[1].Groups[1].Value, "<[^>]+>", "").Trim());
                        string rCheck = WebUtility.HtmlDecode(
                            Regex.Replace(tds[2].Groups[1].Value, "<[^>]+>", "").Trim());
                        if (!string.IsNullOrEmpty(rName))
                            dfcRows.Add(new DataFormatCheckRow
                                { Name = rName, Data = rData, Check = rCheck });
                    }

                    if (dfcRows.Count > 0)
                        scrapedDfc = new DataFormatCheckResult
                        {
                            Overall  = dfcOverall,
                            Standard = dfcStandard ?? "GS1 Application Data Format",
                            Rows     = dfcRows,
                        };
                    break;
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] DFC: {(scrapedDfc is null ? "not found" : $"{scrapedDfc.Rows.Count} rows, {scrapedDfc.Overall}")}");
            }

            // ── Step 4d: Verified datetime string from HTML header ────────────
            //
            // The TruCheck HTML header contains a <p> element with the text:
            //   "Verified: Tue 18-Aug-2026 05:10:32(520ms) PM"
            // This is the local Eastern time (device clock = America/New_York) and
            // is used verbatim as {{REPORT_DATETIME}} in the VCCS PDF so the displayed
            // time matches the TruCheck report exactly — including the (ms) fragment.
            //
            // Note: the in-page <td>-based DateTime header cells show Unix epoch and
            // are intentionally ignored (see ScanDateTime parsing below from filename).
            string? htmlVerifiedString = null;
            {
                var verifiedMatch = Regex.Match(htmlContent,
                    @"Verified:\s*([^<\r\n]+)", RegexOptions.IgnoreCase);
                if (verifiedMatch.Success)
                    htmlVerifiedString = WebUtility.HtmlDecode(
                        verifiedMatch.Groups[1].Value.Trim());
            }

            // ── Step 4e: Verification Grades row — verbatim display strings ──────
            //
            // The TruCheck HTML Verification Grades table has a header row:
            //   Standard | Grade | Aperture | Wavelength | Lighting | Formal Grade
            // followed immediately by the data row (single-mode DM example):
            //   ISO 15415:2024 | 4.0 (A) | 16 | 660 | 45Q | 4.0/16/660/45Q
            //
            // Strategy: FindLastIndex("Formal Grade") in the flat cell list (handles
            // multi-mode pages where two grade sections appear — the 2D section is last).
            // Then read the next 6 positions verbatim.  No parsing, no reformatting.
            string? htmlStandard            = null;
            string? htmlOverallGradeDisplay = null;
            string? htmlAperture            = null;
            string? htmlWavelength          = null;
            string? htmlLighting            = null;
            string? htmlFormalGrade         = null;
            string? htmlLinearStandard      = null;
            string? htmlLinearGradeDisplay  = null;
            string? htmlLinearAperture      = null;
            string? htmlLinearWavelength    = null;
            string? htmlLinearLighting      = null;
            string? htmlLinearFormalGrade   = null;
            {
                static string? E(string s) =>
                    string.IsNullOrWhiteSpace(s) ? null : s.Trim();

                int fgIdx = cells.FindLastIndex(
                    c => c.Equals("Formal Grade", StringComparison.OrdinalIgnoreCase));
                if (fgIdx >= 0 && fgIdx + 6 < cells.Count)
                {
                    htmlStandard            = E(cells[fgIdx + 1]);
                    htmlOverallGradeDisplay = E(cells[fgIdx + 2]);
                    htmlAperture            = E(cells[fgIdx + 3]);
                    htmlWavelength          = E(cells[fgIdx + 4]);
                    htmlLighting            = E(cells[fgIdx + 5]);
                    htmlFormalGrade         = E(cells[fgIdx + 6]);

                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SCRAPER] Grade row: std={htmlStandard ?? "null"} " +
                        $"grade={htmlOverallGradeDisplay ?? "null"} " +
                        $"ap={htmlAperture ?? "null"} wl={htmlWavelength ?? "null"} " +
                        $"light={htmlLighting ?? "null"} formal={htmlFormalGrade ?? "null"}");
                }

                if (isMultiMode)
                {
                    int linearFgIdx = cells.FindIndex(
                        c => c.Equals("Formal Grade", StringComparison.OrdinalIgnoreCase));
                    if (linearFgIdx >= 0 && linearFgIdx + 6 < cells.Count &&
                        linearFgIdx != fgIdx)
                    {
                        htmlLinearStandard     = E(cells[linearFgIdx + 1]);
                        htmlLinearGradeDisplay = E(cells[linearFgIdx + 2]);
                        htmlLinearAperture     = E(cells[linearFgIdx + 3]);
                        htmlLinearWavelength   = E(cells[linearFgIdx + 4]);
                        htmlLinearLighting     = E(cells[linearFgIdx + 5]);
                        htmlLinearFormalGrade  = E(cells[linearFgIdx + 6]);
                    }
                }
            }

            string? embeddedBarcodeImage = null;
            {
                // Only accept an image that is embedded in the HTML artifact itself.
                // Images carried by the reader push payload are intentionally not
                // promoted to TruCheck provenance.
                Match imageMatch = Regex.Match(
                    htmlContent,
                    @"<img\b[^>]*\bsrc\s*=\s*[""']data:image/(?:jpeg|jpg|png);base64,([^""']+)[""']",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (imageMatch.Success)
                    embeddedBarcodeImage = imageMatch.Groups[1].Value.Trim();
            }

            // ── Step 5: DateTime from filename ────────────────────────────────
            //
            // DMST names reports either "yyyy-MM-dd_HH-mm-ss-mmm_<random>.html"
            // or "_F1_<encoded-data>_yyyy-MM-dd_HH-mm-ss-mmm.html". The latter is
            // the normal GS1 DataMatrix form, so the timestamp cannot be assumed to
            // begin at character zero. The device clock is already local
            // (America/New_York via DMCC DEVICE.TIMEZONE); parse it as-is.
            DateTime? scanDateTime = null;
            var fn = Path.GetFileNameWithoutExtension(sourcePath);
            Match fileTimestamp = Regex.Match(
                fn, @"(?<!\d)(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})(?!\d)");
            if (fileTimestamp.Success && DateTime.TryParseExact(
                    fileTimestamp.Groups[1].Value, "yyyy-MM-dd_HH-mm-ss",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                scanDateTime = dt;

            var report = new DmstHtmlReport
            {
                ScanDateTime          = scanDateTime,
                SourceFilePath        = sourcePath,
                HtmlVerifiedString    = htmlVerifiedString,
                HtmlSourceFileName    = hasSyntheticSourcePath
                                          ? null
                                          : Path.GetFileName(sourcePath.Replace('\\', '/')),
                HasSyntheticSourcePath = hasSyntheticSourcePath,
                // A timestamp-less filename is still a valid report when the HTML
                // always-present Verified: header was scraped successfully.
                ParseSucceeded        = scanDateTime.HasValue ||
                                         !string.IsNullOrWhiteSpace(htmlVerifiedString),

                // ── Verbatim Verification Grades row ──────────────────────────
                HtmlStandard            = htmlStandard,
                HtmlOverallGradeDisplay = htmlOverallGradeDisplay,
                HtmlAperture            = htmlAperture,
                HtmlWavelength          = htmlWavelength,
                HtmlLighting            = htmlLighting,
                HtmlFormalGrade         = htmlFormalGrade,
                HtmlSymbology           = GetAny("Symbology", "Symbol Type"),
                HtmlDecodedData          = GetAny("Data", "Encoded Data", "Decoded Data"),
                HtmlApplicationStandard = htmlApplicationStandard ?? Get("Application Standard"),
                HtmlLinearStandard      = htmlLinearStandard,
                HtmlLinearGradeDisplay  = htmlLinearGradeDisplay,
                HtmlLinearAperture      = htmlLinearAperture,
                HtmlLinearWavelength    = htmlLinearWavelength,
                HtmlLinearLighting      = htmlLinearLighting,
                HtmlLinearFormalGrade   = htmlLinearFormalGrade,
                HtmlBarcodeImageBase64  = embeddedBarcodeImage,

                // ── Supplemental fields: not accessible via push XML on fw 6.1.16_sr4 ──
                ECLevel         = Get("Error Correction Level"),   // "M"
                DataMaskPattern = Get("Data Mask Pattern"),         // "2"
                ECI             = Get("ECI"),                       // "000003"
                ImagePolarity   = Get("Image"),                     // "Black on white"

                // ── Bonus: present in HTML, empty/wrong in push XML ────────────
                DataCodewords         = GetInt("Data Codewords"),
                ErrorCorrectionBudget = GetInt("Error Correction Budget"),
                EncodedCharacters     = GetInt("Encoded characters"),
                ErrorsCorrected       = GetInt("Errors Corrected"),
                ErrorCapacityUsed     = GetInt("Error Capacity Used"),
                TotalCodewords        = GetInt("Total Codewords"),

                // ── Cross-validation: header-derived ──────────────────────────
                OverallGrade    = overallGrade,

                // ── Cross-validation: simple characteristics table ─────────────
                // QR HTML: label "QR Size" → "29x29".  DM HTML: label "Matrix Size"
                // → "16x36 (Data: 14x34)" — strip the " (Data:…)" suffix.
                MatrixSize    = Get("QR Size") ?? StripDataSuffix(Get("Matrix Size")),
                NominalXDim   = Get("Nominal X Dim"),
                HorizontalBWG = GetDecimal("Horizontal BWG", stripPercent: true),
                VerticalBWG   = GetDecimal("Vertical BWG",   stripPercent: true),

                // ── Cross-validation: grade parameters table ───────────────────
                UECPercent = GetGradePct("1. Unused Error Correction (UEC)"),
                SCPercent  = GetGradePct("2. Symbol Contrast (SC)"),       // null on IMAGE.LOAD
                ANUPercent = GetGradePct("4. Axial Nonuniformity (ANU)"),
                GNUPercent = GetGradePct("5. Grid Nonuniformity (GNU)"),

                // ── GS1 Data Format Check (scraped from DM TC HTML) ───────────
                ScrapedDataFormatCheck   = scrapedDfc,

                // ── Multi-mode linear symbol ───────────────────────────────────
                IsMultiMode              = isMultiMode,
                LinearSymbology          = linearSymbology,
                LinearDecodedData        = linearDecodedData,
                LinearOverallGrade       = linearOverallGrade,
                LinearOverallGradeNumeric = linearOverallGradeNumeric,
                LinearFormalGrade        = linearFormalGrade,
                LinearAperture           = linearAperture,
                LinearWavelength         = linearWavelength,
                LinearLighting           = linearLighting,
                LinearStandard           = linearStandard,
            };

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] ParseHtml: ECLevel={report.ECLevel ?? "null"} " +
                $"DataMask={report.DataMaskPattern ?? "null"} " +
                $"ECI={report.ECI ?? "null"} " +
                $"Polarity={report.ImagePolarity ?? "null"} " +
                $"DataCW={report.DataCodewords?.ToString() ?? "null"} " +
                $"ECBudget={report.ErrorCorrectionBudget?.ToString() ?? "null"} " +
                $"EncodedChars={report.EncodedCharacters?.ToString() ?? "null"}" +
                (isMultiMode ? $" MultiMode=true LinearSymb={linearSymbology}" : " MultiMode=false"));

            return report;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] ParseHtml exception for '{Path.GetFileName(sourcePath)}': {ex.Message}");
            return new DmstHtmlReport
            {
                SourceFilePath = sourcePath,
                ParseSucceeded = false,
                ParseError     = ex.Message,
            };
        }
    }

    // ── IDisposable ───────────────────────────────────────────────────────────

    public void Dispose() => Stop();

    // ── Internal types ────────────────────────────────────────────────────────

    private sealed record PendingHtmlReport(DmstHtmlReport Report);
}
