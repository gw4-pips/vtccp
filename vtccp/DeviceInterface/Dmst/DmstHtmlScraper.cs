namespace DeviceInterface.Dmst;

using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
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

    // ── Diagnostic capture ────────────────────────────────────────────────────

    /// <summary>
    /// When true, the first HTML report received is copied to
    /// <see cref="DiagnosticCapturePath"/> before being deleted.
    /// Set to true temporarily to capture an HTML sample for parser diagnostics.
    /// ParseHtml() is fully implemented and validated against the 2026-05-25 live sample.
    /// </summary>
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

                // ── Diagnostic capture (first sample only) ──────────────────
                // Preserves the raw HTML for parser implementation.
                // Disable DiagnosticCaptureEnabled once ParseHtml is implemented.
                if (DiagnosticCaptureEnabled && !_diagnosticCaptured)
                {
                    _diagnosticCaptured = true;
                    try
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(DiagnosticCapturePath)!);
                        File.Copy(e.FullPath, DiagnosticCapturePath, overwrite: true);
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-SCRAPER] Diagnostic copy saved → '{DiagnosticCapturePath}'");
                    }
                    catch (Exception copyEx)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-SCRAPER] Diagnostic copy failed: {copyEx.Message}");
                    }
                }

                // Delete the transient DMST output — data is held in memory.
                File.Delete(e.FullPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SCRAPER] Error reading/parsing '{e.FullPath}': {ex.Message}");
            }
        });
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
    internal static DmstHtmlReport ParseHtml(string htmlContent, string sourcePath)
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

            // ── Step 4: overall grade letter from header "D.D (L)" pattern ───
            //
            // Located near cell 19 in the header table; search a window rather
            // than hardcoding the index to absorb minor layout variation.
            string? overallGrade = null;
            var gradeDisplay = cells.Skip(14).Take(12)
                .FirstOrDefault(c => Regex.IsMatch(c, @"^\d+\.\d+\s*\([A-Fa-f]\)$"));
            if (gradeDisplay is not null)
            {
                var m = Regex.Match(gradeDisplay, @"\(([A-Fa-f])\)");
                if (m.Success) overallGrade = m.Groups[1].Value.ToUpperInvariant();
            }

            // ── Step 5: DateTime from filename ────────────────────────────────
            //
            // Format: "yyyy-MM-dd_HH-mm-ss-mmm_<random>.html"
            // The first 19 chars "yyyy-MM-dd_HH-mm-ss" are always present.
            DateTime? scanDateTime = null;
            var fn = Path.GetFileNameWithoutExtension(sourcePath);
            if (fn.Length >= 19 && DateTime.TryParseExact(
                    fn[..19], "yyyy-MM-dd_HH-mm-ss",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                scanDateTime = dt;

            var report = new DmstHtmlReport
            {
                ScanDateTime          = scanDateTime,
                SourceFilePath        = sourcePath,
                ParseSucceeded        = scanDateTime.HasValue,

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
            };

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] ParseHtml: ECLevel={report.ECLevel ?? "null"} " +
                $"DataMask={report.DataMaskPattern ?? "null"} " +
                $"ECI={report.ECI ?? "null"} " +
                $"Polarity={report.ImagePolarity ?? "null"} " +
                $"DataCW={report.DataCodewords?.ToString() ?? "null"} " +
                $"ECBudget={report.ErrorCorrectionBudget?.ToString() ?? "null"} " +
                $"EncodedChars={report.EncodedCharacters?.ToString() ?? "null"}");

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
