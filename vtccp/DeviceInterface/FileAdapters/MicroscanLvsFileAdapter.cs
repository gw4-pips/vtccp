namespace DeviceInterface.FileAdapters;

using ExcelEngine.Models;

/// <summary>
/// File-export adapter for Omron Microscan LVS-95XX series verifiers
/// (LVS-9510, LVS-7510, etc.).
///
/// Integration path: LVS-95XX software writes a result file to a configurable
/// output folder after each scan.  This adapter watches that folder, reads each
/// new file, and fires <see cref="RecordParsed"/> with a <see cref="VerificationRecord"/>
/// whose <see cref="VerificationRecord.VerifierBrand"/> is always <c>"OMRON/LVS"</c> —
/// ensuring the PDF Device header row shows the correct brand label rather than "—".
///
/// <b>Partial implementation</b>: the LVS-95XX export file format has not yet been
/// confirmed.  <see cref="ParseFileAsync"/> reads the raw file content and builds a
/// record using <see cref="BuildRecord"/>; fields that require format-specific parsing
/// are left null until the format is confirmed and the parse logic is completed
/// (tracked in a separate task).
///
/// Every record returned by <see cref="BuildRecord"/> — and therefore by this adapter —
/// carries <c>VerifierBrand = "OMRON/LVS"</c> regardless of how many fields are populated.
///
/// <b>Branding note</b>: "LVS" is the Microscan product line acquired by Omron ~2018.
/// The label "OMRON/LVS" matches the "LVS" / "Omron" / "Microscan" entries in
/// <c>PdfReportGenerator.BrandPatterns</c>.
/// </summary>
public sealed class MicroscanLvsFileAdapter : IDisposable
{
    /// <summary>
    /// All-caps brand label used by this adapter.
    /// Matches the "LVS", "Omron", and "Microscan" entries in
    /// <c>PdfReportGenerator.BrandPatterns</c>.
    /// </summary>
    public const string Brand = "OMRON/LVS";

    // ── Watch folder ──────────────────────────────────────────────────────────
    private readonly string _watchFolder;
    private FileSystemWatcher? _watcher;
    private bool _disposed;

    /// <summary>
    /// Raised each time a new export file is read and converted to a record.
    /// </summary>
    public event EventHandler<VerificationRecord>? RecordParsed;

    /// <param name="watchFolder">
    /// Absolute path to the folder where LVS-95XX software writes its result files.
    /// Configure this path in the LVS software auto-save / export settings.
    /// </param>
    public MicroscanLvsFileAdapter(string watchFolder)
    {
        if (string.IsNullOrWhiteSpace(watchFolder))
            throw new ArgumentException("Watch folder path must not be empty.", nameof(watchFolder));

        _watchFolder = watchFolder;
    }

    /// <summary>
    /// Starts the folder watcher.  Newly created files trigger <see cref="RecordParsed"/>.
    /// </summary>
    public void Start()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        Directory.CreateDirectory(_watchFolder);

        _watcher = new FileSystemWatcher(_watchFolder)
        {
            // TODO: narrow the filter once the export file extension is confirmed.
            Filter              = "*.*",
            NotifyFilter        = NotifyFilters.FileName,
            EnableRaisingEvents = true,
        };

        _watcher.Created += OnFileCreated;

        System.Diagnostics.Debug.WriteLine(
            $"[LVS] Watching for export files in: {_watchFolder}");
    }

    // ── File arrival ──────────────────────────────────────────────────────────

    private async void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        try
        {
            // Small delay to allow the writing application to finish the file.
            await Task.Delay(500).ConfigureAwait(false);

            VerificationRecord record = await ParseFileAsync(e.FullPath).ConfigureAwait(false);
            RecordParsed?.Invoke(this, record);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[LVS] OnFileCreated error for '{e.Name}': {ex.GetType().Name}: {ex.Message}");
        }
    }

    /// <summary>
    /// Reads an LVS-95XX export file and returns a <see cref="VerificationRecord"/>.
    ///
    /// Always sets <c>VerifierBrand = <see cref="Brand"/></c> so the PDF Device header
    /// row shows "OMRON/LVS" instead of "—".
    ///
    /// <b>TODO</b>: parse LVS-specific fields (Symbology, DecodedData, grades, etc.)
    /// once the export file format is confirmed.  Until then, those fields are null
    /// and <see cref="BuildRecord"/> is called with the raw file text.
    /// </summary>
    internal static async Task<VerificationRecord> ParseFileAsync(string filePath)
    {
        string? rawContent = null;
        DateTime timestamp = File.Exists(filePath)
            ? File.GetLastWriteTime(filePath)
            : DateTime.Now;

        try
        {
            rawContent = await File.ReadAllTextAsync(filePath).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[LVS] ParseFileAsync could not read '{Path.GetFileName(filePath)}': " +
                $"{ex.GetType().Name}: {ex.Message}");
        }

        return BuildRecord(rawContent, timestamp);
    }

    /// <summary>
    /// Constructs a <see cref="VerificationRecord"/> from LVS-95XX export file content.
    ///
    /// <b>VerifierBrand is always <c>"OMRON/LVS"</c></b> on every record this method
    /// returns — this is the key guarantee this adapter provides.  All other fields
    /// are populated when the export format is known; they are null in the interim.
    /// </summary>
    /// <param name="rawContent">
    /// Raw text content of the LVS result file.  May be <see langword="null"/>
    /// if the file could not be read; the record is still returned with brand set.
    /// </param>
    /// <param name="timestamp">Verification timestamp; defaults to <see cref="DateTime.Now"/>.</param>
    public static VerificationRecord BuildRecord(
        string?  rawContent = null,
        DateTime? timestamp  = null)
    {
        // TODO: extract Symbology, DecodedData, grades, standard, aperture, wavelength,
        // etc. from rawContent once the LVS-95XX export file format is confirmed.
        //
        // Example parse skeleton:
        //   string? symbology   = ExtractField(rawContent, "Symbology");
        //   string? decodedData = ExtractField(rawContent, "DecodedData");
        //   ...

        return new VerificationRecord
        {
            // VerifierBrand is always set — this is the fix for the PDF header.
            VerifierBrand           = Brand,

            // Timestamp: prefer the file's write time; fall back to now.
            VerificationDateTime    = timestamp ?? DateTime.Now,

            // Symbology is required (non-nullable on the record) — placeholder until parsing.
            Symbology               = "Unknown",

            // All other fields are null pending format-specific parse implementation.
            // They will be populated here once the LVS-95XX export format is confirmed.
        };
    }

    // ── IDisposable ───────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_watcher is not null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Created -= OnFileCreated;
            _watcher.Dispose();
        }
    }
}
