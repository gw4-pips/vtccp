namespace ExcelEngine.Session;

using System.Text.Json;
using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Utilities;
using ExcelEngine.Writer;

/// <summary>
/// Controls the VTCCP job session lifecycle — New Job / Open Job / Close Job.
///
/// Roll identifier modes (see <see cref="RollIncrementMode"/>):
///   Manual        — roll changes only on explicit operator action; caller supplies value.
///   AutoIncrement — starts at RollStartValue; increments by 1 on each SetNewOperatorAndRoll().
///   DateTimeStamp — yyyyMMddHHmmss generated at session open and on SetNewOperatorAndRoll().
///
/// PRODUCT DECISION — roll does NOT auto-increment on StartSession (confirmed 2026-03-26):
///   The operator must explicitly call SetNewOperatorAndRoll() to advance the roll value.
///   A new session opens without a roll change; this matches the DMV TruCheck "Set New
///   Operator/Roll" UI paradigm.  AutoIncrement and DateTimeStamp are VTCCP conveniences
///   that preserve the same contract: no automatic roll change on session open
///   (AutoIncrement resets to RollStartValue for new sessions; resumes restore from sidecar).
///
/// File path stability guarantee:
///   _outputPath / _sidecarPath are resolved ONCE at StartSession() and never
///   re-derived mid-session, so no state change can orphan sidecars or open the wrong file.
///
/// StartSession() path algorithm:
///   A. Try to resume from the initial path (before sidecar merge changes naming fields).
///      If a sidecar + xlsx pair is found → merge sidecar → use initial path as stable path.
///   B. If no sidecar found yet:
///      1. Apply roll-mode initialisation on the new-session state (so {Roll} is correct
///         before filename resolution).
///      2. Re-resolve the path from the now-complete state.
///      3. Try second-chance sidecar lookup at the new path.
///      4. Pin whichever path has the sidecar+file pair; if neither, open at the new path.
///
/// This ordering guarantees that:
///   - On resume, we open the file that actually contains the interrupted data (never fork).
///   - On new session, the roll label is applied before the filename is built.
///
/// Usage:
///   var mgr = new SessionManager(schema);
///   mgr.StartSession(state);                     // opens / creates file
///   mgr.AddRecord(record);                       // appends record; updates sidecar
///   mgr.SetNewOperatorAndRoll("OP2");            // Auto/DateTime: auto-advance roll
///   mgr.SetNewOperatorAndRoll("OP2", roll: 5);  // Manual: caller supplies new value
///   mgr.CloseSession();                          // saves file; deletes sidecar
/// </summary>
public sealed class SessionManager : IDisposable
{
    private readonly ColumnSchema _schema;
    private SessionState? _currentSession;
    private IExcelAdapter? _adapter;
    private ExcelWriter? _writer;

    // Pinned at StartSession(); never re-derived during the session.
    private string? _outputPath;
    private string? _sidecarPath;

    private bool _disposed;

    public SessionState? CurrentSession => _currentSession;
    public bool IsSessionOpen  => _currentSession is not null && _writer is not null;
    public int  RecordsWritten => _currentSession?.RecordCount ?? 0;

    /// <summary>Pinned output file path for the current session, or null if closed.</summary>
    public string? OutputPath => _outputPath;

    public SessionManager(ColumnSchema schema)
    {
        _schema = schema;
    }

    // ── Session lifecycle ─────────────────────────────────────────────────────

    /// <summary>
    /// Start a new job session or resume an interrupted one.
    ///
    /// Path pinning guarantee: the path used to open the file is always where the
    /// sidecar+xlsx pair was found (for resumes) or the fully-resolved new-session
    /// path (for fresh starts). It never changes after this call returns.
    ///
    /// Roll mode (new sessions only; resumes restore roll from sidecar):
    ///   Manual        — RollNumber stays at caller-supplied value.
    ///   AutoIncrement — RollNumber set to RollStartValue.
    ///   DateTimeStamp — RollTimestamp set to yyyyMMddHHmmss.
    ///
    /// Returns the pinned absolute output file path.
    /// </summary>
    public string StartSession(SessionState state)
    {
        if (_writer is not null)
            throw new InvalidOperationException(
                "A session is already open. Call CloseSession() first.");

        _currentSession = state;

        // ── Step A: try resume from the initial (caller-supplied) path ─────────
        string initialPath    = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
        string initialSidecar = GetSidecarPath(initialPath);

        bool resumed = TryMergeSidecar(initialSidecar, state);
        if (resumed)
        {
            // The xlsx is at initialPath; pin it directly.
            // Do NOT re-resolve — the sidecar may have added naming fields that would
            // produce a different path, but the physical file is the one we found.
            _outputPath  = initialPath;
            _sidecarPath = initialSidecar;
        }
        else
        {
            // ── Step B: new session ───────────────────────────────────────────────
            // B1: Apply roll-mode initialisation BEFORE filename resolution so
            //     {Roll} tokens in custom patterns get the effective roll value.
            ApplyRollModeOnStart(state);

            // B2: Re-resolve with now-complete state.
            string newPath    = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
            string newSidecar = GetSidecarPath(newPath);

            // B3: Second-chance resume if path changed (e.g. roll mode affected filename).
            if (newPath != initialPath)
                resumed = TryMergeSidecar(newSidecar, state);

            _outputPath  = newPath;
            _sidecarPath = newSidecar;
        }

        // ── Open / create the Excel file ─────────────────────────────────────
        _adapter = state.OutputFormat == OutputFormat.Xls
            ? new XlsAdapter()
            : new XlsxAdapter();

        _writer = new ExcelWriter(_adapter, _schema, state);
        _writer.Open(_outputPath);

        // ── Persist sidecar immediately ───────────────────────────────────────
        SaveSidecar(_sidecarPath!, state);

        return _outputPath;
    }

    /// <summary>
    /// Append a single VerificationRecord to the active session's output file.
    /// The sidecar is updated after every write for crash safety.
    /// </summary>
    public void AddRecord(VerificationRecord record)
    {
        EnsureOpen();

        // Resolve effective batch number for this record.
        string? batchOverride = null;
        if (_currentSession!.BatchMode == BatchMode.AutoFromGS1)
        {
            string? extracted = AutoBatchExtractor.ExtractBatchLot(record.DecodedData);
            if (extracted is not null)
                batchOverride = extracted;
            // If AI(10) absent (e.g. non-GS1 symbol in a mixed session),
            // batchOverride stays null → mapper uses record.BatchNumber as-is.
        }

        _writer!.AppendRecord(record, batchOverride);
        _currentSession!.RecordCount++;
        SaveSidecar(_sidecarPath!, _currentSession!);
    }

    /// <summary>
    /// Save and close the active session. Removes the sidecar (job complete).
    /// </summary>
    public void CloseSession()
    {
        if (_writer is null) return;

        _writer.Save();
        _writer.Dispose();
        _writer = null;

        _adapter?.Dispose();
        _adapter = null;

        if (_sidecarPath is not null && File.Exists(_sidecarPath))
            File.Delete(_sidecarPath);

        _currentSession = null;
        _outputPath     = null;
        _sidecarPath    = null;
    }

    /// <summary>
    /// Change the active operator and advance the roll identifier.
    ///
    /// Behaviour by mode:
    ///   <c>Manual</c>        — <paramref name="manualRoll"/> sets RollNumber (required for a
    ///                          meaningful change; omit to keep current value unchanged).
    ///   <c>AutoIncrement</c> — RollNumber is incremented by 1.
    ///   <c>DateTimeStamp</c> — a new yyyyMMddHHmmss timestamp is generated.
    ///
    /// The output file path does not change; mid-session operator/roll changes do not
    /// rename the already-open file.
    /// </summary>
    public void SetNewOperatorAndRoll(string operatorId, int? manualRoll = null)
    {
        EnsureOpen();
        _currentSession!.OperatorId = operatorId;

        switch (_currentSession!.RollIncrementMode)
        {
            case RollIncrementMode.Manual:
                if (manualRoll.HasValue)
                    _currentSession!.RollNumber = manualRoll.Value;
                break;

            case RollIncrementMode.AutoIncrement:
                _currentSession!.RollNumber++;
                break;

            case RollIncrementMode.DateTimeStamp:
                _currentSession!.RollTimestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                break;
        }

        SaveSidecar(_sidecarPath!, _currentSession!);
    }

    /// <summary>
    /// Returns the sidecar JSON path for a given output file path.
    /// Convention: {outputFile}.vtccp.json
    /// </summary>
    public static string GetSidecarPath(string outputFilePath)
        => outputFilePath + ".vtccp.json";

    // ── Roll mode helpers ─────────────────────────────────────────────────────

    private static void ApplyRollModeOnStart(SessionState state)
    {
        switch (state.RollIncrementMode)
        {
            case RollIncrementMode.AutoIncrement:
                state.RollNumber = state.RollStartValue;
                break;

            case RollIncrementMode.DateTimeStamp:
                state.RollTimestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                break;

            case RollIncrementMode.Manual:
            default:
                // No automatic change; caller controls RollNumber.
                break;
        }
    }

    // ── Sidecar serialisation ─────────────────────────────────────────────────

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        WriteIndented = true,
    };

    private static void SaveSidecar(string path, SessionState state)
    {
        var sidecar = new SessionSidecar
        {
            JobName     = state.JobName,
            OperatorId  = state.OperatorId,
            BatchNumber = state.BatchNumber,
            CompanyName = state.CompanyName,
            ProductName = state.ProductName,
            CustomNote  = state.CustomNote,
            User1       = state.User1,
            User2       = state.User2,

            BatchMode         = state.BatchMode.ToString(),
            RollIncrementMode = state.RollIncrementMode.ToString(),
            RollNumber        = state.RollNumber,
            RollStartValue    = state.RollStartValue,
            RollTimestamp     = state.RollTimestamp,

            DeviceSerial    = state.DeviceSerial,
            DeviceName      = state.DeviceName,
            FirmwareVersion = state.FirmwareVersion,
            CalibrationDate = state.CalibrationDate,

            OutputFormat    = state.OutputFormat.ToString(),
            OutputDirectory = state.OutputDirectory,
            FileNamePattern = state.FileNamePattern,

            SessionStarted = state.SessionStarted,
            RecordCount    = state.RecordCount,
        };
        File.WriteAllText(path, JsonSerializer.Serialize(sidecar, _jsonOpts));
    }

    private static SessionSidecar? LoadSidecar(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<SessionSidecar>(json);
    }

    /// <summary>
    /// Load the sidecar at <paramref name="sidecarPath"/> and merge it into state.
    /// Both the xlsx and the sidecar must exist for a valid resume.
    /// Returns true if a sidecar was found and applied.
    /// </summary>
    private static bool TryMergeSidecar(string sidecarPath, SessionState state)
    {
        string outputPath = sidecarPath[..^".vtccp.json".Length];
        if (!File.Exists(sidecarPath) || !File.Exists(outputPath))
            return false;

        try
        {
            var saved = LoadSidecar(sidecarPath);
            if (saved is not null)
            {
                ApplySidecarToState(saved, state);
                return true;
            }
        }
        catch (JsonException)
        {
            // Corrupt sidecar — ignore and start fresh.
        }
        return false;
    }

    /// <summary>
    /// Restore saved fields from the sidecar onto the live session state.
    ///
    /// Roll state and counters are always authoritative from the sidecar.
    /// Context fields restore only if the caller left them null, allowing
    /// callers to pre-set overrides before calling StartSession().
    /// </summary>
    private static void ApplySidecarToState(SessionSidecar saved, SessionState state)
    {
        // Roll identifier — always authoritative
        state.RollNumber     = saved.RollNumber;
        state.RollStartValue = saved.RollStartValue;
        state.RollTimestamp  = saved.RollTimestamp;
        if (saved.BatchMode is not null &&
            Enum.TryParse<BatchMode>(saved.BatchMode, out var bm))
            state.BatchMode = bm;
        if (saved.RollIncrementMode is not null &&
            Enum.TryParse<RollIncrementMode>(saved.RollIncrementMode, out var mode))
            state.RollIncrementMode = mode;

        // Session counters — always authoritative
        state.RecordCount    = saved.RecordCount;
        state.SessionStarted = saved.SessionStarted;

        // Context fields — restore only if caller left them null
        if (state.JobName         is null && saved.JobName         is not null) state.JobName         = saved.JobName;
        if (state.OperatorId      is null && saved.OperatorId      is not null) state.OperatorId      = saved.OperatorId;
        if (state.BatchNumber     is null && saved.BatchNumber     is not null) state.BatchNumber     = saved.BatchNumber;
        if (state.CompanyName     is null && saved.CompanyName     is not null) state.CompanyName     = saved.CompanyName;
        if (state.ProductName     is null && saved.ProductName     is not null) state.ProductName     = saved.ProductName;
        if (state.CustomNote      is null && saved.CustomNote      is not null) state.CustomNote      = saved.CustomNote;
        if (state.User1           is null && saved.User1           is not null) state.User1           = saved.User1;
        if (state.User2           is null && saved.User2           is not null) state.User2           = saved.User2;
        if (state.DeviceSerial    is null && saved.DeviceSerial    is not null) state.DeviceSerial    = saved.DeviceSerial;
        if (state.DeviceName      is null && saved.DeviceName      is not null) state.DeviceName      = saved.DeviceName;
        if (state.FirmwareVersion is null && saved.FirmwareVersion is not null) state.FirmwareVersion = saved.FirmwareVersion;
        if (state.CalibrationDate is null && saved.CalibrationDate is not null) state.CalibrationDate = saved.CalibrationDate;
        if (state.OutputDirectory is null && saved.OutputDirectory is not null) state.OutputDirectory = saved.OutputDirectory;
        if (state.FileNamePattern is null && saved.FileNamePattern is not null) state.FileNamePattern = saved.FileNamePattern;
        if (saved.OutputFormat    is not null &&
            Enum.TryParse<OutputFormat>(saved.OutputFormat, out var fmt))
            state.OutputFormat = fmt;
    }

    // ── IDisposable ───────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _writer?.Dispose();
        _adapter?.Dispose();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void EnsureOpen()
    {
        if (_writer is null)
            throw new InvalidOperationException(
                "No session is open. Call StartSession() first.");
    }
}
