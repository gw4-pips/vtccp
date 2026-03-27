namespace ExcelEngine.Session;

using System.Text.Json;
using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Writer;

/// <summary>
/// Controls the VTCCP job session lifecycle, replicating the behaviour of
/// Webscan TruCheck's "Excel Functions" panel (New Job / Open Job / Close Job).
///
/// Roll identifier modes (see <see cref="RollIncrementMode"/>):
///   Manual      — roll changes only on explicit operator action; caller supplies value.
///   AutoIncrement — roll starts at RollStartValue and increments by 1 on each
///                  SetNewOperatorAndRoll() call.
///   DateTimeStamp — roll is a yyyyMMddHHmmss timestamp; generated at session open and
///                  on each SetNewOperatorAndRoll() call.
///
/// File path stability guarantee:
///   _outputPath / _sidecarPath are resolved ONCE at StartSession() and never
///   re-derived mid-session, so no state changes can orphan sidecars or target
///   a different file.
///
/// Resume path consistency (5-step):
///   1. Resolve initial path from caller-supplied state.
///   2. Load and merge sidecar from initial path.
///   3. Re-resolve path from merged state.
///   4. Second-chance sidecar lookup at new path if path changed.
///   5. Pin final path and open file.
///
/// Usage:
///   var mgr = new SessionManager(schema);
///   mgr.StartSession(state);                     // opens / creates file
///   mgr.AddRecord(record);                       // appends record; updates sidecar
///   mgr.SetNewOperatorAndRoll("OP2");            // Auto/DateTime: auto-advance roll
///   mgr.SetNewOperatorAndRoll("OP2", roll: 5);   // Manual: caller supplies new value
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

    public SessionManager(ColumnSchema schema)
    {
        _schema = schema;
    }

    // ── Session lifecycle ──────────────────────────────────────────────────

    /// <summary>
    /// Start a new job session or resume an interrupted one.
    ///
    /// Roll logic (new sessions only; resumes restore from sidecar):
    ///   Manual        — RollNumber stays at caller-supplied value; no automatic change.
    ///   AutoIncrement — RollNumber is set to state.RollStartValue.
    ///   DateTimeStamp — RollTimestamp is generated (yyyyMMddHHmmss).
    ///
    /// Returns the resolved absolute output file path.
    /// </summary>
    public string StartSession(SessionState state)
    {
        if (_writer is not null)
            throw new InvalidOperationException(
                "A session is already open. Call CloseSession() first.");

        _currentSession = state;

        // ── Step 1: initial path from caller-supplied state ────────────────
        string initialPath    = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
        string initialSidecar = GetSidecarPath(initialPath);

        // ── Step 2: attempt sidecar merge at initial path ─────────────────
        bool resumed = TryMergeSidecar(initialSidecar, state);

        // ── Step 3: re-resolve after merge (naming fields may have changed)
        string finalPath    = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
        string finalSidecar = GetSidecarPath(finalPath);

        // ── Step 4: second-chance sidecar lookup if path changed ──────────
        if (!resumed && finalPath != initialPath)
            resumed = TryMergeSidecar(finalSidecar, state);

        // ── Step 5: apply roll-mode logic for brand-new sessions only ─────
        if (!resumed)
            ApplyRollModeOnStart(state);

        // ── Pin the stable path ───────────────────────────────────────────
        _outputPath  = finalPath;
        _sidecarPath = finalSidecar;

        // ── Open / create the Excel file ──────────────────────────────────
        _adapter = state.OutputFormat == OutputFormat.Xls
            ? new XlsAdapter()
            : new XlsxAdapter();

        _writer = new ExcelWriter(_adapter, _schema, state);
        _writer.Open(_outputPath);

        // ── Persist sidecar immediately ───────────────────────────────────
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
        _writer!.AppendRecord(record);
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
    ///   <c>Manual</c>        — <paramref name="manualRoll"/> is required; RollNumber is set to it.
    ///   <c>AutoIncrement</c> — RollNumber is incremented by 1; <paramref name="manualRoll"/> ignored.
    ///   <c>DateTimeStamp</c> — a new yyyyMMddHHmmss timestamp is generated; <paramref name="manualRoll"/> ignored.
    ///
    /// Note: the output file path and sidecar path do not change — changing operator/roll
    /// mid-session does not rename the already-open file.
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

    // ── Public path utilities ─────────────────────────────────────────────

    /// <summary>Pinned output file path for the current session, or null if closed.</summary>
    public string? OutputPath => _outputPath;

    /// <summary>
    /// Returns the sidecar JSON path for a given output file path.
    /// Convention: {outputFile}.vtccp.json
    /// </summary>
    public static string GetSidecarPath(string outputFilePath)
        => outputFilePath + ".vtccp.json";

    // ── Roll mode helpers ─────────────────────────────────────────────────

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

    // ── Sidecar serialisation ─────────────────────────────────────────────

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        WriteIndented = true,
    };

    private static void SaveSidecar(string path, SessionState state)
    {
        var sidecar = new SessionSidecar
        {
            // Job / operator context
            JobName     = state.JobName,
            OperatorId  = state.OperatorId,
            BatchNumber = state.BatchNumber,
            CompanyName = state.CompanyName,
            ProductName = state.ProductName,
            CustomNote  = state.CustomNote,
            User1       = state.User1,
            User2       = state.User2,

            // Roll identifier — all fields preserved so any mode resumes correctly
            RollIncrementMode = state.RollIncrementMode.ToString(),
            RollNumber        = state.RollNumber,
            RollStartValue    = state.RollStartValue,
            RollTimestamp     = state.RollTimestamp,

            // Device metadata
            DeviceSerial    = state.DeviceSerial,
            DeviceName      = state.DeviceName,
            FirmwareVersion = state.FirmwareVersion,
            CalibrationDate = state.CalibrationDate,

            // Output configuration
            OutputFormat    = state.OutputFormat.ToString(),
            OutputDirectory = state.OutputDirectory,
            FileNamePattern = state.FileNamePattern,

            // Session counters
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
    /// Load the sidecar at <paramref name="sidecarPath"/> and merge it into
    /// <paramref name="state"/>. Returns true if a sidecar was found and applied.
    /// Both the xlsx and the sidecar must exist for a valid resume.
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
    /// Restore all saved fields from the sidecar onto the live session state.
    ///
    /// Counters and roll state are always authoritative from the sidecar.
    /// Context fields restore only if the caller left them null, allowing callers
    /// to pre-set overrides before calling StartSession().
    /// </summary>
    private static void ApplySidecarToState(SessionSidecar saved, SessionState state)
    {
        // ── Roll identifier — always authoritative ─────────────────────────
        state.RollNumber     = saved.RollNumber;
        state.RollStartValue = saved.RollStartValue;
        state.RollTimestamp  = saved.RollTimestamp;
        if (saved.RollIncrementMode is not null &&
            Enum.TryParse<RollIncrementMode>(saved.RollIncrementMode, out var mode))
            state.RollIncrementMode = mode;

        // ── Session counters ───────────────────────────────────────────────
        state.RecordCount    = saved.RecordCount;
        state.SessionStarted = saved.SessionStarted;

        // ── Context fields: restore only if caller left them null ──────────
        if (state.JobName        is null && saved.JobName        is not null) state.JobName        = saved.JobName;
        if (state.OperatorId     is null && saved.OperatorId     is not null) state.OperatorId     = saved.OperatorId;
        if (state.BatchNumber    is null && saved.BatchNumber    is not null) state.BatchNumber    = saved.BatchNumber;
        if (state.CompanyName    is null && saved.CompanyName    is not null) state.CompanyName    = saved.CompanyName;
        if (state.ProductName    is null && saved.ProductName    is not null) state.ProductName    = saved.ProductName;
        if (state.CustomNote     is null && saved.CustomNote     is not null) state.CustomNote     = saved.CustomNote;
        if (state.User1          is null && saved.User1          is not null) state.User1          = saved.User1;
        if (state.User2          is null && saved.User2          is not null) state.User2          = saved.User2;
        if (state.DeviceSerial   is null && saved.DeviceSerial   is not null) state.DeviceSerial   = saved.DeviceSerial;
        if (state.DeviceName     is null && saved.DeviceName     is not null) state.DeviceName     = saved.DeviceName;
        if (state.FirmwareVersion is null && saved.FirmwareVersion is not null) state.FirmwareVersion = saved.FirmwareVersion;
        if (state.CalibrationDate is null && saved.CalibrationDate is not null) state.CalibrationDate = saved.CalibrationDate;
        if (state.OutputDirectory is null && saved.OutputDirectory is not null) state.OutputDirectory = saved.OutputDirectory;
        if (state.FileNamePattern is null && saved.FileNamePattern is not null) state.FileNamePattern = saved.FileNamePattern;
        if (saved.OutputFormat    is not null && Enum.TryParse<OutputFormat>(saved.OutputFormat, out var fmt))
            state.OutputFormat = fmt;
    }

    // ── IDisposable ────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _writer?.Dispose();
        _adapter?.Dispose();
    }

    // ── Private helpers ────────────────────────────────────────────────────

    private void EnsureOpen()
    {
        if (_writer is null)
            throw new InvalidOperationException(
                "No session is open. Call StartSession() first.");
    }
}
