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
/// Responsibilities:
///   1. Track current operator, job name, roll number, and other session metadata.
///   2. Resolve the output file path at session-open time; pin it for the lifetime
///      of the session so sidecar operations always target the correct file.
///   3. Append VerificationRecords to the active output file.
///   4. Persist all session state to a JSON sidecar so an interrupted session
///      can be fully resumed (all context fields, not just counters).
///
/// Roll number semantics (matches Webscan TruCheck "Set New Operator/Roll" behaviour):
///   - StartSession with a fresh (no sidecar) file: caller controls initial RollNumber.
///   - StartSession that finds a matching sidecar (resume): saved roll is restored.
///     Roll is NOT incremented on resume, because the operator did not start a new roll —
///     they are continuing the interrupted one.
///   - SetNewOperatorAndRoll() increments the roll number mid-session and updates the sidecar.
///
/// File path stability guarantee:
///   The output path and sidecar path are resolved ONCE at StartSession and stored
///   in _outputPath / _sidecarPath.  No subsequent call re-derives these paths, so
///   state changes (e.g. pattern tokens changing) cannot orphan sidecars or target
///   the wrong file.
///
/// Resume path consistency:
///   1. Initial path resolved from caller-supplied state (minimal context).
///   2. Sidecar loaded and merged into state.
///   3. Path re-resolved from merged state.
///   4. If path changed, update _outputPath / _sidecarPath; re-check for a sidecar
///      at the new path (second-chance resume).
///   5. Open the file at the final stable path.
///
/// Usage:
///   var mgr = new SessionManager(schema);
///   mgr.StartSession(state);          // opens or creates output file
///   mgr.AddRecord(record);            // appends a verification record
///   mgr.CloseSession();               // saves and closes the file; deletes sidecar
/// </summary>
public sealed class SessionManager : IDisposable
{
    private readonly ColumnSchema _schema;
    private SessionState? _currentSession;
    private IExcelAdapter? _adapter;
    private ExcelWriter? _writer;

    // Pinned at StartSession; never re-derived during a session.
    private string? _outputPath;
    private string? _sidecarPath;

    private bool _disposed;

    public SessionState? CurrentSession => _currentSession;
    public bool IsSessionOpen => _currentSession is not null && _writer is not null;
    public int RecordsWritten => _currentSession?.RecordCount ?? 0;

    public SessionManager(ColumnSchema schema)
    {
        _schema = schema;
    }

    // ── Session lifecycle ─────────────────────────────────────────────────────

    /// <summary>
    /// Start a new job session (or resume an interrupted one).
    ///
    /// Resume behaviour:
    ///   If a sidecar exists for the resolved output file, all saved context is merged
    ///   back onto <paramref name="state"/> before the file is opened.  The path is then
    ///   re-resolved from the merged state (in case naming fields change), and the final
    ///   stable path is pinned for the session lifetime.
    ///
    /// Roll number:
    ///   - New session: uses state.RollNumber as supplied by the caller.
    ///   - Resume: restored from sidecar; not auto-incremented.
    ///   - Call SetNewOperatorAndRoll() to advance the roll mid-session.
    ///
    /// Returns the resolved absolute output file path.
    /// </summary>
    public string StartSession(SessionState state)
    {
        if (_writer is not null)
            throw new InvalidOperationException(
                "A session is already open. Call CloseSession() first.");

        _currentSession = state;

        // Step 1: Resolve initial path from caller-supplied state.
        string initialPath   = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
        string initialSidecar = GetSidecarPath(initialPath);

        // Step 2: Attempt to load and merge sidecar from initial path.
        bool resumed = TryMergeSidecar(initialSidecar, state);

        // Step 3: Re-resolve the path from merged state (naming fields may have changed).
        string finalPath    = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
        string finalSidecar = GetSidecarPath(finalPath);

        // Step 4: If path changed and no sidecar was loaded yet, attempt second-chance
        // sidecar lookup at the new path.
        if (!resumed && finalPath != initialPath)
            TryMergeSidecar(finalSidecar, state);

        // Step 5: Pin the stable path for this session.
        _outputPath  = finalPath;
        _sidecarPath = finalSidecar;

        // Step 6: Open (or create) the Excel file via the correct adapter.
        _adapter = state.OutputFormat == OutputFormat.Xls
            ? new XlsAdapter()
            : new XlsxAdapter();

        _writer = new ExcelWriter(_adapter, _schema, state);
        _writer.Open(_outputPath);

        // Step 7: Persist sidecar immediately after opening.
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
    /// Save and close the active session. Removes the sidecar file (job complete).
    /// </summary>
    public void CloseSession()
    {
        if (_writer is null) return;

        _writer.Save();
        _writer.Dispose();
        _writer = null;

        _adapter?.Dispose();
        _adapter = null;

        // Remove sidecar — job is closed cleanly.
        if (_sidecarPath is not null && File.Exists(_sidecarPath))
            File.Delete(_sidecarPath);

        _currentSession = null;
        _outputPath     = null;
        _sidecarPath    = null;
    }

    /// <summary>
    /// Set or change the operator and advance the roll number (matching Webscan
    /// "Set New Operator/Roll" button behaviour). The sidecar is updated immediately.
    /// </summary>
    public void SetNewOperatorAndRoll(string operatorId)
    {
        EnsureOpen();
        _currentSession!.OperatorId = operatorId;
        _currentSession!.RollNumber++;
        // Note: _outputPath and _sidecarPath do NOT change — the operator change
        // does not rename the already-open file.
        SaveSidecar(_sidecarPath!, _currentSession!);
    }

    // ── Public path utilities ─────────────────────────────────────────────────

    /// <summary>Returns the pinned output file path for the current session, or null if closed.</summary>
    public string? OutputPath => _outputPath;

    /// <summary>
    /// Returns the output file path for the given session state (without opening it).
    /// </summary>
    public static string GetOutputPath(SessionState state)
        => ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);

    /// <summary>
    /// Returns the sidecar JSON path for a given output file path.
    /// Convention: {outputFile}.vtccp.json
    /// </summary>
    public static string GetSidecarPath(string outputFilePath)
        => outputFilePath + ".vtccp.json";

    // ── Sidecar serialisation ─────────────────────────────────────────────────

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        WriteIndented = true,
    };

    private static void SaveSidecar(string path, SessionState state)
    {
        var sidecar = new SessionSidecar
        {
            // Job / operator context — full snapshot
            JobName         = state.JobName,
            OperatorId      = state.OperatorId,
            RollNumber      = state.RollNumber,
            BatchNumber     = state.BatchNumber,
            CompanyName     = state.CompanyName,
            ProductName     = state.ProductName,
            CustomNote      = state.CustomNote,
            User1           = state.User1,
            User2           = state.User2,

            // Device metadata (Phase 2)
            DeviceSerial    = state.DeviceSerial,
            DeviceName      = state.DeviceName,
            FirmwareVersion = state.FirmwareVersion,
            CalibrationDate = state.CalibrationDate,

            // Output configuration
            OutputFormat    = state.OutputFormat.ToString(),
            OutputDirectory = state.OutputDirectory,
            FileNamePattern = state.FileNamePattern,

            // Session counters
            SessionStarted  = state.SessionStarted,
            RecordCount     = state.RecordCount,
        };
        File.WriteAllText(path, JsonSerializer.Serialize(sidecar, _jsonOpts));
    }

    private static SessionSidecar? LoadSidecar(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<SessionSidecar>(json);
    }

    /// <summary>
    /// Try to load the sidecar at <paramref name="sidecarPath"/> and merge it into
    /// <paramref name="state"/>. Returns true if a sidecar was found and applied.
    /// </summary>
    private static bool TryMergeSidecar(string sidecarPath, SessionState state)
    {
        // Both the output file AND sidecar must exist to resume.
        string outputPath = sidecarPath[..^".vtccp.json".Length]; // strip suffix
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
    /// Counters (RecordCount, RollNumber, SessionStarted) are always restored as
    /// authoritative. Context fields restore only if the caller left them null,
    /// allowing callers to pre-set overrides before StartSession.
    /// </summary>
    private static void ApplySidecarToState(SessionSidecar saved, SessionState state)
    {
        // Counters are authoritative in the sidecar.
        state.RecordCount    = saved.RecordCount;
        state.RollNumber     = saved.RollNumber;
        state.SessionStarted = saved.SessionStarted;

        // Context fields: restore from sidecar only if caller left them null.
        if (state.JobName        is null && saved.JobName        is not null) state.JobName        = saved.JobName;
        if (state.OperatorId     is null && saved.OperatorId     is not null) state.OperatorId     = saved.OperatorId;
        if (state.BatchNumber    is null && saved.BatchNumber    is not null) state.BatchNumber    = saved.BatchNumber;
        if (state.CompanyName    is null && saved.CompanyName    is not null) state.CompanyName    = saved.CompanyName;
        if (state.ProductName    is null && saved.ProductName    is not null) state.ProductName    = saved.ProductName;
        if (state.CustomNote     is null && saved.CustomNote     is not null) state.CustomNote     = saved.CustomNote;
        if (state.User1          is null && saved.User1          is not null) state.User1          = saved.User1;
        if (state.User2          is null && saved.User2          is not null) state.User2          = saved.User2;

        // Device metadata (Phase 2)
        if (state.DeviceSerial    is null && saved.DeviceSerial    is not null) state.DeviceSerial    = saved.DeviceSerial;
        if (state.DeviceName      is null && saved.DeviceName      is not null) state.DeviceName      = saved.DeviceName;
        if (state.FirmwareVersion is null && saved.FirmwareVersion is not null) state.FirmwareVersion = saved.FirmwareVersion;
        if (state.CalibrationDate is null && saved.CalibrationDate is not null) state.CalibrationDate = saved.CalibrationDate;

        // Output configuration
        if (state.OutputDirectory is null && saved.OutputDirectory is not null) state.OutputDirectory = saved.OutputDirectory;
        if (state.FileNamePattern is null && saved.FileNamePattern is not null) state.FileNamePattern = saved.FileNamePattern;
        if (saved.OutputFormat    is not null && Enum.TryParse<OutputFormat>(saved.OutputFormat, out var fmt))
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
            throw new InvalidOperationException("No session is open. Call StartSession() first.");
    }
}
