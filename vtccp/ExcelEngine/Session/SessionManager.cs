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
///   2. Resolve the output file path from the session context (Webscan convention).
///   3. Append VerificationRecords to the active output file.
///   4. Persist all session state to a JSON sidecar so an interrupted session
///      can be fully resumed (all context fields, not just counters).
///   5. Roll number semantics:
///        - A new session (no sidecar) starts with the caller-supplied RollNumber.
///        - When the same output file already exists AND a sidecar exists (resume),
///          the saved roll number is restored; it is NOT auto-incremented on resume
///          because the operator did not physically "start a new roll".
///        - `SetNewOperatorAndRoll()` increments the roll number mid-session,
///          matching the Webscan "Set New Operator/Roll" button behaviour.
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
    /// Resume behaviour: if both the output file and its sidecar already exist for
    /// this session, all saved context is merged back onto <paramref name="state"/>
    /// (operator, job, company, product, batch, notes, device metadata, roll number,
    /// and record count). Callers can still override any field before calling AddRecord.
    ///
    /// Roll number: for a brand-new session the caller supplies <c>state.RollNumber</c>;
    /// for a resume the sidecar's roll number is restored. Call
    /// <see cref="SetNewOperatorAndRoll"/> to advance the roll mid-session.
    ///
    /// Returns the resolved absolute output file path.
    /// </summary>
    public string StartSession(SessionState state)
    {
        if (_writer is not null)
            throw new InvalidOperationException(
                "A session is already open. Call CloseSession() first.");

        _currentSession = state;

        // Resolve the output file path (may depend on FileNamePattern, if set).
        string outputPath = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);
        string sidecarPath = GetSidecarPath(outputPath);

        // Attempt resume if both the output file and sidecar exist.
        if (File.Exists(sidecarPath) && File.Exists(outputPath))
        {
            try
            {
                var saved = LoadSidecar(sidecarPath);
                if (saved is not null)
                    ApplySidecarToState(saved, state);
            }
            catch (JsonException)
            {
                // Corrupt sidecar — start fresh (file will be appended from scratch).
            }
        }

        // Open (or create) the Excel file via the correct adapter.
        _adapter = state.OutputFormat == OutputFormat.Xls
            ? new XlsAdapter()
            : new XlsxAdapter();

        _writer = new ExcelWriter(_adapter, _schema, state);
        _writer.Open(outputPath);

        // Persist the sidecar immediately after opening.
        SaveSidecar(sidecarPath, state);

        return outputPath;
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
        // Update sidecar after each record for crash-safety.
        string outputPath = ExcelFileManager.ResolveOutputPath(_currentSession!, _currentSession!.OutputFormat);
        SaveSidecar(GetSidecarPath(outputPath), _currentSession!);
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
        if (_currentSession is not null)
        {
            string outputPath = ExcelFileManager.ResolveOutputPath(_currentSession, _currentSession.OutputFormat);
            string sidecarPath = GetSidecarPath(outputPath);
            if (File.Exists(sidecarPath))
                File.Delete(sidecarPath);
        }

        _currentSession = null;
    }

    /// <summary>
    /// Set or change the operator for the current session (matching Webscan "Set New Operator/Roll").
    /// Increments the roll number and updates the sidecar.
    /// </summary>
    public void SetNewOperatorAndRoll(string operatorId)
    {
        EnsureOpen();
        _currentSession!.OperatorId = operatorId;
        _currentSession!.RollNumber++;
        // Persist the roll change immediately.
        string outputPath = ExcelFileManager.ResolveOutputPath(_currentSession!, _currentSession!.OutputFormat);
        SaveSidecar(GetSidecarPath(outputPath), _currentSession!);
    }

    // ── Public path utilities ─────────────────────────────────────────────────

    /// <summary>
    /// Returns the output file path for the given session (without opening it).
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
            // Job / operator context
            JobName         = state.JobName,
            OperatorId      = state.OperatorId,
            RollNumber      = state.RollNumber,
            BatchNumber     = state.BatchNumber,
            CompanyName     = state.CompanyName,
            ProductName     = state.ProductName,
            CustomNote      = state.CustomNote,
            User1           = state.User1,
            User2           = state.User2,

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
    /// Restore all saved fields from the sidecar onto the live session state.
    /// Only non-null sidecar fields overwrite the caller-supplied values, so
    /// callers can pre-set fields they want to override before StartSession.
    /// Record count and roll number are always restored (they are authoritative
    /// in the sidecar — the caller cannot know the current file row count or roll).
    /// </summary>
    private static void ApplySidecarToState(SessionSidecar saved, SessionState state)
    {
        // Counters are authoritative in the sidecar.
        state.RecordCount  = saved.RecordCount;
        state.RollNumber   = saved.RollNumber;
        state.SessionStarted = saved.SessionStarted;

        // Context fields: restore from sidecar only if caller left them null/default.
        if (state.JobName     is null && saved.JobName     is not null) state.JobName     = saved.JobName;
        if (state.OperatorId  is null && saved.OperatorId  is not null) state.OperatorId  = saved.OperatorId;
        if (state.BatchNumber is null && saved.BatchNumber is not null) state.BatchNumber = saved.BatchNumber;
        if (state.CompanyName is null && saved.CompanyName is not null) state.CompanyName = saved.CompanyName;
        if (state.ProductName is null && saved.ProductName is not null) state.ProductName = saved.ProductName;
        if (state.CustomNote  is null && saved.CustomNote  is not null) state.CustomNote  = saved.CustomNote;
        if (state.User1       is null && saved.User1       is not null) state.User1       = saved.User1;
        if (state.User2       is null && saved.User2       is not null) state.User2       = saved.User2;

        // Device metadata (Phase 2)
        if (state.DeviceSerial    is null && saved.DeviceSerial    is not null) state.DeviceSerial    = saved.DeviceSerial;
        if (state.DeviceName      is null && saved.DeviceName      is not null) state.DeviceName      = saved.DeviceName;
        if (state.FirmwareVersion is null && saved.FirmwareVersion is not null) state.FirmwareVersion = saved.FirmwareVersion;
        if (state.CalibrationDate is null && saved.CalibrationDate is not null) state.CalibrationDate = saved.CalibrationDate;

        // Output config
        if (state.OutputDirectory is null && saved.OutputDirectory is not null) state.OutputDirectory = saved.OutputDirectory;
        if (state.FileNamePattern is null && saved.FileNamePattern is not null) state.FileNamePattern = saved.FileNamePattern;
        if (saved.OutputFormat is not null && Enum.TryParse<OutputFormat>(saved.OutputFormat, out var fmt))
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
