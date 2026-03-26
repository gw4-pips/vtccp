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
///   4. Persist session state to a JSON sidecar so an interrupted session can resume.
///   5. Manage roll number increment on each StartSession call.
///
/// Usage:
///   var mgr = new SessionManager(schema);
///   mgr.StartSession(state);          // opens or creates output file
///   mgr.AddRecord(record);            // appends a verification record
///   mgr.CloseSession();               // saves and closes the file
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
    /// Start a new job session. If a sidecar file exists for the same job on the same day,
    /// the session is resumed (appending to the existing output file).
    /// Roll number is automatically incremented unless overridden.
    /// </summary>
    public string StartSession(SessionState state)
    {
        if (_writer is not null)
            throw new InvalidOperationException(
                "A session is already open. Call CloseSession() first.");

        _currentSession = state;

        // Resolve the output file path.
        string outputPath = ExcelFileManager.ResolveOutputPath(state, state.OutputFormat);

        // Check sidecar for resume data.
        string sidecarPath = GetSidecarPath(outputPath);
        if (File.Exists(sidecarPath) && File.Exists(outputPath))
        {
            try
            {
                var saved = LoadSidecar(sidecarPath);
                // Resume: preserve roll number and record count from previous session.
                if (saved is not null)
                {
                    state.RecordCount = saved.RecordCount;
                    if (state.RollNumber <= 1 && saved.RollNumber > 0)
                        state.RollNumber = saved.RollNumber;
                }
            }
            catch (JsonException)
            {
                // Corrupt sidecar — start fresh.
            }
        }

        // Open (or create) the Excel file via the correct adapter.
        _adapter = state.OutputFormat == OutputFormat.Xls
            ? new XlsAdapter()
            : new XlsxAdapter();

        _writer = new ExcelWriter(_adapter, _schema, state);
        _writer.Open(outputPath);

        // Persist the sidecar immediately.
        SaveSidecar(sidecarPath, state);

        return outputPath;
    }

    /// <summary>
    /// Append a single VerificationRecord to the active session's output file.
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
    /// Increments the roll number.
    /// </summary>
    public void SetNewOperatorAndRoll(string operatorId)
    {
        EnsureOpen();
        _currentSession!.OperatorId = operatorId;
        _currentSession!.RollNumber++;
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
            JobName       = state.JobName,
            OperatorId    = state.OperatorId,
            RollNumber    = state.RollNumber,
            BatchNumber   = state.BatchNumber,
            CompanyName   = state.CompanyName,
            ProductName   = state.ProductName,
            RecordCount   = state.RecordCount,
            SessionStarted= state.SessionStarted,
            OutputFormat  = state.OutputFormat.ToString(),
        };
        File.WriteAllText(path, JsonSerializer.Serialize(sidecar, _jsonOpts));
    }

    private static SessionSidecar? LoadSidecar(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<SessionSidecar>(json);
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
