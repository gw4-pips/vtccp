namespace ExcelEngine.Models;

/// <summary>
/// Captures the operator-entered session context that applies to all records within a job.
/// These fields come entirely from the VTCCP UI (not from the device) because DMV devices
/// cannot pass all header fields through DMCC programming — and certain characters (&, /, etc.)
/// conflict with DMCC command syntax.
/// </summary>
public sealed class SessionState
{
    public string? JobName     { get; set; }
    public string? OperatorId  { get; set; }
    public string? BatchNumber { get; set; }
    public string? CompanyName { get; set; }
    public string? ProductName { get; set; }
    public string? CustomNote  { get; set; }

    // ── Roll identifier ──────────────────────────────────────────────────────

    /// <summary>
    /// How the roll identifier is produced and advanced.
    /// Default: <see cref="RollIncrementMode.Manual"/> (explicit operator action required).
    /// </summary>
    public RollIncrementMode RollIncrementMode { get; set; } = RollIncrementMode.Manual;

    /// <summary>
    /// Numeric roll counter.
    /// - <c>Manual</c>: operator-supplied value; VTCCP never changes it automatically.
    /// - <c>AutoIncrement</c>: starts at <see cref="RollStartValue"/>; incremented by 1
    ///   on each <c>SetNewOperatorAndRoll()</c> call.
    /// - <c>DateTimeStamp</c>: not used as the roll label; see <see cref="RollTimestamp"/>.
    /// </summary>
    public int RollNumber { get; set; } = 1;

    /// <summary>
    /// Starting value for <see cref="RollIncrementMode.AutoIncrement"/> sessions.
    /// <c>SessionManager.StartSession()</c> sets <see cref="RollNumber"/> to this
    /// value when opening a brand-new (non-resume) session.
    /// </summary>
    public int RollStartValue { get; set; } = 1;

    /// <summary>
    /// Timestamp string used as the roll label in <see cref="RollIncrementMode.DateTimeStamp"/>
    /// mode.  Format: yyyyMMddHHmmss.  Set by <c>SessionManager</c> at session open and on
    /// each <c>SetNewOperatorAndRoll()</c> call; preserved across sidecar resumes.
    /// </summary>
    public string? RollTimestamp { get; set; }

    /// <summary>
    /// The effective roll identifier string used in Excel cell values, the title row,
    /// and the <c>{Roll}</c> file-name pattern token.
    /// Derived from the active <see cref="RollIncrementMode"/>:
    /// - Manual / AutoIncrement → <see cref="RollNumber"/> formatted as decimal string.
    /// - DateTimeStamp → <see cref="RollTimestamp"/> (yyyyMMddHHmmss).
    /// </summary>
    public string RollLabel => RollIncrementMode == RollIncrementMode.DateTimeStamp
        ? (RollTimestamp ?? RollNumber.ToString())
        : RollNumber.ToString();

    // ── Device-supplied fields (Phase 2) ────────────────────────────────────

    public string?   DeviceSerial    { get; set; }
    public string?   DeviceName      { get; set; }
    public string?   FirmwareVersion { get; set; }
    public DateTime? CalibrationDate { get; set; }

    // ── Output preferences ───────────────────────────────────────────────────

    public OutputFormat OutputFormat  { get; set; } = OutputFormat.Xlsx;
    public string?      OutputDirectory { get; set; }

    /// <summary>
    /// Optional custom file-name pattern.  Supported tokens:
    ///   {Job}      — SanitizeFileName(JobName) or "VTCCP"
    ///   {Op}       — SanitizeFileName(OperatorId) or ""
    ///   {Roll}     — RollLabel (decimal number or yyyyMMddHHmmss depending on mode)
    ///   {Date}     — SessionStarted "yyyy-MM-dd"
    ///   {DateTime} — SessionStarted "yyyy-MM-dd_HH-mm"
    /// When null the default Webscan TruCheck convention is used ({Job}_{Date} or VTCCP_{Date}).
    /// Example: "{Job}_{Op}_Roll{Roll}_{Date}"
    /// </summary>
    public string? FileNamePattern { get; set; }

    // ── Session tracking ─────────────────────────────────────────────────────

    public DateTime SessionStarted { get; set; } = DateTime.Now;
    public int      RecordCount    { get; set; } = 0;

    /// <summary>User-defined fields (User 1, User 2 in Webscan schema)</summary>
    public string? User1 { get; set; }
    public string? User2 { get; set; }
}
