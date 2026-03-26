namespace ExcelEngine.Session;

/// <summary>
/// Serialized snapshot of an in-progress session. Written to a .vtccp.json sidecar file
/// alongside the Excel output so an interrupted session can be resumed.
/// All fields from SessionState that affect header content or continuation are stored.
/// </summary>
internal sealed class SessionSidecar
{
    // Job / operator context
    public string? JobName        { get; set; }
    public string? OperatorId     { get; set; }
    public int     RollNumber     { get; set; }
    public string? BatchNumber    { get; set; }
    public string? CompanyName    { get; set; }
    public string? ProductName    { get; set; }
    public string? CustomNote     { get; set; }

    // User-defined fields
    public string? User1          { get; set; }
    public string? User2          { get; set; }

    // Device metadata (populated in Phase 2; preserved across resumes)
    public string?   DeviceSerial    { get; set; }
    public string?   DeviceName      { get; set; }
    public string?   FirmwareVersion { get; set; }
    public DateTime? CalibrationDate { get; set; }

    // Output configuration
    public string?   OutputFormat    { get; set; }
    public string?   OutputDirectory { get; set; }
    public string?   FileNamePattern { get; set; }

    // Session counters / timestamps
    public DateTime SessionStarted  { get; set; }
    public int      RecordCount     { get; set; }

    // Resume-only flag: when true the sidecar was written by an in-progress session
    // that started a new roll. Consumers should not increment roll again on resume.
    public bool     IsNewRollSaved  { get; set; }
}
