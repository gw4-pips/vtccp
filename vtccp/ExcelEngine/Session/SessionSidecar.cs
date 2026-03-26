namespace ExcelEngine.Session;

/// <summary>
/// Serialized snapshot of an in-progress session. Written to a .vtccp.json sidecar file
/// alongside the Excel output so an interrupted session can be resumed.
/// </summary>
internal sealed class SessionSidecar
{
    public string? JobName        { get; set; }
    public string? OperatorId     { get; set; }
    public int     RollNumber     { get; set; }
    public string? BatchNumber    { get; set; }
    public string? CompanyName    { get; set; }
    public string? ProductName    { get; set; }
    public int     RecordCount    { get; set; }
    public DateTime SessionStarted { get; set; }
    public string? OutputFormat   { get; set; }
}
