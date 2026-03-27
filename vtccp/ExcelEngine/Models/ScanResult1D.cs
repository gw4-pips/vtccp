namespace ExcelEngine.Models;

/// <summary>
/// Result from a single scan pass in ISO 15416 1D verification.
/// DMV TruCheck records up to 10 scan passes per verification event.
/// Each scan measures: Edge, Ref (Reflectance), SC, MinEC, MOD, DEF, DCOD, DEC, LQZ, RQZ, HQZ → per-scan Grade.
/// </summary>
public sealed class ScanResult1D
{
    public int ScanNumber { get; init; }

    public decimal? Edge { get; init; }
    public string? Reflectance { get; init; }  // Rl/Rd format, e.g. "89/4"
    public decimal? SC { get; init; }
    public decimal? MinEC { get; init; }
    public decimal? MOD { get; init; }
    public decimal? Defect { get; init; }
    public string? DCOD { get; init; }  // e.g. "10/10"
    public decimal? DEC { get; init; }
    public decimal? LQZ { get; init; }
    public decimal? RQZ { get; init; }
    public decimal? HQZ { get; init; }

    public GradingResult? PerScanGrade { get; init; }
}
