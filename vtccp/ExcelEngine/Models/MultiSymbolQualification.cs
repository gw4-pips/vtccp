namespace ExcelEngine.Models;

public enum MultiSymbolQualificationStatus
{
    Verified,
    Qualified,
    Unverified,
    Rejected,
}

public sealed class MultiSymbolIdentityEvidence
{
    public int Ordinal { get; init; }
    public string? Symbology { get; init; }
    public string Family { get; init; } = string.Empty;
    public string? Gtin14 { get; init; }
}

public sealed class MultiSymbolQualification
{
    public MultiSymbolQualificationStatus Status { get; init; }
    public IReadOnlyList<string> Reasons { get; init; } = [];
    public IReadOnlyList<MultiSymbolIdentityEvidence> Symbols { get; init; } = [];
    public IReadOnlyList<int> MatchingSymbols { get; init; } = [];
    public IReadOnlyList<int> MismatchingSymbols { get; init; } = [];
}