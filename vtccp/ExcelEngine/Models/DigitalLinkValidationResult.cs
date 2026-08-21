namespace ExcelEngine.Models;

/// <summary>
/// The outcome of VCCS's independent GS1 syntax check.
/// This is intentionally distinct from a verifier's native Data Format Check.
/// </summary>
public enum DigitalLinkValidationStatus
{
    NotApplicable,
    Valid,
    Invalid,
    Unavailable,
}

/// <summary>
/// VCCS-owned validation metadata for a decoded GS1 Digital Link URI or
/// GS1 Element String.
/// It must never be presented as a native verifier result or used to replace
/// verifier-provided grades or Data Format Check rows.
/// </summary>
public sealed record class DigitalLinkValidationResult
{
    public const string VccsSource = "VCCS / GS1 Digital Link syntax validation";
    public const string VccsElementStringSource = "VCCS / GS1 Element String syntax validation";

    public DigitalLinkValidationStatus Status { get; init; } =
        DigitalLinkValidationStatus.NotApplicable;

    /// <summary>Always identifies this result as VCCS-owned validation.</summary>
    public string Source { get; init; } = VccsSource;

    /// <summary>Official GS1 Syntax Engine version when the engine was used.</summary>
    public string? EngineVersion { get; init; }

    /// <summary>Short human-readable reason or diagnostic; never vendor output.</summary>
    public string? Detail { get; init; }
}