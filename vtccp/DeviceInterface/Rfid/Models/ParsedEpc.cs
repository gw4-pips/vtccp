namespace DeviceInterface.Rfid.Models;

/// <summary>EPC encoding scheme — derived from the header byte.</summary>
public enum EpcScheme
{
    Unknown    = 0,
    Sgtin96    = 0x30,
    Sgtin198   = 0x36,
}

/// <summary>
/// Decoded fields from an EPC (Electronic Product Code).
/// Only populated fields are non-null; all others remain null for the parsed scheme.
/// </summary>
public sealed record ParsedEpc
{
    /// <summary>Raw EPC bytes that were parsed.</summary>
    public required byte[] EpcBytes { get; init; }

    /// <summary>Hex representation of the raw EPC.</summary>
    public string EpcHex => Convert.ToHexString(EpcBytes);

    /// <summary>Identified EPC encoding scheme.</summary>
    public EpcScheme Scheme { get; init; }

    // ── SGTIN fields (populated for Sgtin96 and Sgtin198) ─────────────────────

    /// <summary>
    /// Filter value (3 bits, 0–7).
    /// Conveys packaging hierarchy: 0=All, 1=POS item, 2=Case, etc.
    /// NOT the same as GTIN-14 indicator digit (which is embedded in ItemReference).
    /// </summary>
    public int? Filter { get; init; }

    /// <summary>GS1 partition value (0–6); determines GCP length.</summary>
    public int? Partition { get; init; }

    /// <summary>
    /// GS1 Company Prefix as a decimal string, zero-padded to the correct digit count.
    /// e.g. "0614141" for a 7-digit GCP.
    /// </summary>
    public string? CompanyPrefix { get; init; }

    /// <summary>
    /// Item Reference as a decimal string, zero-padded to (13-L) digits.
    /// Includes the GTIN indicator digit as its leading digit for SGTIN encoding.
    /// </summary>
    public string? ItemReference { get; init; }

    /// <summary>
    /// Serial number as a decimal string (SGTIN-96: up to 38 bits = max 274,877,906,943)
    /// or an alphanumeric string (SGTIN-198: up to 20 ASCII characters).
    /// </summary>
    public string? Serial { get; init; }

    /// <summary>
    /// Reconstructed GTIN-14 (14 decimal digits including check digit).
    /// Formula: CompanyPrefix(L) + ItemReference(13-L) + GS1CheckDigit.
    /// </summary>
    public string? Gtin14 { get; init; }

    /// <summary>
    /// Error/warning message if parsing succeeded partially or with caveats.
    /// Null on a clean parse.
    /// </summary>
    public string? ParseWarning { get; init; }
}
