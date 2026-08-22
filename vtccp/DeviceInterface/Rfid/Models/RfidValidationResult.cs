namespace DeviceInterface.Rfid.Models;

/// <summary>Outcome of an RFID cross-validation check against barcode data.</summary>
public enum RfidValidationStatus
{
    /// <summary>Both RFID and barcode were read; all checked fields match.</summary>
    Pass = 0,

    /// <summary>
    /// RFID was read but one or more fields do not match the barcode
    /// (GTIN-14 mismatch, serial mismatch, etc.).
    /// </summary>
    Fail = 1,

    /// <summary>No RFID tag was detected within the scan window.</summary>
    NoTag = 2,

    /// <summary>RFID was read but the EPC could not be parsed (unknown scheme or corrupt).</summary>
    ParseError = 3,

    /// <summary>RFID reader is not configured or not connected; validation was skipped.</summary>
    Skipped = 4,

    /// <summary>Multiple distinct EPC values were detected; ambiguous result.</summary>
    MultipleTagsDetected = 5,
}

/// <summary>Outcome of looking up an EPC's encoded GS1 Company Prefix.</summary>
public enum GcpValidationStatus
{
    /// <summary>No GCP result is available because the check was not applicable or not run.</summary>
    NotChecked = 0,

    /// <summary>The prefix exists in the GS1 table and its registered length matches the EPC partition.</summary>
    Valid = 1,

    /// <summary>The prefix exists in the GS1 table but its registered length disagrees with the EPC partition.</summary>
    Invalid = 2,

    /// <summary>The prefix does not exist in the loaded GS1 table.</summary>
    NotFound = 3,
}

/// <summary>
/// Complete result of one RFID cross-validation cycle, bound to the barcode
/// VerificationRecord that triggered the paired RFID scan.
/// </summary>
public sealed record RfidValidationResult
{
    public RfidValidationStatus Status { get; init; } = RfidValidationStatus.Skipped;

    // ── Raw EPC data ────────────────────────────────────────────────────────────

    /// <summary>
    /// All distinct EPC values detected during the scan window.
    /// Empty when Status is NoTag or Skipped.
    /// </summary>
    public IReadOnlyList<EpcReadResult> RawReads { get; init; } = [];

    /// <summary>
    /// The single EPC selected for validation (first read when unique, null when ambiguous or absent).
    /// </summary>
    public EpcReadResult? SelectedRead { get; init; }

    // ── Parsed EPC ──────────────────────────────────────────────────────────────

    public ParsedEpc? ParsedEpc { get; init; }

    // ── Cross-validation fields ─────────────────────────────────────────────────

    /// <summary>GTIN-14 from the RFID EPC (null when EPC could not be decoded).</summary>
    public string? RfidGtin14 { get; init; }

    /// <summary>GTIN-14 extracted from the barcode DecodedData (AI 01).</summary>
    public string? BarcodeGtin14 { get; init; }

    /// <summary>True when RfidGtin14 == BarcodeGtin14 (and both non-null).</summary>
    public bool Gtin14Matches => RfidGtin14 is not null && RfidGtin14 == BarcodeGtin14;

    /// <summary>Serial number from the RFID EPC.</summary>
    public string? RfidSerial { get; init; }

    /// <summary>Serial number from the barcode (AI 21), null if not present.</summary>
    public string? BarcodeSerial { get; init; }

    /// <summary>True when both serials are non-null and equal.</summary>
    public bool SerialMatches => RfidSerial is not null && RfidSerial == BarcodeSerial;

    // ── GCP validation ──────────────────────────────────────────────────────────

    /// <summary>Explicit GS1 Company Prefix lookup outcome.</summary>
    public GcpValidationStatus GcpStatus { get; init; } = GcpValidationStatus.NotChecked;

    /// <summary>
    /// GCP length registered for the company prefix in the lookup table.
    /// Null when the prefix was not found or the lookup was not checked.
    /// </summary>
    public int? GcpRegisteredLength { get; init; }

    /// <summary>
    /// Legacy Boolean projection of <see cref="GcpStatus"/>.
    /// True means valid, false means a found prefix had an incompatible encoded length,
    /// and null means not checked or not found.
    /// </summary>
    public bool? GcpValid => GcpStatus switch
    {
        GcpValidationStatus.Valid => true,
        GcpValidationStatus.Invalid => false,
        _ => null,
    };

    // ── Timing ──────────────────────────────────────────────────────────────────

    /// <summary>How long the RFID scan window was open (ms).</summary>
    public int ScanWindowMs { get; init; }

    // ── Human-readable summary ──────────────────────────────────────────────────

    /// <summary>
    /// Semicolon-separated mismatch tokens. Possible tokens:
    ///   GTIN14:NoBarcodeData          — barcode carried no AI (01) to compare against
    ///   GTIN14:RFID=&lt;x&gt;,BC=&lt;y&gt;       — GTIN-14 values differ
    ///   Serial:MissingFromTag         — barcode has AI (21) but RFID tag EPC has no serial
    ///   Serial:RFID=&lt;x&gt;,BC=&lt;y&gt;       — serial values differ
    /// GCP registration is intentionally excluded — it is not a pass/fail criterion.
    /// Null on Pass or when no comparison was possible.
    /// </summary>
    public string? MismatchDetail { get; init; }
}
