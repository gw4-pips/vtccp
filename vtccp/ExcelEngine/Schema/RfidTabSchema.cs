namespace ExcelEngine.Schema;

/// <summary>
/// Column definitions for the "RFID Scans" auxiliary worksheet.
///
/// Each row represents one RFID cross-validation result, keyed to a barcode
/// verification row by timestamp and row number.
///
/// Columns are intentionally minimal for Phase 0 POC; extend as needed.
/// </summary>
public static class RfidTabSchema
{
    // ── Column position constants (1-based, matching Excel column letters) ──────

    public const int ColTimestamp       = 1;   // A — UTC timestamp of the scan
    public const int ColBarcodeRow      = 2;   // B — row number in main sheet (for lookup)
    public const int ColEpcHex          = 3;   // C — raw EPC as hex string
    public const int ColTid             = 4;   // D — TID memory bank hex string (chip identity)
    public const int ColScheme          = 5;   // E — e.g. "SGTIN-96"
    public const int ColCompanyPrefix   = 6;   // F — GS1 Company Prefix
    public const int ColRfidGtin14      = 7;   // G — GTIN-14 from EPC
    public const int ColRfidSerial      = 8;   // H — Serial from EPC
    public const int ColBarcodeGtin14   = 9;   // I — GTIN-14 from barcode (AI 01)
    public const int ColBarcodeSerial   = 10;  // J — Serial from barcode (AI 21)
    public const int ColGtin14Match     = 11;  // K — TRUE / FALSE
    public const int ColSerialMatch     = 12;  // L — TRUE / FALSE / N/A
    public const int ColGcpValid        = 13;  // M — TRUE / FALSE / N/A
    public const int ColValidationStatus= 14;  // N — Pass / Fail / NoTag / ParseError / etc.
    public const int ColScanWindowMs    = 15;  // O — actual scan window duration
    public const int ColMismatchDetail  = 16;  // P — semicolon-separated mismatch field list
    public const int ColTagCount        = 17;  // Q — number of distinct EPCs detected

    public const int TotalColumns       = 17;

    /// <summary>Column headers in display order (index = ColXxx - 1).</summary>
    public static readonly string[] Headers =
    [
        "Timestamp (UTC)",
        "Barcode Row",
        "EPC (Hex)",
        "TID (Hex)",
        "EPC Scheme",
        "Company Prefix",
        "RFID GTIN-14",
        "RFID Serial",
        "Barcode GTIN-14",
        "Barcode Serial",
        "GTIN-14 Match",
        "Serial Match",
        "GCP Valid",
        "Validation Status",
        "Scan Window (ms)",
        "Mismatch Detail",
        "Tags Detected",
    ];

    /// <summary>
    /// Preferred column widths in Excel character units.
    /// Index matches <see cref="Headers"/>.
    /// </summary>
    public static readonly double[] ColumnWidths =
    [
        22,  // Timestamp
        11,  // Barcode Row
        32,  // EPC Hex
        28,  // TID Hex
        12,  // Scheme
        16,  // Company Prefix
        18,  // RFID GTIN-14
        20,  // RFID Serial
        18,  // Barcode GTIN-14
        20,  // Barcode Serial
        13,  // GTIN-14 Match
        12,  // Serial Match
        10,  // GCP Valid
        18,  // Validation Status
        16,  // Scan Window
        50,  // Mismatch Detail
        13,  // Tags Detected
    ];
}
