using DeviceInterface.Rfid.Models;
using ExcelEngine.Adapters;
using ExcelEngine.Schema;

namespace DeviceInterface.Rfid;

/// <summary>
/// Writes RFID cross-validation results to the "RFID Scans" auxiliary worksheet.
///
/// IMPORTANT: The caller (ExcelWriter / AppendRecord) must call
/// adapter.EnsureSheet("Main") — or the primary sheet name — after calling any
/// method here, to restore the adapter's active sheet to the main data sheet.
///
/// Pattern mirrors CwValuesSheetWriter / ModValuesSheetWriter in ExcelEngine.
/// </summary>
public sealed class RfidTabWriter
{
    private const string SheetName = "RFID Scans";

    // ARGB colour constants (0xFF = fully opaque alpha)
    private const uint ArgbHeader         = 0xFF4472C4; // blue header (matches main sheet)
    private const uint ArgbPass           = 0x00000000; // no fill (pass = default)
    private const uint ArgbFail           = 0xFFFFCCCC; // light red
    private const uint ArgbNoTag          = 0xFFFFF2CC; // light amber
    private const uint ArgbMultipleTags   = 0xFFFFE0CC; // light orange
    private const uint ArgbParseError     = 0xFFFFD9CC; // salmon
    private const uint ArgbSkipped        = 0xFFF2F2F2; // light grey

    private readonly IExcelAdapter _adapter;
    private int _nextRow = 2;  // row 1 = header; data starts at row 2
    private bool _headersWritten;

    public RfidTabWriter(IExcelAdapter adapter)
        => _adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));

    // ── Public API ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Ensure the "RFID Scans" sheet exists and write the header row on first call.
    /// Safe to call multiple times. Caller must restore the main sheet afterwards.
    /// </summary>
    public void EnsureSheet()
    {
        _adapter.EnsureSheet(SheetName);
        if (_headersWritten) return;
        WriteHeaders();
        _headersWritten = true;
    }

    /// <summary>
    /// Append one RFID validation result row to the sheet.
    /// Caller must restore the main sheet afterwards via adapter.EnsureSheet(mainSheet).
    /// </summary>
    /// <param name="result">Completed validation result for this scan cycle.</param>
    /// <param name="barcodeRowNumber">
    /// 1-based row number in the main verification sheet that triggered this scan.
    /// </param>
    public void AppendResult(RfidValidationResult result, int barcodeRowNumber)
    {
        _adapter.EnsureSheet(SheetName);

        int row = _nextRow;

        // A — Timestamp
        DateTime ts = result.SelectedRead?.ReadTime.UtcDateTime ?? DateTime.UtcNow;
        _adapter.WriteDateTime(row, RfidTabSchema.ColTimestamp, ts, "yyyy-mm-dd hh:mm:ss");

        // B — Barcode Row cross-reference
        _adapter.WriteNumber(row, RfidTabSchema.ColBarcodeRow, barcodeRowNumber);

        // C — EPC hex
        _adapter.WriteString(row, RfidTabSchema.ColEpcHex, result.SelectedRead?.EpcHex ?? "");

        // D — EPC scheme
        _adapter.WriteString(row, RfidTabSchema.ColScheme, FormatScheme(result.ParsedEpc?.Scheme));

        // E — Company prefix
        _adapter.WriteString(row, RfidTabSchema.ColCompanyPrefix, result.ParsedEpc?.CompanyPrefix ?? "");

        // F — RFID GTIN-14
        _adapter.WriteString(row, RfidTabSchema.ColRfidGtin14, result.RfidGtin14 ?? "");

        // G — RFID serial
        _adapter.WriteString(row, RfidTabSchema.ColRfidSerial, result.RfidSerial ?? "");

        // H — Barcode GTIN-14
        _adapter.WriteString(row, RfidTabSchema.ColBarcodeGtin14, result.BarcodeGtin14 ?? "");

        // I — Barcode serial
        _adapter.WriteString(row, RfidTabSchema.ColBarcodeSerial, result.BarcodeSerial ?? "");

        // J — GTIN-14 match
        _adapter.WriteString(row, RfidTabSchema.ColGtin14Match,
            result.BarcodeGtin14 is null ? "N/A" : result.Gtin14Matches ? "TRUE" : "FALSE");

        // K — Serial match
        bool serialComparable = result.BarcodeSerial is not null && result.RfidSerial is not null;
        _adapter.WriteString(row, RfidTabSchema.ColSerialMatch,
            serialComparable ? (result.SerialMatches ? "TRUE" : "FALSE") : "N/A");

        // L — GCP valid
        _adapter.WriteString(row, RfidTabSchema.ColGcpValid,
            result.GcpValid.HasValue ? (result.GcpValid.Value ? "TRUE" : "FALSE") : "N/A");

        // M — Validation status
        _adapter.WriteString(row, RfidTabSchema.ColValidationStatus, result.Status.ToString());

        // N — Scan window (ms)
        _adapter.WriteNumber(row, RfidTabSchema.ColScanWindowMs, result.ScanWindowMs);

        // O — Mismatch detail
        _adapter.WriteString(row, RfidTabSchema.ColMismatchDetail, result.MismatchDetail ?? "");

        // P — Tags detected count
        _adapter.WriteNumber(row, RfidTabSchema.ColTagCount, result.RawReads.Count);

        // Row styling
        uint bgColor = GetRowColor(result.Status);
        if (bgColor != ArgbPass)
            _adapter.SetRowBackground(row, RfidTabSchema.TotalColumns, bgColor);

        _nextRow++;
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private void WriteHeaders()
    {
        _adapter.EnsureSheet(SheetName);

        for (int c = 0; c < RfidTabSchema.Headers.Length; c++)
        {
            int col = c + 1;
            _adapter.WriteString(1, col, RfidTabSchema.Headers[c]);
            if (c < RfidTabSchema.ColumnWidths.Length)
                _adapter.SetColumnWidth(col, RfidTabSchema.ColumnWidths[c]);
        }

        _adapter.SetRowBold(1, RfidTabSchema.TotalColumns);
        _adapter.SetRowBackground(1, RfidTabSchema.TotalColumns, ArgbHeader);
    }

    private static uint GetRowColor(RfidValidationStatus status) => status switch
    {
        RfidValidationStatus.Pass                 => ArgbPass,
        RfidValidationStatus.Fail                 => ArgbFail,
        RfidValidationStatus.NoTag                => ArgbNoTag,
        RfidValidationStatus.MultipleTagsDetected => ArgbMultipleTags,
        RfidValidationStatus.ParseError           => ArgbParseError,
        RfidValidationStatus.Skipped              => ArgbSkipped,
        _                                         => ArgbPass,
    };

    private static string FormatScheme(EpcScheme? scheme) => scheme switch
    {
        EpcScheme.Sgtin96  => "SGTIN-96",
        EpcScheme.Sgtin198 => "SGTIN-198",
        EpcScheme.Unknown  => "Unknown",
        null               => "",
        _                  => scheme.ToString()!,
    };
}
