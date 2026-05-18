namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;

/// <summary>
/// Writes scan JPEG images to the "Images" worksheet.
///
/// Sheet layout per record block:
///   Row N+0: Label row — "yyyy-MM-dd HH:mm:ss | Symbology | Grade | DecodedData (first 40 chars)"
///            bold, col 1. Col 2 holds the full base64 string for D2 reverse-report round-trip.
///   Row N+1: Image row — tall (160pt); JPEG embedded anchored at col 1.
///   Row N+2: blank separator
///   Row N+3: blank separator
///
/// Multiple records are appended sequentially in the same sheet.
///
/// IMPORTANT: The caller (ExcelWriter) must call adapter.EnsureSheet("Main") after
/// WriteRecord returns to restore the Main sheet as the active write target.
///
/// The base64 string in col 2 of the label row is intentionally wide — Excel cells support
/// up to 32,767 characters; scan images from the DM475V are ≤ ~400 KB base64 (~533,000 chars),
/// which exceeds that limit for large symbols. In practice, QR v3 = ~22,000 chars (safe) and
/// DM 24×24 = ~18,000 chars (safe). If a base64 payload exceeds 32,767 chars it is truncated
/// with a trailing "…[TRUNCATED]" marker so the cell is still readable; the sidecar (SessionSidecar)
/// always stores the full payload and is the authoritative source for D2 round-trip.
/// </summary>
public sealed class ImagesSheetWriter
{
    public const string SheetName = "Images";

    private const double ImageRowHeightPt   = 160.0;  // ~213px — fits a ~220×220 scan JPEG
    private const int    MaxBase64CellChars = 32_000;  // Excel cell limit is 32,767; leave headroom

    private readonly IExcelAdapter _adapter;
    private int  _nextRow;
    private bool _sheetEnsured;

    public ImagesSheetWriter(IExcelAdapter adapter)
    {
        _adapter      = adapter;
        _nextRow      = 1;
        _sheetEnsured = false;
    }

    /// <summary>
    /// Write one record's JPEG image block to the "Images" sheet.
    /// No-ops if <paramref name="record"/>.JpegImageBase64 is null.
    /// Switches the adapter's active sheet to "Images" for the duration.
    /// The caller must call adapter.EnsureSheet("Main") after this returns.
    /// </summary>
    public void WriteRecord(VerificationRecord record)
    {
        if (record.JpegImageBase64 is null)
            return;

        EnsureSheetActive();

        // ── Label row ─────────────────────────────────────────────────────────
        string dtStr  = record.VerificationDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        string sym    = record.Symbology;
        string grade  = record.OverallGrade?.ToString() ?? "-";
        string data   = record.DecodedData is not null
                            ? (record.DecodedData.Length > 40
                               ? record.DecodedData[..40] + "\u2026"   // "…"
                               : record.DecodedData)
                            : "";
        string label  = $"{dtStr}  |  {sym}  |  {grade}  |  {data}";

        _adapter.WriteString(_nextRow, 1, label);
        _adapter.SetCellBold(_nextRow, 1);

        // Store base64 in col 2 — truncate if it would exceed the Excel cell limit.
        // The full payload lives in the sidecar; this is a best-effort convenience copy.
        string b64Cell = record.JpegImageBase64.Length <= MaxBase64CellChars
                            ? record.JpegImageBase64
                            : record.JpegImageBase64[..MaxBase64CellChars] + "[TRUNCATED]";
        _adapter.WriteString(_nextRow, 2, b64Cell);
        _nextRow++;

        // ── Image row ─────────────────────────────────────────────────────────
        _adapter.SetRowHeight(_nextRow, ImageRowHeightPt);

        byte[] jpegBytes;
        try
        {
            jpegBytes = Convert.FromBase64String(record.JpegImageBase64);
        }
        catch (FormatException ex)
        {
            // Malformed base64 payload (truncated push, encoding error, etc.).
            // Write an error marker in the image cell and skip embedding so the
            // rest of record append continues normally.
            _adapter.WriteString(_nextRow, 1, $"[IMAGE DECODE ERROR: {ex.Message}]");
            _nextRow += 3;   // image row + two blank separator rows
            return;
        }

        _adapter.WriteEmbeddedImage(_nextRow, 1, jpegBytes);
        _nextRow++;

        // ── Two-row blank separator between records ───────────────────────────
        _nextRow += 2;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void EnsureSheetActive()
    {
        int existingRows = _adapter.EnsureSheet(SheetName);
        if (!_sheetEnsured)
        {
            _nextRow      = existingRows > 0 ? existingRows + 2 : 1;
            _sheetEnsured = true;
        }
    }
}
