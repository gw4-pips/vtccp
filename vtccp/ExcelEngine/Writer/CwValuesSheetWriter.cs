namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;

/// <summary>
/// Writes codeword-array and encodation-analysis data to the "Codeword Values" worksheet.
///
/// Sheet layout per record block:
///   Row N+0:    Section label (datetime | sym | grade | size | CW counts)
///               Bold; VCCS-navy background.
///   Row N+1:    Codeword table column header: Index | Codeword | IsCorrected | Type
///               Bold; sub-header blue background.
///   Row N+2+:   Codeword data rows.
///               A divider row is inserted at the data/ECC boundary (when DataCodewords is known).
///               Data rows:      pale-blue background.
///               ECC rows:       pale-yellow background.
///               Corrected rows: pale-red background (overrides data/ECC colour).
///   Blank row:  separates codeword table from encodation table.
///   Encodation header: # | Name | Mode | Result  (bold; sub-header background)
///   Encodation rows:   one row per EncodationEntry.
///   Two blank separator rows between records.
///
/// IMPORTANT: The caller (ExcelWriter) must call adapter.EnsureSheet("Main") after
/// WriteRecord returns to restore Main as the active write target.
/// </summary>
public sealed class CwValuesSheetWriter
{
    public const string SheetName = "Codeword Values";

    private const uint ArgbSectionHeader = 0xFF1E3A5F;   // VCCS navy
    private const uint ArgbSubHeader     = 0xFF2E5480;   // lighter navy
    private const uint ArgbDataRow       = 0xFFEBF5FB;   // pale blue (data codewords)
    private const uint ArgbEccRow        = 0xFFFFF9C4;   // pale yellow (ECC codewords)
    private const uint ArgbCorrected     = 0xFFFFCDD2;   // pale red (corrected error)
    private const uint ArgbBoundary      = 0xFFB0BEC5;   // grey (data/ECC divider marker)

    private const int ColCount = 4;   // Index | Codeword | IsCorrected | Type

    private readonly IExcelAdapter _adapter;
    private int  _nextRow;
    private bool _sheetEnsured;

    public CwValuesSheetWriter(IExcelAdapter adapter)
    {
        _adapter      = adapter;
        _nextRow      = 1;
        _sheetEnsured = false;
    }

    /// <summary>
    /// Write one record's codeword and encodation analysis tables to the sheet.
    /// No-ops if <paramref name="record"/>.CodewordValues is null.
    /// Switches the adapter's active sheet to "Codeword Values" for the duration.
    /// The caller must call adapter.EnsureSheet("Main") after this returns.
    /// </summary>
    public void WriteRecord(VerificationRecord record)
    {
        var data = record.CodewordValues;
        if (data is null) return;

        EnsureSheetActive();

        int dataCW = data.DataCodewords ?? 0;
        int eccCW  = data.EccCodewords  ?? (data.Codewords.Count - dataCW);

        // ── Section label row ─────────────────────────────────────────────────
        string dtStr  = data.ScanDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        string grade  = record.OverallGrade?.ToString() ?? "-";
        string size   = record.MatrixSize ?? "-";
        string sym    = record.Symbology;
        string label  = $"{dtStr}  |  {sym}  |  {grade}  |  {size}" +
                        $"  |  total={data.Codewords.Count}  data={dataCW}  ECC={eccCW}";

        _adapter.WriteString(_nextRow, 1, label);
        _adapter.SetCellBold(_nextRow, 1);
        _adapter.SetRowBackground(_nextRow, ColCount, ArgbSectionHeader);
        _nextRow++;

        // ── Codeword table ────────────────────────────────────────────────────
        WriteSubHeader("Index", "Codeword", "IsCorrected", "Type");

        for (int i = 0; i < data.Codewords.Count; i++)
        {
            // Boundary marker row at data/ECC split
            if (data.DataCodewords.HasValue && i == data.DataCodewords.Value)
            {
                _adapter.WriteString(_nextRow, 1,
                    $"─── ECC codewords start at index {i} ───");
                _adapter.SetRowBackground(_nextRow, ColCount, ArgbBoundary);
                _nextRow++;
            }

            bool isCorrected = i < data.IsCorrected.Count && data.IsCorrected[i];
            bool isEcc       = data.DataCodewords.HasValue && i >= data.DataCodewords.Value;
            uint bgColor     = isCorrected ? ArgbCorrected
                             : isEcc       ? ArgbEccRow
                             :               ArgbDataRow;

            _adapter.WriteNumber(_nextRow, 1, i, null);
            _adapter.WriteNumber(_nextRow, 2, data.Codewords[i], null);
            _adapter.WriteString(_nextRow, 3, isCorrected ? "1" : "0");
            _adapter.WriteString(_nextRow, 4, isEcc ? "ECC" : "Data");
            _adapter.SetRowBackground(_nextRow, ColCount, bgColor);
            _nextRow++;
        }

        _nextRow++;   // blank row between codeword and encodation tables

        // ── Encodation analysis table ─────────────────────────────────────────
        if (data.EncodationAnalysis.Count > 0)
        {
            WriteSubHeader("#", "Name", "Mode", "Result");

            for (int j = 0; j < data.EncodationAnalysis.Count; j++)
            {
                var ea = data.EncodationAnalysis[j];
                _adapter.WriteNumber(_nextRow, 1, j, null);
                _adapter.WriteString(_nextRow, 2, ea.Name);
                _adapter.WriteString(_nextRow, 3, ea.Mode);
                _adapter.WriteString(_nextRow, 4, ea.Result);
                _nextRow++;
            }
        }

        // ── Two-row blank separator between records ───────────────────────────
        _nextRow += 2;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void WriteSubHeader(string c1, string c2, string c3, string c4)
    {
        _adapter.WriteString(_nextRow, 1, c1);
        _adapter.WriteString(_nextRow, 2, c2);
        _adapter.WriteString(_nextRow, 3, c3);
        _adapter.WriteString(_nextRow, 4, c4);
        _adapter.SetRowBold(_nextRow, ColCount);
        _adapter.SetRowBackground(_nextRow, ColCount, ArgbSubHeader);
        _nextRow++;
    }

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
