namespace ExcelEngine.Models;

/// <summary>
/// Codeword and encodation-analysis data extracted from the firmware push arrays:
///   q.codewordArray          — one entry per codeword (data + ECC combined)
///   q.encodationAnalysisArray — one entry per encoded character/segment
///
/// Confirmed v1.28 device scan (16×36 symbol):
///   codewordArray.length=56   = 32 data + 24 ECC codewords total ✓
///   encodationAnalysisArray.length=33 = EncodedCharacters field ✓
///
/// Written to the "Codeword Values" worksheet by CwValuesSheetWriter.
/// Attached to VerificationRecord.CodewordValues — populated by the push parser.
/// </summary>
public sealed class CodewordValuesData
{
    /// <summary>
    /// Codeword byte values 0–255, indexed 0..(total−1).
    /// Data codewords occupy indices 0..DataCodewords−1;
    /// ECC codewords follow at DataCodewords..Total−1.
    /// Length = DataCodewords + ECCCodewords (total, confirmed by v1.28 cwLen=56).
    /// </summary>
    public required IReadOnlyList<int> Codewords { get; init; }

    /// <summary>
    /// isCorrected flag per codeword (true = this codeword contained an error that
    /// was corrected by the ECC algorithm during decoding).
    /// Same length and indexing as <see cref="Codewords"/>.
    /// </summary>
    public required IReadOnlyList<bool> IsCorrected { get; init; }

    /// <summary>
    /// Number of data codewords in the symbol.
    /// Derived from MatrixSize via the ECC200 size lookup in
    /// <see cref="Ecc200DataCodewordsTable"/>.
    /// Null when MatrixSize is unrecognized (lookup miss).
    /// </summary>
    public int? DataCodewords { get; init; }

    /// <summary>
    /// Number of ECC codewords = Codewords.Count − DataCodewords (when DataCodewords is known).
    /// </summary>
    public int? EccCodewords => DataCodewords.HasValue ? Codewords.Count - DataCodewords.Value : null;

    /// <summary>
    /// Encodation analysis: one entry per encoded character or segment.
    /// Length matches VerificationRecord.EncodedCharacters.
    /// Element shape confirmed v1.28: {name:string, mode:string, result:string}.
    /// </summary>
    public required IReadOnlyList<EncodationEntry> EncodationAnalysis { get; init; }

    /// <summary>Timestamp of the scan (from VerificationRecord.VerificationDateTime).</summary>
    public DateTime ScanDateTime { get; init; }

    /// <summary>
    /// Human-readable label for the sheet section header.
    /// E.g. "2026-05-18 11:24:40  |  Data Matrix  |  A  |  (010)..."
    /// </summary>
    public string? Label { get; init; }

    // ── ECC200 DataCodewords lookup ───────────────────────────────────────────
    //
    // Maps MatrixSize string (as emitted in the push XML) → data codeword count.
    // Source: ISO/IEC 16022 Table 7.
    //
    // Square sizes confirmed from standard; rectangular 16×36=32 DEVICE-CONFIRMED
    // (v1.28 cwLen=56=32+24 ✓).  All other rectangular values from ISO/IEC 16022 —
    // verify against device before relying on them for production use.
    //
    // ECC codewords = total (Codewords.Count) − data.
    //
    public static readonly IReadOnlyDictionary<string, int> Ecc200DataCodewordsTable =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            // ── Square ECC200 ──────────────────────────────────────────────────
            { "10x10",    3 },
            { "12x12",    5 },
            { "14x14",    8 },
            { "16x16",   12 },
            { "18x18",   18 },
            { "20x20",   22 },
            { "22x22",   30 },
            { "24x24",   36 },
            { "26x26",   44 },
            { "32x32",   62 },
            { "36x36",   86 },
            { "40x40",  114 },
            { "44x44",  144 },
            { "48x48",  174 },
            { "52x52",  204 },
            { "64x64",  280 },
            { "72x72",  368 },
            { "80x80",  456 },
            { "88x88",  576 },
            { "96x96",  696 },
            { "104x104", 816 },
            { "120x120", 1050 },
            { "132x132", 1304 },
            { "144x144", 1558 },

            // ── Rectangular ECC200 (16×36 device-confirmed; others from ISO 16022) ──
            { "8x18",    10 },
            { "8x32",    20 },
            { "12x26",   32 },
            { "12x36",   44 },
            { "16x36",   32 },   // DEVICE-CONFIRMED v1.28: cwLen=56, data=32, ECC=24 ✓
            { "16x48",   64 },
        };

    /// <summary>
    /// Look up the data codeword count for the given MatrixSize string.
    /// Returns null if the size is not in the table.
    /// </summary>
    public static int? LookupDataCodewords(string? matrixSize)
    {
        if (string.IsNullOrWhiteSpace(matrixSize)) return null;
        return Ecc200DataCodewordsTable.TryGetValue(matrixSize.Trim(), out var count)
            ? count : null;
    }
}

/// <summary>
/// One entry from the firmware encodationAnalysisArray.
/// Element shape confirmed v1.28: {name:string, mode:string, result:string}.
/// </summary>
/// <param name="Name">Character or segment name/value as reported by firmware.</param>
/// <param name="Mode">Encodation mode (e.g. "ASCII", "C40", "Base256").</param>
/// <param name="Result">Encoding result or status string from firmware.</param>
public sealed record EncodationEntry(string Name, string Mode, string Result);
