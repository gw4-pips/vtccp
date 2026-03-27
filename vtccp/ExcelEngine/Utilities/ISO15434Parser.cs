namespace ExcelEngine.Utilities;

/// <summary>
/// Parses ISO/IEC 15434 (ANSI MH10.8.2) transport-syntax strings to extract
/// batch/lot information from the Data Identifier 4L (primary) or related DIs
/// 10L, 1L, and L.
///
/// Covers all standards that use the ISO 15434 envelope with MH10.8.2 DIs:
///   MIL-STD-129  — Military shipment and storage labels (typically PDF417)
///   MIL-STD-130  — DoD UID item marking (typically Data Matrix)
///   ANSI MH10.8.2 — Commercial supply chain / healthcare
///
/// Envelope format:
///   [)> RS {format-indicator} GS (DI+data GS)+ RS [EOT]
///
/// Common format indicators: 06 = ANSI MH10.8.2, 05 = EDI, 12 = ASC X12.
/// All format indicators are accepted; the indicator value is not validated.
///
/// DataMan-style text escapes accepted in addition to raw control characters:
///   &lt;RS&gt; → 0x1E (Record Separator)
///   &lt;GS&gt; → 0x1D (Group Separator)
///   &lt;EOT&gt; → 0x04 (End of Transmission — optional)
/// </summary>
public static class ISO15434Parser
{
    private const char   RsChar        = '\u001E';
    private const char   GsChar        = '\u001D';
    private const char   EotChar       = '\u0004';
    private const string EnvelopeStart = "[)>";

    /// <summary>
    /// Data Identifiers representing batch/lot, tested in priority order.
    /// ANSI MH10.8.2: 4L = Lot/Batch (primary), 10L = alternate, 1L = lot tracking, L = legacy.
    /// </summary>
    private static readonly string[] BatchDIs = ["4L", "10L", "1L", "L"];

    /// <summary>
    /// Returns the batch/lot value extracted from a 15434 envelope string, or null
    /// if the string is not in 15434 format or contains no recognizable batch DI.
    /// </summary>
    public static string? ExtractBatchLot(string? raw)
    {
        if (raw is null) return null;

        string data   = Normalize(raw);
        int    envIdx = data.IndexOf(EnvelopeStart, StringComparison.Ordinal);
        if (envIdx < 0) return null;

        // Advance past [)> — next char must be RS.
        int pos = envIdx + EnvelopeStart.Length;
        if (pos >= data.Length || data[pos] != RsChar) return null;
        pos++; // past RS

        // Skip format indicator (e.g. "06") + the following GS.
        int gsIdx = data.IndexOf(GsChar, pos);
        if (gsIdx < 0) return null;
        pos = gsIdx + 1; // now positioned at the first DI character

        // Scan GS-delimited DI fields until RS (end-of-record) or EOT or end-of-string.
        while (pos < data.Length && data[pos] != RsChar && data[pos] != EotChar)
        {
            int fieldEnd = MinPositive(data.IndexOf(GsChar, pos),
                                       data.IndexOf(RsChar, pos));
            if (fieldEnd < 0) fieldEnd = data.Length;

            string field = data[pos..fieldEnd];

            foreach (string di in BatchDIs)
            {
                if (field.StartsWith(di, StringComparison.OrdinalIgnoreCase))
                    return field[di.Length..];
            }

            pos = fieldEnd + 1;
        }

        return null;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Replaces DataMan-style text placeholders with the corresponding control characters.
    /// </summary>
    private static string Normalize(string s) =>
        s.Replace("<RS>",  "\u001E")
         .Replace("<GS>",  "\u001D")
         .Replace("<EOT>", "\u0004");

    /// <summary>Returns the smaller of two values, ignoring negatives (not-found sentinels).</summary>
    private static int MinPositive(int a, int b)
    {
        if (a < 0) return b;
        if (b < 0) return a;
        return Math.Min(a, b);
    }
}
