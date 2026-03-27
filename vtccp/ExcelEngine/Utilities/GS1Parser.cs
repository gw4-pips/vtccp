namespace ExcelEngine.Utilities;

/// <summary>
/// Minimal GS1 Application Identifier parser for VTCCP use cases.
///
/// Handles FNC1 as:
///   • Raw GS character (U+001D / ASCII 29) — Cognex DataMan DMCC raw output
///   • Text placeholder "&lt;F1&gt;" — DataMan decoded-data display format
///
/// Parsing strategy:
///   The parser forward-scans using a built-in table of fixed-length AIs.
///   Fixed-length AIs (e.g. AI 01 = 14 data chars) are consumed by length so
///   the scanner can correctly step past them even when no FNC1 separator
///   appears between consecutive fixed-length elements.
///   Variable-length AIs not present in the table are terminated by the next
///   FNC1 or by the end of the string.
///
/// Phase 1 limitation:
///   The fixed-length table covers the AIs most common in healthcare/pharma
///   GS1 DataMatrix labelling.  4-char AIs in the 31xx–36xx measurement range
///   are handled as a group (all map to 6 data chars).  Uncommon 3-char and
///   4-char AIs outside this range are treated as variable-length.
/// </summary>
public static class GS1Parser
{
    private const char Gs = '\u001d';           // raw GS / FNC1 (ASCII 29)
    private const string FncPlaceholder = "<F1>";

    // Data lengths (chars of DATA only, not including the AI digits themselves).
    // Source: GS1 General Specifications v24, Table 7-1.
    private static readonly Dictionary<string, int> FixedDataLen =
        new(StringComparer.Ordinal)
    {
        // ── 2-char AIs ──────────────────────────────────────────────────────
        ["00"] = 18, // SSCC
        ["01"] = 14, // GTIN-14
        ["02"] = 14, // GTIN of contained trade item
        ["03"] = 14,
        ["04"] = 16,
        ["11"] =  6, // Production date (YYMMDD)
        ["12"] =  6, // Due date
        ["13"] =  6, // Packaging date
        ["15"] =  6, // Best before date
        ["16"] =  6, // Sell-by date
        ["17"] =  6, // Expiry date
        ["18"] =  6, // Harvest date
        ["19"] =  6,
        ["20"] =  2, // Product variant
        // ── 4-char AIs: 3100–3699 (all have 6 data chars) ───────────────────
        // Handled below via prefix matching rather than explicit enumeration.
    };

    // 4-char AI prefixes whose data length is known (all = 6 chars).
    private static readonly string[] FourCharPrefixes6 =
        ["31", "32", "33", "34", "35", "36"];

    /// <summary>
    /// Extract the value of a specific Application Identifier from a GS1 string.
    /// Returns <see langword="null"/> if the AI is absent or the input is empty.
    /// </summary>
    public static string? ExtractAI(string? encodedData, string aiCode)
    {
        if (string.IsNullOrEmpty(encodedData) || string.IsNullOrEmpty(aiCode))
            return null;

        // Normalise FNC1 representations to raw GS.
        var data = encodedData.Replace(FncPlaceholder, Gs.ToString(),
                                       StringComparison.Ordinal);

        int pos = 0;
        int len = data.Length;

        // Skip optional leading FNC1.
        if (pos < len && data[pos] == Gs) pos++;

        while (pos < len)
        {
            // Skip any inline FNC1 separators.
            while (pos < len && data[pos] == Gs) pos++;
            if (pos >= len) break;

            // ── Identify the AI at the current position ──────────────────────
            // Try 4-char first, then 2-char.
            // (GS1 doesn't use 3-char AIs in the healthcare/pharma space covered here.)

            string? ai      = null;
            int     aiLen   = 0;
            int     dataLen = 0;

            // 4-char AI attempt.
            if (pos + 4 <= len)
            {
                var candidate4 = data.Substring(pos, 4);
                if (IsAllDigits(candidate4) && TryGetFixedDataLen4(candidate4, out dataLen))
                {
                    ai = candidate4; aiLen = 4;
                }
            }

            // 2-char AI attempt (only if 4-char didn't match).
            if (ai is null && pos + 2 <= len)
            {
                var candidate2 = data.Substring(pos, 2);
                if (IsAllDigits(candidate2))
                {
                    ai = candidate2; aiLen = 2;
                }
            }

            if (ai is null) break; // Cannot identify AI — stop.

            // ── Check if this is the target AI ───────────────────────────────
            if (ai == aiCode)
            {
                // Variable-length: read to next GS or end of string.
                int start = pos + aiLen;
                int end   = data.IndexOf(Gs, start);
                return end < 0
                    ? data[start..]
                    : data[start..end];
            }

            // ── Advance past this AI + its data ──────────────────────────────
            dataLen = 0;
            bool fixed4 = ai.Length == 4 && TryGetFixedDataLen4(ai, out dataLen);
            bool fixed2 = !fixed4 && FixedDataLen.TryGetValue(ai, out dataLen);

            if (fixed4 || fixed2)
            {
                pos += aiLen + dataLen;
            }
            else
            {
                // Variable-length: advance to next GS or end.
                int nextGs = data.IndexOf(Gs, pos + aiLen);
                pos = nextGs < 0 ? len : nextGs;
            }
        }

        return null;
    }

    /// <summary>Extract AI(10) Batch/Lot Number. Returns <see langword="null"/> if not present.</summary>
    public static string? ExtractBatchLot(string? encodedData) =>
        ExtractAI(encodedData, "10");

    // ── Private helpers ───────────────────────────────────────────────────────

    private static bool IsAllDigits(string s)
    {
        foreach (char c in s)
            if (c < '0' || c > '9') return false;
        return true;
    }

    private static bool TryGetFixedDataLen4(string ai4, out int dataLen)
    {
        // 31xx–36xx: all have 6 data chars.
        if (ai4.Length == 4)
        {
            var prefix2 = ai4[..2];
            foreach (var p in FourCharPrefixes6)
            {
                if (prefix2 == p) { dataLen = 6; return true; }
            }
        }
        dataLen = 0;
        return false;
    }
}
