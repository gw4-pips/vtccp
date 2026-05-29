namespace DeviceInterface.Dmst;

/// <summary>
/// Formats barcode decoded-data strings for human display and Excel storage,
/// matching the conventions used by the DataMan Setup Tool (DMST / TruCheck).
///
/// ── Write path (device → display / Excel) ─────────────────────────────────
///   1. DmstResultParser.SubstituteForbiddenXmlChars() runs BEFORE XML parsing and
///      converts every forbidden XML 1.0 byte to its angle-bracket mnemonic text
///      (&lt;RS&gt;, &lt;EOT&gt;, etc.) except GS (0x1D) which becomes "|".  After parsing,
///      those text tokens are already in DecodedData — SubstituteControlChars below
///      acts only as a safety net for any literal byte that somehow survived.
///   2. Any remaining literal control chars in DecodedData are replaced with
///      mnemonics: &lt;ETX&gt;, &lt;EOT&gt;, &lt;RS&gt;, etc.
///   3. "&lt;F1&gt;" is prepended for GS1 symbologies (AIM ID indicates FNC1 in first
///      position), and every "|" (GS/FNC1 separator) becomes "&lt;F1&gt;" too.
///
/// ── Read-back path (Excel → IMAGE.LOAD / round-trip) ─────────────────────
///   Reverses step 2 (strip "&lt;F1&gt;") and step 1 (restore mnemonics to bytes),
///   then replaces "|" with 0x1D so the payload is suitable for device replay.
/// </summary>
public static class BarcodeDataFormatter
{
    // ── GS1 FNC1-in-first-position AIM IDs ───────────────────────────────────
    // Each entry is the full three-character AIM ID string as it appears in
    // the SymbologyId push XML element.  Only identifiers where the barcode
    // standard places FNC1 as the very first encoded character are listed;
    // FNC1-in-second-position variants are deliberately excluded.
    private static readonly HashSet<string> _fnc1FirstAimIds = new(StringComparer.Ordinal)
    {
        "]d2",   // GS1 DataMatrix ECC200 (FNC1 in first position)
        "]C1",   // GS1-128 / Code 128 with FNC1 in first position
        "]Q3",   // GS1 QR Code (FNC1 in first position, modifier 3)
        "]J3",   // GS1 DotCode
        "]e0",   // GS1-128 legacy AIM designation
    };

    // ── Control-character substitution table ──────────────────────────────────
    // Maps literal control bytes to their conventional angle-bracket mnemonics.
    // 0x09 (TAB), 0x0A (LF), 0x0D (CR) are permitted in XML and left as-is.
    // 0x1D (GS) is already "|" after DmstResultParser pre-parse sanitisation.
    private static readonly (char Char, string Mnemonic)[] _controlMnemonics =
    [
        ('\x01', "<SOH>"),
        ('\x02', "<STX>"),
        ('\x03', "<ETX>"),
        ('\x04', "<EOT>"),
        ('\x05', "<ENQ>"),
        ('\x06', "<ACK>"),
        ('\x07', "<BEL>"),
        ('\x08', "<BS>"),
        // 0x09 TAB — left as-is
        // 0x0A LF  — left as-is
        ('\x0B', "<VT>"),
        ('\x0C', "<FF>"),
        // 0x0D CR  — left as-is
        ('\x0E', "<SO>"),
        ('\x0F', "<SI>"),
        ('\x10', "<DLE>"),
        ('\x11', "<DC1>"),
        ('\x12', "<DC2>"),
        ('\x13', "<DC3>"),
        ('\x14', "<DC4>"),
        ('\x15', "<NAK>"),
        ('\x16', "<SYN>"),
        ('\x17', "<ETB>"),
        ('\x18', "<CAN>"),
        ('\x19', "<EM>"),
        ('\x1A', "<SUB>"),
        ('\x1B', "<ESC>"),
        ('\x1C', "<FS>"),
        // 0x1D GS — already "|"; omitted here so pipe is not double-converted
        ('\x1E', "<RS>"),
        ('\x1F', "<US>"),
    ];

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Formats a raw decoded-data string (after DmstResultParser's GS→| pass)
    /// for display and Excel storage.
    ///
    /// Remaining literal control bytes are replaced with mnemonics such as
    /// "&lt;ETX&gt;", "&lt;EOT&gt;", "&lt;RS&gt;".  For GS1 FNC1-in-first-position
    /// symbologies the string is prefixed with "&lt;F1&gt;", matching DMST/TruCheck
    /// display conventions.  Pipe characters (GS1 AI separators) are left as-is.
    ///
    /// All resulting characters are Excel-safe (no control bytes in the cell value).
    /// </summary>
    /// <param name="rawData">
    ///   DecodedData after the GS→| pre-parse substitution in DmstResultParser.
    ///   May be null or empty.
    /// </param>
    /// <param name="aimId">
    ///   AIM ID from the push XML &lt;SymbologyId&gt; element (e.g. "]d2", "]C1").
    ///   Null is treated as unknown — no &lt;F1&gt; prefix will be added.
    /// </param>
    public static string FormatForDisplay(string? rawData, string? aimId)
    {
        if (string.IsNullOrEmpty(rawData))
            return rawData ?? string.Empty;

        string cleaned = SubstituteControlChars(rawData);

        if (!IsGS1FNC1First(aimId))
            return cleaned;

        // In GS1 DataMatrix (and GS1-128, GS1 QR, etc.) the FNC1 codeword (232 / 0xE8)
        // serves as both the initial GS1 indicator and every subsequent field separator.
        // DmstResultParser converts FNC1/GS bytes (0x1D) to "|" before XML parsing.
        // Replace all of those pipes with <F1> to match DMST TruCheck display output,
        // e.g. <F1>0100355513710213<F1>2110000328934717<F1>280331<F1>101197170
        return "<F1>" + cleaned.Replace("|", "<F1>", StringComparison.Ordinal);
    }

    /// <summary>
    /// Reverses <see cref="FormatForDisplay"/>: strips the "&lt;F1&gt;" prefix (if
    /// any) and restores mnemonic sequences to their original byte values, then
    /// replaces "|" with GS (0x1D).  Use this to reconstruct the raw payload for
    /// IMAGE.LOAD or D2 round-trip operations.
    /// </summary>
    public static string RestoreFromDisplay(string? formatted, string? aimId)
    {
        if (string.IsNullOrEmpty(formatted))
            return formatted ?? string.Empty;

        string result = formatted;

        // All <F1> tokens (leading indicator and internal separators) represent
        // FNC1/GS — collapse them all to | first, then normalise to 0x1D below.
        result = result.Replace("<F1>", "|", StringComparison.Ordinal);

        foreach (var (ch, mnemonic) in _controlMnemonics)
            result = result.Replace(mnemonic, ch.ToString(), StringComparison.Ordinal);

        result = result.Replace("|", "\x1D", StringComparison.Ordinal);

        return result;
    }

    // ── Internal helpers ──────────────────────────────────────────────────────

    private static bool IsGS1FNC1First(string? aimId) =>
        aimId is not null && _fnc1FirstAimIds.Contains(aimId);

    private static string SubstituteControlChars(string s)
    {
        // Fast path: most barcode payloads contain no remaining control bytes.
        bool hasControl = false;
        foreach (char c in s)
        {
            if (c < '\x20' && c != '\x09' && c != '\x0A' && c != '\x0D')
            { hasControl = true; break; }
        }
        if (!hasControl) return s;

        var sb = new System.Text.StringBuilder(s.Length + 32);
        foreach (char c in s)
        {
            if (c >= '\x20' || c == '\x09' || c == '\x0A' || c == '\x0D')
            {
                sb.Append(c);
                continue;
            }
            string? mnemonic = null;
            foreach (var (ch, m) in _controlMnemonics)
                if (c == ch) { mnemonic = m; break; }
            sb.Append(mnemonic ?? $"<0x{(int)c:X2}>");
        }
        return sb.ToString();
    }
}
