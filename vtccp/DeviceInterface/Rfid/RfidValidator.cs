using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// Cross-validates RFID EPC data against the barcode <see cref="VerificationRecord"/>.
///
/// Validation checks:
///   1. EPC parse: did the raw EPC decode successfully?
///   2. GTIN-14 match: does the EPC-derived GTIN-14 match AI (01) from the barcode?
///   3. Serial match: does the EPC serial match AI (21) from the barcode (when present)?
///   4. GCP validity: is the company prefix registered in the GS1 GCP table?
///
/// All comparisons are exact-string after normalisation (no whitespace, leading zeros preserved).
/// </summary>
public sealed class RfidValidator
{
    private readonly GcpValidator? _gcpValidator;

    /// <param name="gcpValidator">
    /// Optional GCP validator. When null, GCP validation is skipped.
    /// </param>
    public RfidValidator(GcpValidator? gcpValidator = null)
        => _gcpValidator = gcpValidator;

    /// <summary>
    /// Determines whether the native TruCheck GS1 parser supplied a usable
    /// validation result for the record and, when it did, whether it failed.
    ///
    /// Only the verbatim Data Format Check scraped from the correlated TruCheck
    /// HTML is authoritative here. Barcode quality grades and application-pass
    /// values must not be used as a proxy for GS1 parser availability.
    /// </summary>
    public static TruCheckValidationAssessment AssessTruCheckValidation(
        VerificationRecord barcodeRecord)
    {
        ArgumentNullException.ThrowIfNull(barcodeRecord);

        DataFormatCheckResult? dfc = barcodeRecord.HtmlDataFormatCheck;
        bool usable = dfc is { Rows.Count: > 0 } &&
            dfc.Overall is OverallPassFail.Pass or OverallPassFail.Fail;

        return new TruCheckValidationAssessment(
            Usable: usable,
            Failed: usable && dfc!.Overall == OverallPassFail.Fail);
    }

    /// <summary>
    /// Produce a <see cref="RfidValidationResult"/> from the raw RFID reads and the
    /// barcode record that triggered the scan.
    /// </summary>
    public RfidValidationResult Validate(
        IReadOnlyList<EpcReadResult> reads,
        VerificationRecord barcodeRecord,
        int scanWindowMs)
    {
        // ── No tag detected ────────────────────────────────────────────────────
        if (reads.Count == 0)
        {
            return new RfidValidationResult
            {
                Status       = RfidValidationStatus.NoTag,
                RawReads     = reads,
                ScanWindowMs = scanWindowMs,
            };
        }

        // ── Multiple distinct EPCs ─────────────────────────────────────────────
        EpcReadResult? selected;
        if (reads.Count > 1)
        {
            // Keep the first; flag ambiguity
            selected = reads[0];
        }
        else
        {
            selected = reads[0];
        }

        // ── Parse EPC ─────────────────────────────────────────────────────────
        var parsedEpc = EpcParser.Parse(selected.EpcBytes);
        if (parsedEpc.Scheme == EpcScheme.Unknown || parsedEpc.Gtin14 is null)
        {
            return new RfidValidationResult
            {
                Status       = RfidValidationStatus.ParseError,
                RawReads     = reads,
                SelectedRead = selected,
                ParsedEpc    = parsedEpc,
                ScanWindowMs = scanWindowMs,
            };
        }

        // ── Extract barcode GS1 fields ─────────────────────────────────────────
        string? barcodeGtin14 = ExtractAi01(barcodeRecord.DecodedData);
        string? barcodeSerial = ExtractAi21(barcodeRecord.DecodedData);

        // ── GCP validation ─────────────────────────────────────────────────────
        GcpValidationStatus gcpStatus = _gcpValidator?.Validate(parsedEpc)
            ?? GcpValidationStatus.NotChecked;
        int? gcpRegisteredLength = _gcpValidator?.TryGetRegisteredLength(
            parsedEpc,
            out int registeredLength) == true
            ? registeredLength
            : null;

        // ── Compare GTIN-14 ────────────────────────────────────────────────────
        string rfidGtin14 = parsedEpc.Gtin14!;
        bool gtin14Match  = barcodeGtin14 is not null && rfidGtin14 == barcodeGtin14;

        // ── Compare Serial ─────────────────────────────────────────────────────
        string? rfidSerial = parsedEpc.Serial;
        bool serialMatch   = barcodeSerial is not null
                          && rfidSerial is not null
                          && rfidSerial == barcodeSerial;

        // ── Build mismatch detail ──────────────────────────────────────────────
        var mismatches = new List<string>();

        if (barcodeGtin14 is null)
            mismatches.Add("GTIN14:NoBarcodeData");
        else if (!gtin14Match)
            mismatches.Add($"GTIN14:RFID={rfidGtin14},BC={barcodeGtin14}");

        if (barcodeSerial is not null && rfidSerial is null)
            mismatches.Add("Serial:MissingFromTag");
        else if (barcodeSerial is not null && rfidSerial is not null && !serialMatch)
            mismatches.Add($"Serial:RFID={rfidSerial},BC={barcodeSerial}");

        // GCP registration status is informational — matching GTIN and Serial is
        // sufficient for a PASS. GCP status is retained on the result for display.

        // ── Determine overall status ───────────────────────────────────────────
        bool pass = gtin14Match
            && (barcodeSerial is null || serialMatch);

        RfidValidationStatus status = reads.Count > 1
            ? RfidValidationStatus.MultipleTagsDetected
            : pass ? RfidValidationStatus.Pass : RfidValidationStatus.Fail;

        return new RfidValidationResult
        {
            Status          = status,
            RawReads        = reads,
            SelectedRead    = selected,
            ParsedEpc       = parsedEpc,
            RfidGtin14      = rfidGtin14,
            BarcodeGtin14   = barcodeGtin14,
            RfidSerial      = rfidSerial,
            BarcodeSerial   = barcodeSerial,
            GcpStatus       = gcpStatus,
            GcpRegisteredLength = gcpRegisteredLength,
            ScanWindowMs    = scanWindowMs,
            MismatchDetail  = mismatches.Count > 0 ? string.Join(";", mismatches) : null,
        };
    }

    // ── GS1 AI extraction ──────────────────────────────────────────────────────

    /// <summary>
    /// Extract AI (01) GTIN-14 from a decoded barcode data string.
    /// Handles both GS1 Digital Link URLs (https://.../01/GTIN14/21/Serial)
    /// and GS1 Element String format (]d2 / FNC1-separated AIs).
    /// Returns the 14-digit GTIN-14 string, or null if AI (01) is not found.
    /// </summary>
    public static string? ExtractAi01(string? decodedData)
    {
        if (string.IsNullOrEmpty(decodedData)) return null;
        string payload = StripAimId(decodedData);
        if (IsDigitalLinkUrl(payload))
            return ExtractDlAi(payload, "01");
        return FindFixedLengthAi(payload, "01", 14);
    }

    /// <summary>
    /// Extract AI (21) serial number from a decoded barcode data string.
    /// Handles both GS1 Digital Link URLs and GS1 Element String format.
    /// AI (21) is variable-length, terminated by FNC1 (0x1D) or end-of-data.
    /// Returns null if AI (21) is not present.
    /// </summary>
    public static string? ExtractAi21(string? decodedData)
    {
        if (string.IsNullOrEmpty(decodedData)) return null;
        string payload = StripAimId(decodedData);
        if (IsDigitalLinkUrl(payload))
            return ExtractDlAi(payload, "21");
        return FindVariableLengthAi(payload, "21");
    }

    /// <summary>Strip the AIM identifier prefix (e.g. "]d2", "]C1", "]Q1") from the data string.</summary>
    private static string StripAimId(string data)
    {
        if (data.Length >= 3 && data[0] == ']')
            return data[3..];
        return data;
    }

    /// <summary>
    /// Returns true when <paramref name="payload"/> is a GS1 Digital Link URI
    /// (starts with https:// or http://).
    /// </summary>
    private static bool IsDigitalLinkUrl(string payload)
        => payload.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
        || payload.StartsWith("http://",  StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Extract a GS1 AI value from a GS1 Digital Link URI.
    ///
    /// Path-based (most common):
    ///   https://example.com/01/00012345678905/21/ABC123
    ///   Segments are paired: /AI/value[/AI/value…]
    ///
    /// Query-string-based (less common):
    ///   https://example.com/gtin?01=00012345678905&amp;21=ABC123
    ///
    /// Percent-encoding:
    ///   Values may contain percent-encoded characters (e.g. a serial with a
    ///   literal "/" encoded as %2F).  Uri.AbsolutePath preserves %2F in .NET
    ///   Core/.NET 5+ (per RFC 3986 §3.3) so splitting on '/' never
    ///   misinterprets an encoded slash as a path delimiter.  The raw
    ///   segment is then decoded with Uri.UnescapeDataString before it is
    ///   returned, so callers always receive the unencoded value.
    ///
    /// Compressed segments (OUT OF SCOPE):
    ///   GS1 Digital Link v1.3 §7.8 defines a compact binary encoding for
    ///   numeric-only path segments (e.g. /AIdkMQ encodes GTIN+serial in
    ///   a base64url-like scheme).  That encoding is NOT decoded here.
    ///   Compressed DL URIs are rare in the field — no Cognex reader
    ///   firmware tested produces them — and decoding requires implementing
    ///   the full GS1 DL numeric compressor.  If a compressed URI arrives
    ///   this method returns null (no match found), which the validator
    ///   treats as a barcode-field-not-present (no mismatch asserted).
    ///   Add decompression support here if compressed URIs are observed in
    ///   production.
    /// </summary>
    private static string? ExtractDlAi(string url, string ai)
    {
        try
        {
            var uri = new Uri(url);

            // 1. Path-based: /AI/value pairs.
            //    Uri.AbsolutePath preserves percent-encoded characters (e.g.
            //    %2F stays as-is and is NOT treated as a path separator), so
            //    the split is always correct.  Uri.UnescapeDataString then
            //    decodes the value before returning it to the caller.
            string[] segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < segments.Length - 1; i++)
            {
                if (segments[i] == ai)
                    return Uri.UnescapeDataString(segments[i + 1]);
            }

            // 2. Query-string-based: AI=value pairs.
            //    uri.Query preserves percent-encoding; UnescapeDataString
            //    decodes each value before it is returned.
            if (!string.IsNullOrEmpty(uri.Query))
            {
                string query = uri.Query.TrimStart('?');
                foreach (string pair in query.Split('&', StringSplitOptions.RemoveEmptyEntries))
                {
                    int eq = pair.IndexOf('=');
                    if (eq > 0 && pair[..eq] == ai)
                        return Uri.UnescapeDataString(pair[(eq + 1)..]);
                }
            }
        }
        catch { /* malformed URL — fall through to null */ }
        return null;
    }

    /// <summary>
    /// Find a fixed-length AI value in the GS1 Element String payload.
    /// GS1 data: AI digits immediately followed by value; FNC1 (0x1D) acts as separator.
    /// </summary>
    private static string? FindFixedLengthAi(string payload, string ai, int valueLen)
    {
        string normalized = NormalizeElementStringSeparators(payload);
        foreach (string segment in normalized.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment.StartsWith(ai, StringComparison.Ordinal)
                && segment.Length >= ai.Length + valueLen)
            {
                string value = segment.Substring(ai.Length, valueLen);
                if (value.All(char.IsAsciiDigit)) return value;
            }
        }
        return null;
    }

    /// <summary>Find a variable-length AI value terminated by FNC1 or end of data.</summary>
    private static string? FindVariableLengthAi(string payload, string ai)
    {
        string normalized = NormalizeElementStringSeparators(payload);
        foreach (string segment in normalized.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            int position = 0;
            while (position + 2 <= segment.Length)
            {
                if (segment.AsSpan(position).StartsWith(ai, StringComparison.Ordinal)
                    && segment.Length > position + ai.Length)
                {
                    return segment[(position + ai.Length)..];
                }

                // A GS1 separator is not required after a fixed-length AI.
                // Advance over the fixed data so AI(21) can be discovered after
                // a leading AI(01), as in <F1>01006961147042882172803282009.
                if (TryGetFixedLengthAi(segment, position, out int dataLength))
                {
                    position += 2 + dataLength;
                    continue;
                }

                break;
            }
        }
        return null;
    }

    private static bool TryGetFixedLengthAi(string value, int position, out int dataLength)
    {
        dataLength = 0;
        if (position + 2 > value.Length) return false;

        return value.AsSpan(position, 2) switch
        {
            "00" => SetFixedLength(18, out dataLength),
            "01" => SetFixedLength(14, out dataLength),
            "02" => SetFixedLength(14, out dataLength),
            "11" or "12" or "13" or "15" or "16" or "17" => SetFixedLength(6, out dataLength),
            _ => false,
        };
    }

    private static bool SetFixedLength(int length, out int dataLength)
    {
        dataLength = length;
        return true;
    }

    /// <summary>
    /// Normalises the three FNC1/GS representations that can reach VTCCP:
    /// raw ASCII GS (0x1D), DMST/TruCheck's literal &lt;F1&gt; marker, and the
    /// conventional literal &lt;GS&gt; marker. Treating each as the same delimiter
    /// lets Element String extraction work identically for push XML and HTML data.
    /// </summary>
    private static string NormalizeElementStringSeparators(string payload)
        => payload
            .Replace('\x1D', '|')
            .Replace("<F1>", "|", StringComparison.OrdinalIgnoreCase)
            .Replace("<GS>", "|", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// Native TruCheck GS1 parser state used to choose the VeriWedge fallback path.
/// </summary>
public readonly record struct TruCheckValidationAssessment(bool Usable, bool Failed)
{
    /// <summary>
    /// VeriWedge is used only when native TruCheck validation is unavailable or
    /// when its usable validation result failed.
    /// </summary>
    public bool RequiresVeriWedge => !Usable || Failed;
}
