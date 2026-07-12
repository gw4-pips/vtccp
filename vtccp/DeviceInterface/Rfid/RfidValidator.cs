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
    /// Optional GCP validator. When null, GCP validation is skipped (GcpValid = null in result).
    /// </param>
    public RfidValidator(GcpValidator? gcpValidator = null)
        => _gcpValidator = gcpValidator;

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
        bool? gcpValid = _gcpValidator?.Validate(parsedEpc);

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

        if (barcodeGtin14 is not null && !gtin14Match)
            mismatches.Add($"GTIN14:RFID={rfidGtin14},BC={barcodeGtin14}");

        if (barcodeSerial is not null && rfidSerial is not null && !serialMatch)
            mismatches.Add($"Serial:RFID={rfidSerial},BC={barcodeSerial}");

        if (gcpValid == false)
            mismatches.Add($"GCP:NotRegistered={parsedEpc.CompanyPrefix}");

        // ── Determine overall status ───────────────────────────────────────────
        bool pass = gtin14Match
            && (barcodeSerial is null || serialMatch)
            && gcpValid != false;

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
            GcpValid        = gcpValid,
            ScanWindowMs    = scanWindowMs,
            MismatchDetail  = mismatches.Count > 0 ? string.Join(";", mismatches) : null,
        };
    }

    // ── GS1 AI extraction ──────────────────────────────────────────────────────

    /// <summary>
    /// Extract AI (01) GTIN-14 from a decoded barcode data string.
    /// Handles GS1 DataMatrix (]d2 prefix, FNC1=0x1D) and raw AI strings.
    /// Returns the 14-digit GTIN-14 string, or null if AI (01) is not found.
    /// </summary>
    public static string? ExtractAi01(string? decodedData)
    {
        if (string.IsNullOrEmpty(decodedData)) return null;
        string payload = StripAimId(decodedData);
        return FindFixedLengthAi(payload, "01", 14);
    }

    /// <summary>
    /// Extract AI (21) serial number from a decoded barcode data string.
    /// AI (21) is variable-length, terminated by FNC1 (0x1D) or end-of-data.
    /// Returns null if AI (21) is not present.
    /// </summary>
    public static string? ExtractAi21(string? decodedData)
    {
        if (string.IsNullOrEmpty(decodedData)) return null;
        string payload = StripAimId(decodedData);
        return FindVariableLengthAi(payload, "21");
    }

    /// <summary>Strip the AIM identifier prefix (e.g. "]d2", "]C1") from the data string.</summary>
    private static string StripAimId(string data)
    {
        if (data.Length >= 3 && data[0] == ']')
            return data[3..];
        return data;
    }

    /// <summary>
    /// Find a fixed-length AI value in the GS1 payload.
    /// GS1 data: AI digits immediately followed by value; FNC1 (0x1D) acts as separator.
    /// </summary>
    private static string? FindFixedLengthAi(string payload, string ai, int valueLen)
    {
        // Normalise FNC1 (0x1D) to pipe for simple parsing
        string normalized = payload.Replace('\x1D', '|');
        // Split on | to get individual AI+value segments
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
        string normalized = payload.Replace('\x1D', '|');
        foreach (string segment in normalized.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment.StartsWith(ai, StringComparison.Ordinal) && segment.Length > ai.Length)
            {
                return segment[ai.Length..];
            }
        }
        return null;
    }
}
