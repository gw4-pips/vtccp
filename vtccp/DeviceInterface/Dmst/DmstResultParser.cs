namespace DeviceInterface.Dmst;

using System.Xml.Linq;
using ExcelEngine.Models;

/// <summary>
/// Parses a DMST result XML string (from a DMCC GET SYMBOL.RESULT response body,
/// or pushed via a DMST TCP connection) into a <see cref="VerificationRecord"/>.
///
/// The parser is driven by a <see cref="VerificationXmlMap"/> that maps XML element
/// names to record fields — swap the map to support different firmware versions.
///
/// Tolerant design: missing or malformed elements are silently skipped; the resulting
/// record will have null values for those fields rather than throwing.
/// </summary>
public static class DmstResultParser
{
    /// <summary>
    /// Parses the XML and returns a populated <see cref="VerificationRecord"/>,
    /// pre-seeded with the device metadata fields from <paramref name="deviceContext"/>.
    /// </summary>
    /// <param name="xml">Full DMST XML string (with or without the DMCCResponse envelope).</param>
    /// <param name="map">Element-name map; pass <c>null</c> for default Cognex DMV mapping.</param>
    /// <param name="deviceContext">
    /// Optional partial record carrying device metadata (serial, firmware, etc.)
    /// that the device doesn't embed inside every result XML.  Session-level
    /// fields (JobName, OperatorId, BatchNumber) should be supplied here and
    /// will be preserved in the returned record.
    /// </param>
    public static VerificationRecord Parse(
        string                xml,
        VerificationXmlMap?   map           = null,
        VerificationRecord?   deviceContext = null)
    {
        map ??= new VerificationXmlMap();

        XDocument doc;
        try   { doc = XDocument.Parse(xml); }
        catch { return Fallback(deviceContext, "XML parse failed"); }

        // Navigate to the result container element.
        XElement? container =
            doc.Descendants(map.ResultContainer).FirstOrDefault()
            ?? doc.Root;   // fallback: search from root

        if (container is null)
            return Fallback(deviceContext, "ResultContainer not found");

        // ── Helper closures ───────────────────────────────────────────────────
        string? Str(string elem) =>
            container.Descendants(elem).FirstOrDefault()?.Value.Trim()
                is { Length: > 0 } v ? v : null;

        decimal? Dec(string elem) =>
            decimal.TryParse(Str(elem), System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : null;

        int? Int(string elem) =>
            int.TryParse(Str(elem), out var i) ? i : null;

        GradingResult? Grade(string elem)
        {
            string? letter = Str(elem);
            if (letter is null) return null;
            // Letter grade is the element text; numeric may be embedded as attribute.
            XElement? el  = container.Descendants(elem).FirstOrDefault();
            decimal   num = 0m;
            if (el?.Attribute("numeric") is { } a)
                decimal.TryParse(a.Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out num);
            return GradingResult.FromLetterAndNumeric(letter, num, "");
        }

        // ── Symbology ─────────────────────────────────────────────────────────
        string?         symbName   = Str(map.SymbologyName) ?? deviceContext?.Symbology ?? "Unknown";
        SymbologyFamily symbFamily = map.ClassifySymbology(symbName);

        // Normalise DataMan symbology strings to VTCCP canonical form.
        string symbology = NormaliseSymbologyName(symbName);

        // ── Timestamp ─────────────────────────────────────────────────────────
        DateTime verifyDt = DateTime.Now;
        if (Str(map.DateTime) is { } dtStr)
            _ = DateTime.TryParse(dtStr, out verifyDt);

        // ── Overall grade ─────────────────────────────────────────────────────
        string? gradeLetterStr = Str(map.OverallGrade);
        decimal gradeNumeric   = Dec(map.OverallGradeNumeric) ?? 0m;
        string? formalGrade    = Str(map.FormalGrade);

        // Derive pass/fail from letter grade when no explicit element is present.
        // A/B → Pass (above any standard minimum); F → Fail; C/D → NotApplicable
        // (threshold-dependent; device may or may not mark those as pass).
        string derivedPassFail = gradeLetterStr?.Trim().ToUpper() switch
        {
            "A" or "B" => "PASS",
            "F"        => "FAIL",
            _          => "",
        };

        GradingResult? overall = gradeLetterStr is not null
            ? GradingResult.FromLetterAndNumeric(gradeLetterStr, gradeNumeric, derivedPassFail)
            : null;

        // ── Verification settings ─────────────────────────────────────────────
        int?    aperture  = Int(map.ApertureRef);
        int?    wavelength = Int(map.Wavelength);
        string? lighting  = Str(map.Lighting);
        string? standard  = Str(map.Standard);

        // ── 2D parameters ─────────────────────────────────────────────────────
        decimal? uecPct  = Dec(map.UECPercent);
        GradingResult? uecGrade  = Grade(map.UECGrade);
        decimal? scPct   = Dec(map.SCPercent);
        GradingResult? scGrade   = Grade(map.SCGrade);
        GradingResult? modGrade  = Grade(map.MODGrade);
        GradingResult? rmGrade   = Grade(map.RMGrade);
        decimal? anuPct  = Dec(map.ANUPercent);
        GradingResult? anuGrade  = Grade(map.ANUGrade);
        decimal? gnuPct  = Dec(map.GNUPercent);
        GradingResult? gnuGrade  = Grade(map.GNUGrade);
        GradingResult? fpdGrade  = Grade(map.FPDGrade);
        GradingResult? decGrade  = Grade(map.DecodeGrade);
        decimal? agVal   = Dec(map.AGValue);
        GradingResult? agGrade   = Grade(map.AGGrade);

        // ── 2D matrix characteristics ─────────────────────────────────────────
        string? matrixSize        = Str(map.MatrixSize);
        decimal? hBwg             = Dec(map.HorizontalBWG);
        decimal? vBwg             = Dec(map.VerticalBWG);
        int? encodedChars         = Int(map.EncodedCharacters);
        int? totalCw              = Int(map.TotalCodewords);
        int? dataCw               = Int(map.DataCodewords);
        int? ecBudget             = Int(map.ErrorCorrectionBudget);
        int? ecCorrected          = Int(map.ErrorsCorrected);
        int? ecCapUsed            = Int(map.ErrorCapacityUsed);
        string? ecType            = Str(map.ErrorCorrectionType);
        decimal? nomXDim2D        = Dec(map.NominalXDim);
        decimal? ppm              = Dec(map.PixelsPerModule);
        string? contrastUniformity = Str(map.ContrastUniformity);
        string? mrd               = Str(map.MRD);

        ImagePolarity polarity = ImagePolarity.Unknown;
        if (Str(map.ImagePolarity) is { } pol)
            Enum.TryParse(pol.Replace(" ", ""), true, out polarity);

        // ── 2D quiet zones / borders (single-region) ──────────────────────────
        GradingResult? llsGrade = Grade(map.LLSGrade);
        GradingResult? blsGrade = Grade(map.BLSGrade);
        GradingResult? lqzGrade = Grade(map.LQZGrade);
        GradingResult? bqzGrade = Grade(map.BQZGrade);
        GradingResult? tqzGrade = Grade(map.TQZGrade);
        GradingResult? rqzGrade = Grade(map.RQZGrade);

        // ── 2D transition ratios / clock tracks ───────────────────────────────
        decimal? ttrPct       = Dec(map.TTRPercent);
        GradingResult? ttrGrade = Grade(map.TTRGrade);
        decimal? rtrPct       = Dec(map.RTRPercent);
        GradingResult? rtrGrade = Grade(map.RTRGrade);
        GradingResult? tctGrade = Grade(map.TCTGrade);
        GradingResult? rctGrade = Grade(map.RCTGrade);

        // ── 2D quadrant parameters ────────────────────────────────────────────
        GradingResult? ulqzGrade = Grade(map.ULQZGrade);
        GradingResult? urqzGrade = Grade(map.URQZGrade);
        GradingResult? ruqzGrade = Grade(map.RUQZGrade);
        GradingResult? rlqzGrade = Grade(map.RLQZGrade);

        decimal? ulqttrPct  = Dec(map.ULQTTRPercent); GradingResult? ulqttrGrade = Grade(map.ULQTTRGrade);
        decimal? urqttrPct  = Dec(map.URQTTRPercent); GradingResult? urqttrGrade = Grade(map.URQTTRGrade);
        decimal? llqttrPct  = Dec(map.LLQTTRPercent); GradingResult? llqttrGrade = Grade(map.LLQTTRGrade);
        decimal? lrqttrPct  = Dec(map.LRQTTRPercent); GradingResult? lrqttrGrade = Grade(map.LRQTTRGrade);

        decimal? ulqrtrPct  = Dec(map.ULQRTRPercent); GradingResult? ulqrtrGrade = Grade(map.ULQRTRGrade);
        decimal? urqrtrPct  = Dec(map.URQRTRPercent); GradingResult? urqrtrGrade = Grade(map.URQRTRGrade);
        decimal? llqrtrPct  = Dec(map.LLQRTRPercent); GradingResult? llqrtrGrade = Grade(map.LLQRTRGrade);
        decimal? lrqrtrPct  = Dec(map.LRQRTRPercent); GradingResult? lrqrtrGrade = Grade(map.LRQRTRGrade);

        GradingResult? ulqtctGrade = Grade(map.ULQTCTGrade); GradingResult? urqtctGrade = Grade(map.URQTCTGrade);
        GradingResult? llqtctGrade = Grade(map.LLQTCTGrade); GradingResult? lrqtctGrade = Grade(map.LRQTCTGrade);
        GradingResult? ulqrctGrade = Grade(map.ULQRCTGrade); GradingResult? urqrctGrade = Grade(map.URQRCTGrade);
        GradingResult? llqrctGrade = Grade(map.LLQRCTGrade); GradingResult? lrqrctGrade = Grade(map.LRQRCTGrade);

        // ── 1D ISO 15416 summary parameters ──────────────────────────────────
        GradingResult? symbolAnsiGrade = Grade(map.SymbolAnsiGrade);
        decimal? avgEdge    = Dec(map.AvgEdge);
        string?  avgRlRd    = Str(map.AvgRlRd);
        decimal? avgSC      = Dec(map.AvgSC);
        decimal? avgMinEC   = Dec(map.AvgMinEC);
        decimal? avgMOD     = Dec(map.AvgMOD);
        decimal? avgDefect  = Dec(map.AvgDefect);
        string?  avgDcod    = Str(map.AvgDcod);
        decimal? avgDEC     = Dec(map.AvgDEC);
        decimal? avgLQZ     = Dec(map.AvgLQZ);
        decimal? avgRQZ     = Dec(map.AvgRQZ);
        decimal? avgHQZ     = Dec(map.AvgHQZ);
        decimal? avgMinQZ   = Dec(map.AvgMinQZ);
        decimal? bwgPct     = Dec(map.BWGPercent);
        decimal? magnif     = Dec(map.Magnification);
        decimal? ratio      = Dec(map.Ratio);
        decimal? nomXDim1D  = Dec(map.NominalXDim1D);

        // ── Per-scan results (1D) ─────────────────────────────────────────────
        var scanResults = new List<ScanResult1D>();
        XElement? scanContainer = container.Descendants(map.ScanResults).FirstOrDefault();
        if (scanContainer is not null)
        {
            foreach (XElement scan in scanContainer.Elements(map.ScanElement))
            {
                string? numAttr = scan.Attribute(map.ScanNumber)?.Value;
                int.TryParse(numAttr, out int scanNum);

                decimal? ScanDec(string elem) =>
                    decimal.TryParse(scan.Element(elem)?.Value,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;

                scanResults.Add(new ScanResult1D
                {
                    ScanNumber = scanNum > 0 ? scanNum : scanResults.Count + 1,
                    Edge       = ScanDec(map.ScanEdge),
                    SC         = ScanDec(map.ScanSC),
                    MinEC      = ScanDec(map.ScanMinEC),
                    MOD        = ScanDec(map.ScanMOD),
                    Defect     = ScanDec(map.ScanDefect),
                    DEC        = ScanDec(map.ScanDEC),
                    LQZ        = ScanDec(map.ScanLQZ),
                    RQZ        = ScanDec(map.ScanRQZ),
                    HQZ        = ScanDec(map.ScanHQZ),
                });
            }
        }

        // ── Assemble record ───────────────────────────────────────────────────
        return new VerificationRecord
        {
            // Identity
            VerificationDateTime = verifyDt,
            Symbology            = symbology,
            SymbologyFamily      = symbFamily,
            DecodedData          = Str(map.DecodedData),
            FormalGrade          = formalGrade,
            OverallGrade         = overall,

            // Verification settings
            Aperture   = aperture,
            Wavelength = wavelength,
            Lighting   = lighting,
            Standard   = standard,

            // Session context preserved from device context
            OperatorId      = deviceContext?.OperatorId,
            JobName         = deviceContext?.JobName,
            BatchNumber     = deviceContext?.BatchNumber,
            RollNumber      = deviceContext?.RollNumber,
            CompanyName     = deviceContext?.CompanyName,
            ProductName     = deviceContext?.ProductName,
            CustomNote      = deviceContext?.CustomNote,
            User1           = deviceContext?.User1,
            User2           = deviceContext?.User2,
            DeviceSerial    = deviceContext?.DeviceSerial,
            DeviceName      = deviceContext?.DeviceName,
            FirmwareVersion = deviceContext?.FirmwareVersion,
            CalibrationDate = deviceContext?.CalibrationDate,

            // 2D quality
            UEC_Percent     = uecPct,
            UEC_Grade       = uecGrade,
            SC_Percent      = scPct,
            SC_Grade        = scGrade,
            MOD_Grade       = modGrade,
            RM_Grade        = rmGrade,
            ANU_Percent     = anuPct,
            ANU_Grade       = anuGrade,
            GNU_Percent     = gnuPct,
            GNU_Grade       = gnuGrade,
            FPD_Grade       = fpdGrade,
            DECODE_Grade    = decGrade,
            AG_Value        = agVal,
            AG_Grade        = agGrade,

            // 2D matrix
            MatrixSize            = matrixSize,
            HorizontalBWG         = hBwg,
            VerticalBWG           = vBwg,
            EncodedCharacters     = encodedChars,
            TotalCodewords        = totalCw,
            DataCodewords         = dataCw,
            ErrorCorrectionBudget = ecBudget,
            ErrorsCorrected       = ecCorrected,
            ErrorCapacityUsed     = ecCapUsed,
            ErrorCorrectionType   = ecType,
            NominalXDim_2D        = nomXDim2D,
            PixelsPerModule       = ppm,
            ImagePolarity         = polarity,
            ContrastUniformity    = contrastUniformity,
            MRD                   = mrd,

            // 2D quiet zones / borders
            LLS_Grade = llsGrade,
            BLS_Grade = blsGrade,
            LQZ_Grade = lqzGrade,
            BQZ_Grade = bqzGrade,
            TQZ_Grade = tqzGrade,
            RQZ_Grade = rqzGrade,

            // 2D transition ratios / clock tracks
            TTR_Percent = ttrPct, TTR_Grade = ttrGrade,
            RTR_Percent = rtrPct, RTR_Grade = rtrGrade,
            TCT_Grade   = tctGrade,
            RCT_Grade   = rctGrade,

            // 2D quadrant
            ULQZ_Grade = ulqzGrade, URQZ_Grade = urqzGrade,
            RUQZ_Grade = ruqzGrade, RLQZ_Grade = rlqzGrade,

            ULQTTR_Percent = ulqttrPct, ULQTTR_Grade = ulqttrGrade,
            URQTTR_Percent = urqttrPct, URQTTR_Grade = urqttrGrade,
            LLQTTR_Percent = llqttrPct, LLQTTR_Grade = llqttrGrade,
            LRQTTR_Percent = lrqttrPct, LRQTTR_Grade = lrqttrGrade,

            ULQRTR_Percent = ulqrtrPct, ULQRTR_Grade = ulqrtrGrade,
            URQRTR_Percent = urqrtrPct, URQRTR_Grade = urqrtrGrade,
            LLQRTR_Percent = llqrtrPct, LLQRTR_Grade = llqrtrGrade,
            LRQRTR_Percent = lrqrtrPct, LRQRTR_Grade = lrqrtrGrade,

            ULQTCT_Grade = ulqtctGrade, URQTCT_Grade = urqtctGrade,
            LLQTCT_Grade = llqtctGrade, LRQTCT_Grade = lrqtctGrade,
            ULQRCT_Grade = ulqrctGrade, URQRCT_Grade = urqrctGrade,
            LLQRCT_Grade = llqrctGrade, LRQRCT_Grade = lrqrctGrade,

            // 1D
            SymbolAnsiGrade = symbolAnsiGrade,
            Avg_Edge        = avgEdge,
            Avg_RlRd        = avgRlRd,
            Avg_SC          = avgSC,
            Avg_MinEC       = avgMinEC,
            Avg_MOD         = avgMOD,
            Avg_Defect      = avgDefect,
            Avg_DCOD        = avgDcod,
            Avg_DEC         = avgDEC,
            Avg_LQZ         = avgLQZ,
            Avg_RQZ         = avgRQZ,
            Avg_HQZ         = avgHQZ,
            Avg_MinQZ       = avgMinQZ,
            BWG_Percent     = bwgPct,
            Magnification   = magnif,
            Ratio           = ratio,
            NominalXDim_1D  = nomXDim1D,
            ScanResults     = scanResults,
        };
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Normalises DataMan device symbology name strings to the VTCCP canonical forms
    /// used in the schema (e.g. "UPCA", "EAN13", "GS1 DataMatrix").
    /// </summary>
    private static string NormaliseSymbologyName(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "Unknown";
        return raw.Trim() switch
        {
            "UPC-A"        => "UPCA",
            "UPC-E"        => "UPCE",
            "EAN-8"        => "EAN8",
            "EAN-13"       => "EAN13",
            "Code 128"     => "Code128",
            "Code 39"      => "Code39",
            "I 2/5"        => "ITF",
            "QR Code"      => "QRCode",
            "GS1 QR Code"  => "GS1 QRCode",
            "Data Matrix"  => "DataMatrix",
            "Data Matrix ECC 200" => "DataMatrix",
            _              => raw.Trim(),
        };
    }

    // ── Plain-text entry point ────────────────────────────────────────────────

    /// <summary>
    /// Creates a minimal <see cref="VerificationRecord"/> from a single plain-text
    /// push line received via the DataMan Network Client (Format Data → Basic/Standard).
    ///
    /// Only <see cref="VerificationRecord.DecodedData"/> and session-context fields
    /// are populated; all quality-grade fields remain null.
    ///
    /// The line may optionally be tab- or comma-delimited in the order:
    ///   content [, decodeTimeMs [, symbology]]
    /// If it contains no delimiter the whole string is treated as the decoded content.
    /// </summary>
    public static VerificationRecord ParseText(
        string              line,
        VerificationRecord? deviceContext = null)
    {
        // Try to split tab-delimited or comma-delimited fields.
        // DataMan Standard format typically uses the configured delimiter; we try both.
        string[] parts = line.Contains('\t')
            ? line.Split('\t')
            : line.Contains(',')
                ? line.Split(',')
                : [line];

        string  content  = parts.Length > 0 ? parts[0].Trim() : line.Trim();
        string? symbRaw  = parts.Length > 2 ? parts[2].Trim()   // field 3 = symbology
                         : parts.Length > 1 ? null               // field 2 = decode-time (unused)
                         : null;
        string  symbology = NormaliseSymbologyName(symbRaw);

        return new VerificationRecord
        {
            VerificationDateTime = DateTime.Now,
            DecodedData          = content,
            Symbology            = symbology,
            SymbologyFamily      = new VerificationXmlMap().ClassifySymbology(symbology),

            // Session context preserved from device context
            OperatorId      = deviceContext?.OperatorId,
            JobName         = deviceContext?.JobName,
            BatchNumber     = deviceContext?.BatchNumber,
            RollNumber      = deviceContext?.RollNumber,
            CompanyName     = deviceContext?.CompanyName,
            ProductName     = deviceContext?.ProductName,
            CustomNote      = deviceContext?.CustomNote,
            User1           = deviceContext?.User1,
            User2           = deviceContext?.User2,
            DeviceSerial    = deviceContext?.DeviceSerial,
            DeviceName      = deviceContext?.DeviceName,
            FirmwareVersion = deviceContext?.FirmwareVersion,
            CalibrationDate = deviceContext?.CalibrationDate,
        };
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static VerificationRecord Fallback(VerificationRecord? ctx, string reason) =>
        new()
        {
            Symbology       = reason,
            SymbologyFamily = SymbologyFamily.Unknown,
            OperatorId      = ctx?.OperatorId,
            JobName         = ctx?.JobName,
            DeviceSerial    = ctx?.DeviceSerial,
            FirmwareVersion = ctx?.FirmwareVersion,
        };
}
