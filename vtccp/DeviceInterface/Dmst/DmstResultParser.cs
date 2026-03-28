namespace DeviceInterface.Dmst;

using System.Globalization;
using System.Xml.Linq;
using ExcelEngine.Models;

/// <summary>
/// Parses a Cognex DataMan ReadXml result (firmware 6.x &lt;result&gt; format) or a legacy
/// DMST push string (&lt;DMCCResponse&gt;/&lt;DMSymVerResponse&gt;) into a
/// <see cref="VerificationRecord"/>.
///
/// Firmware 6.x XML structure:
///   &lt;result&gt;
///     &lt;general&gt;&lt;symbology&gt;Data Matrix&lt;/symbology&gt; ...&lt;/general&gt;
///     &lt;trucheck_verificaiton_result&gt;               (sic — typo in firmware)
///       &lt;CalibrationDate&gt;...&lt;/CalibrationDate&gt;
///       &lt;SymbolData&gt;
///         &lt;SymbologyType&gt;DataMatrix&lt;/SymbologyType&gt;
///         &lt;DecodedData&gt;...&lt;/DecodedData&gt;
///         &lt;ReportSection sectionType="GradingInfo"&gt;
///           &lt;GradeInfo&gt;&lt;Standard&gt;ISO 15415:2011&lt;/Standard&gt;
///             &lt;Grade&gt;3.0&lt;/Grade&gt;&lt;ValueGrade&gt;B&lt;/ValueGrade&gt;
///             &lt;FormalGrade&gt;3.0/10/660/45Q&lt;/FormalGrade&gt;
///             &lt;Aperture&gt;10&lt;/Aperture&gt;&lt;Wavelength&gt;660&lt;/Wavelength&gt;
///           &lt;/GradeInfo&gt;
///         &lt;/ReportSection&gt;
///         &lt;ReportSection sectionType="GradeHistory"&gt;
///           &lt;VerificationOverallPass&gt;1&lt;/VerificationOverallPass&gt;
///         &lt;/ReportSection&gt;
///         &lt;ReportSection sectionTitle="ISO15415 Quality Parameters"&gt;
///           &lt;Parameter&gt;&lt;Number&gt;1&lt;/Number&gt;&lt;Grade&gt;4.0&lt;/Grade&gt;&lt;Value&gt;100.0%&lt;/Value&gt;&lt;/Parameter&gt;
///           ...
///         &lt;/ReportSection&gt;
///         &lt;ReportSection sectionTitle="General Characteristics"&gt;
///           &lt;Parameter&gt;&lt;Name&gt;Matrix Size&lt;/Name&gt;&lt;Data&gt;22x22 (Data: 20x20)&lt;/Data&gt;&lt;/Parameter&gt;
///           ...
///         &lt;/ReportSection&gt;
///       &lt;/SymbolData&gt;
///     &lt;/trucheck_verificaiton_result&gt;
///     &lt;general&gt;&lt;full_string encoding="base64"&gt;...&lt;/full_string&gt;&lt;/general&gt;
///   &lt;/result&gt;
///
/// Tolerant design: missing/malformed elements are silently skipped.
/// </summary>
public static class DmstResultParser
{
    // ── Public entry points ───────────────────────────────────────────────────

    public static VerificationRecord Parse(
        string               xml,
        VerificationXmlMap?  map           = null,
        VerificationRecord?  deviceContext = null)
    {
        map ??= new VerificationXmlMap();

        XDocument doc;
        try   { doc = XDocument.Parse(xml); }
        catch { return Fallback(deviceContext, "XML parse failed"); }

        // Navigate to the result container element.
        // Firmware 6.x: <result> root (no DMSymVerResponse wrapper).
        // Legacy DMST: <DMCCResponse>/<DMSymVerResponse>
        XElement? container =
            doc.Descendants(map.ResultContainer).FirstOrDefault()
            ?? doc.Root;

        if (container is null)
            return Fallback(deviceContext, "ResultContainer not found");

        // ── Helper closures ───────────────────────────────────────────────────

        string? Str(string elem) =>
            container.Descendants(elem).FirstOrDefault()?.Value.Trim()
                is { Length: > 0 } v ? v : null;

        decimal? Dec(string elem) =>
            decimal.TryParse(Str(elem), NumberStyles.Any,
                CultureInfo.InvariantCulture, out var d) ? d : null;

        int? Int(string elem) =>
            int.TryParse(Str(elem), out var i) ? i : null;

        // Legacy grade helper: looks for a named element whose text is the
        // letter grade and optionally a "numeric" attribute.
        GradingResult? GradeLegacy(string elem)
        {
            string? letter = Str(elem);
            if (letter is null) return null;
            XElement? el = container.Descendants(elem).FirstOrDefault();
            decimal   num = 0m;
            if (el?.Attribute("numeric") is { } a)
                decimal.TryParse(a.Value, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out num);
            return GradingResult.FromLetterAndNumeric(letter, num, "");
        }

        // ── Firmware 6.x parameter-section helpers ────────────────────────────

        // ISO 15415 Quality Parameters: <ReportSection sectionTitle="ISO15415 Quality Parameters">
        XElement? isoParamsSection = container
            .Descendants("ReportSection")
            .FirstOrDefault(s =>
                (string?)s.Attribute("sectionTitle") == "ISO15415 Quality Parameters");

        // General Characteristics: <ReportSection sectionTitle="General Characteristics">
        XElement? genCharSection = container
            .Descendants("ReportSection")
            .FirstOrDefault(s =>
                (string?)s.Attribute("sectionTitle") == "General Characteristics");

        // Grade History: <ReportSection sectionType="GradeHistory">
        XElement? gradeHistSection = container
            .Descendants("ReportSection")
            .FirstOrDefault(s =>
                (string?)s.Attribute("sectionType") == "GradeHistory");

        // Grading Info: <ReportSection sectionType="GradingInfo"> / first ISO GradeInfo
        XElement? gradingInfoSection = container
            .Descendants("ReportSection")
            .FirstOrDefault(s =>
                (string?)s.Attribute("sectionType") == "GradingInfo");

        XElement? isoGradeInfo = gradingInfoSection?
            .Elements("GradeInfo")
            .FirstOrDefault(g => g.Element("Standard")?.Value.Contains("ISO") == true)
            ?? gradingInfoSection?.Elements("GradeInfo").FirstOrDefault();

        // Lookup a quality parameter by Number (e.g. "1", "3a") → GradingResult
        GradingResult? ParamGrade(string number)
        {
            XElement? p = isoParamsSection?.Elements("Parameter")
                .FirstOrDefault(e => e.Element("Number")?.Value.Trim() == number);
            if (p is null) return null;
            string? gStr = p.Element("Grade")?.Value.Trim();
            if (string.IsNullOrEmpty(gStr)) return null;
            bool isNum = decimal.TryParse(gStr, NumberStyles.Any,
                CultureInfo.InvariantCulture, out decimal gNum);
            string letter = isNum ? NumericToLetterGrade(gNum) : gStr;
            string check  = p.Element("Check")?.Value.Trim() ?? "";
            return GradingResult.FromLetterAndNumeric(letter, isNum ? gNum : 0m, check);
        }

        // Lookup a quality parameter <Value> (strips trailing %) → decimal
        decimal? ParamValuePct(string number)
        {
            XElement? p = isoParamsSection?.Elements("Parameter")
                .FirstOrDefault(e => e.Element("Number")?.Value.Trim() == number);
            string? raw = p?.Element("Value")?.Value.Trim().TrimEnd('%').Trim();
            if (!decimal.TryParse(raw, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out decimal v)) return null;
            return v;
        }

        // Lookup a quality parameter <Data> string
        string? ParamData(string number)
        {
            XElement? p = isoParamsSection?.Elements("Parameter")
                .FirstOrDefault(e => e.Element("Number")?.Value.Trim() == number);
            return p?.Element("Data")?.Value.Trim() is { Length: > 0 } s ? s : null;
        }

        // Lookup a General Characteristics entry by Name → <Data> string
        string? CharData(string name)
        {
            XElement? p = genCharSection?.Elements("Parameter")
                .FirstOrDefault(e => e.Element("Name")?.Value.Trim() == name);
            return p?.Element("Data")?.Value.Trim() is { Length: > 0 } s ? s : null;
        }

        // Lookup a General Characteristics entry → decimal (strips units)
        decimal? CharDataDec(string name, char stripChar = '%')
        {
            string? raw = CharData(name)?.TrimEnd(stripChar).Trim();
            if (!decimal.TryParse(raw, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out decimal v)) return null;
            return v;
        }

        // Lookup a General Characteristics entry → int
        int? CharDataInt(string name)
        {
            if (!int.TryParse(CharData(name), out int v)) return null;
            return v;
        }

        // ── Symbology ─────────────────────────────────────────────────────────
        // Firmware 6.x: <SymbologyType>DataMatrix</SymbologyType> (in SymbolData)
        //           or: <general>/<symbology>Data Matrix</symbology>
        // Legacy:       <SymbologyName>
        string? symbName =
            Str(map.SymbologyName)                                          // SymbologyType
            ?? container.Descendants("symbology").FirstOrDefault()?.Value   // general/symbology
            ?? deviceContext?.Symbology
            ?? "Unknown";

        SymbologyFamily symbFamily = map.ClassifySymbology(symbName);
        string symbology = NormaliseSymbologyName(symbName);

        // ── Timestamp ─────────────────────────────────────────────────────────
        // Firmware 6.x: timestamp is inside the base64-encoded <full_string>.
        // Decode it to extract <DateTime>; fall back to DateTime.Now.
        DateTime verifyDt = DateTime.Now;
        if (Str(map.DateTime) is { } dtStrDirect)
        {
            DateTime.TryParse(dtStrDirect, out verifyDt);
        }
        else
        {
            // Try the base64-encoded inner XML
            XElement? fullStringEl = container
                .Descendants("full_string").FirstOrDefault();
            if (fullStringEl?.Attribute("encoding")?.Value == "base64"
                && fullStringEl.Value is { Length: > 0 } b64)
            {
                try
                {
                    string inner = System.Text.Encoding
                        .UTF8.GetString(Convert.FromBase64String(b64));
                    XDocument innerDoc = XDocument.Parse(inner);
                    if (innerDoc.Descendants("DateTime")
                            .FirstOrDefault()?.Value is { } dtInner)
                        DateTime.TryParse(dtInner, out verifyDt);
                }
                catch { }
            }
        }

        // ── CalibrationDate from XML ──────────────────────────────────────────
        // <trucheck_verificaiton_result>/<CalibrationDate> (firmware 6.x)
        DateTime? calibDate = deviceContext?.CalibrationDate;
        if (Str("CalibrationDate") is { } calibStr
            && DateTime.TryParse(calibStr,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out DateTime calibParsed))
            calibDate = calibParsed;

        // ── Overall grade ─────────────────────────────────────────────────────
        // Firmware 6.x: letter = <ValueGrade> (B), numeric = isoGradeInfo/<Grade> (3.0)
        string? gradeLetterStr = Str(map.OverallGrade)    // "ValueGrade" → "B"
                              ?? isoGradeInfo?.Element("ValueGrade")?.Value.Trim();

        decimal gradeNumeric = 0m;
        if (isoGradeInfo?.Element("Grade")?.Value is { } isoGradeStr)
            decimal.TryParse(isoGradeStr, NumberStyles.Any,
                CultureInfo.InvariantCulture, out gradeNumeric);
        else if (Dec(map.OverallGradeNumeric) is { } mn)
            gradeNumeric = mn;

        string? formalGrade =
            isoGradeInfo?.Element("FormalGrade")?.Value.Trim()
            ?? Str(map.FormalGrade);

        // Derive pass/fail from letter grade
        string derivedPassFail = gradeLetterStr?.Trim().ToUpper() switch
        {
            "A" or "B" => "PASS",
            "F"        => "FAIL",
            _          => "",
        };

        // Check GradeHistory/VerificationOverallPass to confirm pass/fail
        if (gradeHistSection?.Element("VerificationOverallPass")?.Value is { } ovp)
            derivedPassFail = ovp.Trim() == "1" ? "PASS" : "FAIL";

        GradingResult? overall = gradeLetterStr is not null
            ? GradingResult.FromLetterAndNumeric(gradeLetterStr, gradeNumeric, derivedPassFail)
            : null;

        // ── Verification settings ─────────────────────────────────────────────
        // Firmware 6.x: inside isoGradeInfo directly; also found via Descendants
        int? aperture  = isoGradeInfo?.Element("Aperture")  is { } ael
                            && int.TryParse(ael.Value, out int ap) ? ap : Int(map.ApertureRef);
        int? wavelength = isoGradeInfo?.Element("Wavelength") is { } wel
                            && int.TryParse(wel.Value, out int wv) ? wv : Int(map.Wavelength);
        string? lighting = isoGradeInfo?.Element("Lighting")?.Value.Trim() ?? Str(map.Lighting);
        string? standard = isoGradeInfo?.Element("Standard")?.Value.Trim() ?? Str(map.Standard);

        // ── 2D quality parameters ─────────────────────────────────────────────
        // Try legacy map names first (DMSymVerResponse), then parameter-number lookup.
        decimal?       uecPct   = Dec(map.UECPercent)  ?? ParamValuePct("1");
        GradingResult? uecGrade = GradeLegacy(map.UECGrade)  ?? ParamGrade("1");

        decimal?       scPct    = Dec(map.SCPercent)   ?? ParamValuePct("2");
        string?        scRlRd   = Str(map.SCRlRd)      ?? ParamData("2");
        GradingResult? scGrade  = GradeLegacy(map.SCGrade)   ?? ParamGrade("2");

        GradingResult? modGrade  = GradeLegacy(map.MODGrade)  ?? ParamGrade("3a");
        GradingResult? rmGrade   = GradeLegacy(map.RMGrade)   ?? ParamGrade("3b");

        decimal?       anuPct   = Dec(map.ANUPercent)  ?? ParamValuePct("4");
        GradingResult? anuGrade = GradeLegacy(map.ANUGrade)  ?? ParamGrade("4");

        decimal?       gnuPct   = Dec(map.GNUPercent)  ?? ParamValuePct("5");
        GradingResult? gnuGrade = GradeLegacy(map.GNUGrade)  ?? ParamGrade("5");

        GradingResult? fpdGrade  = GradeLegacy(map.FPDGrade)  ?? ParamGrade("6");

        GradingResult? llsGrade  = GradeLegacy(map.LLSGrade)  ?? ParamGrade("7");
        GradingResult? blsGrade  = GradeLegacy(map.BLSGrade)  ?? ParamGrade("8");
        GradingResult? lqzGrade  = GradeLegacy(map.LQZGrade)  ?? ParamGrade("9");
        GradingResult? bqzGrade  = GradeLegacy(map.BQZGrade)  ?? ParamGrade("10");
        GradingResult? tqzGrade  = GradeLegacy(map.TQZGrade)  ?? ParamGrade("11");
        GradingResult? rqzGrade  = GradeLegacy(map.RQZGrade)  ?? ParamGrade("12");

        decimal?       ttrPct    = Dec(map.TTRPercent)  ?? ParamValuePct("13");
        GradingResult? ttrGrade  = GradeLegacy(map.TTRGrade)   ?? ParamGrade("13");
        decimal?       rtrPct    = Dec(map.RTRPercent)  ?? ParamValuePct("14");
        GradingResult? rtrGrade  = GradeLegacy(map.RTRGrade)   ?? ParamGrade("14");

        GradingResult? tctGrade  = GradeLegacy(map.TCTGrade)   ?? ParamGrade("15");
        GradingResult? rctGrade  = GradeLegacy(map.RCTGrade)   ?? ParamGrade("16");

        decimal?       agVal    = Dec(map.AGValue)     ?? ParamValuePct("17");
        GradingResult? agGrade  = GradeLegacy(map.AGGrade)    ?? ParamGrade("17");

        GradingResult? decGrade  = GradeLegacy(map.DecodeGrade) ?? ParamGrade("18");

        // ── 2D quadrant parameters (≥32×32) — legacy only, not in firmware 6.x ──
        GradingResult? ulqzGrade  = GradeLegacy(map.ULQZGrade);
        GradingResult? urqzGrade  = GradeLegacy(map.URQZGrade);
        GradingResult? ruqzGrade  = GradeLegacy(map.RUQZGrade);
        GradingResult? rlqzGrade  = GradeLegacy(map.RLQZGrade);

        decimal? ulqttrPct = Dec(map.ULQTTRPercent); GradingResult? ulqttrGrade = GradeLegacy(map.ULQTTRGrade);
        decimal? urqttrPct = Dec(map.URQTTRPercent); GradingResult? urqttrGrade = GradeLegacy(map.URQTTRGrade);
        decimal? llqttrPct = Dec(map.LLQTTRPercent); GradingResult? llqttrGrade = GradeLegacy(map.LLQTTRGrade);
        decimal? lrqttrPct = Dec(map.LRQTTRPercent); GradingResult? lrqttrGrade = GradeLegacy(map.LRQTTRGrade);

        decimal? ulqrtrPct = Dec(map.ULQRTRPercent); GradingResult? ulqrtrGrade = GradeLegacy(map.ULQRTRGrade);
        decimal? urqrtrPct = Dec(map.URQRTRPercent); GradingResult? urqrtrGrade = GradeLegacy(map.URQRTRGrade);
        decimal? llqrtrPct = Dec(map.LLQRTRPercent); GradingResult? llqrtrGrade = GradeLegacy(map.LLQRTRGrade);
        decimal? lrqrtrPct = Dec(map.LRQRTRPercent); GradingResult? lrqrtrGrade = GradeLegacy(map.LRQRTRGrade);

        GradingResult? ulqtctGrade = GradeLegacy(map.ULQTCTGrade); GradingResult? urqtctGrade = GradeLegacy(map.URQTCTGrade);
        GradingResult? llqtctGrade = GradeLegacy(map.LLQTCTGrade); GradingResult? lrqtctGrade = GradeLegacy(map.LRQTCTGrade);
        GradingResult? ulqrctGrade = GradeLegacy(map.ULQRCTGrade); GradingResult? urqrctGrade = GradeLegacy(map.URQRCTGrade);
        GradingResult? llqrctGrade = GradeLegacy(map.LLQRCTGrade); GradingResult? lrqrctGrade = GradeLegacy(map.LRQRCTGrade);

        // ── 2D matrix characteristics ─────────────────────────────────────────
        // Firmware 6.x: all in General Characteristics ReportSection <Data> elements.
        string? matrixSize = Str(map.MatrixSize) ?? CharData("Matrix Size");

        decimal? hBwg = Dec(map.HorizontalBWG);
        if (!hBwg.HasValue)
            hBwg = CharDataDec("Horizontal BWG");

        decimal? vBwg = Dec(map.VerticalBWG);
        if (!vBwg.HasValue)
            vBwg = CharDataDec("Vertical BWG");

        int? encodedChars = Int(map.EncodedCharacters) ?? CharDataInt("Encoded characters");
        int? totalCw      = Int(map.TotalCodewords)    ?? CharDataInt("Total Codewords");
        int? dataCw       = Int(map.DataCodewords)     ?? CharDataInt("Data Codewords");
        int? ecBudget     = Int(map.ErrorCorrectionBudget) ?? CharDataInt("Error Correction Budget");
        int? ecCorrected  = Int(map.ErrorsCorrected)   ?? CharDataInt("Errors Corrected");
        int? ecCapUsed    = Int(map.ErrorCapacityUsed) ?? CharDataInt("Error Capacity Used");
        string? ecType    = Str(map.ErrorCorrectionType) ?? CharData("Error Correction Type");

        // Nominal X Dim — strip units ("13.1 mil" → 13.1)
        decimal? nomXDim2D = Dec(map.NominalXDim);
        if (!nomXDim2D.HasValue)
        {
            string? ndRaw = CharData("Nominal X Dim")?.Split(' ').FirstOrDefault();
            if (decimal.TryParse(ndRaw, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out decimal nd))
                nomXDim2D = nd;
        }

        decimal? ppm = Dec(map.PixelsPerModule);
        if (!ppm.HasValue)
        {
            string? ppmRaw = CharData("Pixels per Module");
            if (decimal.TryParse(ppmRaw, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out decimal ppmv))
                ppm = ppmv;
        }

        string? contrastUniformity = Str(map.ContrastUniformity)
                                  ?? CharData("Contrast Uniformity");
        string? mrd   = Str(map.MRD) ?? CharData("MRD");

        // Image polarity: <Data>Black on white</Data>
        ImagePolarity polarity = ImagePolarity.Unknown;
        string? imagePol = Str(map.ImagePolarity) ?? CharData("Image");
        if (imagePol is not null)
            Enum.TryParse(imagePol.Replace(" ", ""), true, out polarity);

        // ── 1D ISO 15416 summary parameters (legacy / not present in firmware 6.x) ──
        GradingResult? symbolAnsiGrade = GradeLegacy(map.SymbolAnsiGrade);
        decimal? avgEdge   = Dec(map.AvgEdge);
        string?  avgRlRd   = Str(map.AvgRlRd);
        decimal? avgSC     = Dec(map.AvgSC);
        decimal? avgMinEC  = Dec(map.AvgMinEC);
        decimal? avgMOD    = Dec(map.AvgMOD);
        decimal? avgDefect = Dec(map.AvgDefect);
        string?  avgDcod   = Str(map.AvgDcod);
        decimal? avgDEC    = Dec(map.AvgDEC);
        decimal? avgLQZ    = Dec(map.AvgLQZ);
        decimal? avgRQZ    = Dec(map.AvgRQZ);
        decimal? avgHQZ    = Dec(map.AvgHQZ);
        decimal? avgMinQZ  = Dec(map.AvgMinQZ);
        decimal? bwgPct    = Dec(map.BWGPercent);
        decimal? magnif    = Dec(map.Magnification);
        decimal? ratio     = Dec(map.Ratio);
        decimal? nomXDim1D = Dec(map.NominalXDim1D);

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
                        NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;

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

            // Session context (caller-supplied via deviceContext; XML values override where available)
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
            CalibrationDate = calibDate,          // now extracted from XML when present

            // 2D quality
            UEC_Percent = uecPct,
            UEC_Grade   = uecGrade,
            SC_Percent  = scPct,
            SC_RlRd     = scRlRd,
            SC_Grade    = scGrade,
            MOD_Grade   = modGrade,
            RM_Grade    = rmGrade,
            ANU_Percent = anuPct,
            ANU_Grade   = anuGrade,
            GNU_Percent = gnuPct,
            GNU_Grade   = gnuGrade,
            FPD_Grade   = fpdGrade,
            DECODE_Grade = decGrade,
            AG_Value    = agVal,
            AG_Grade    = agGrade,

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

    // ── Plain-text entry point ────────────────────────────────────────────────

    /// <summary>
    /// Creates a minimal <see cref="VerificationRecord"/> from a single plain-text
    /// push line received via the DataMan Network Client (Format Data → Basic/Standard).
    /// </summary>
    public static VerificationRecord ParseText(
        string              line,
        VerificationRecord? deviceContext = null)
    {
        string[] parts = line.Contains('\t')
            ? line.Split('\t')
            : line.Contains(',')
                ? line.Split(',')
                : [line];

        string  content  = parts.Length > 0 ? parts[0].Trim() : line.Trim();
        string? symbRaw  = parts.Length > 2 ? parts[2].Trim() : null;
        string  symbology = NormaliseSymbologyName(symbRaw);

        return new VerificationRecord
        {
            VerificationDateTime = DateTime.Now,
            DecodedData          = content,
            Symbology            = symbology,
            SymbologyFamily      = new VerificationXmlMap().ClassifySymbology(symbology),

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

    /// <summary>
    /// Converts an ISO 15415 numeric grade (0–4) to its letter equivalent.
    /// 4 → A, 3 → B, 2 → C, 1 → D, 0 → F.
    /// </summary>
    private static string NumericToLetterGrade(decimal num) => num switch
    {
        >= 3.5m => "A",
        >= 2.5m => "B",
        >= 1.5m => "C",
        >= 0.5m => "D",
        _       => "F",
    };

    /// <summary>
    /// Normalises DataMan device symbology name strings to VTCCP canonical forms.
    /// </summary>
    private static string NormaliseSymbologyName(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "Unknown";
        return raw.Trim() switch
        {
            "UPC-A"                  => "UPCA",
            "UPC-E"                  => "UPCE",
            "EAN-8"                  => "EAN8",
            "EAN-13"                 => "EAN13",
            "Code 128"               => "Code128",
            "Code 39"                => "Code39",
            "I 2/5"                  => "ITF",
            "QR Code"                => "QRCode",
            "GS1 QR Code"            => "GS1 QRCode",
            "Data Matrix"            => "DataMatrix",
            "Data Matrix ECC 200"    => "DataMatrix",
            _                        => raw.Trim(),
        };
    }

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
