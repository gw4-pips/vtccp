namespace DeviceInterface.Webscan;

using DeviceInterface.Rfid;
using System.Text.RegularExpressions;
using ExcelEngine.Models;

/// <summary>
/// Literal data extracted from a Webscan TruCheck HTML export.
///
/// This is intentionally separate from DmstHtmlReport. Webscan TruCheck is a
/// USB verifier with its own file-export format; it has no DataMan HTTP result
/// body and must not be correlated through the DataMan scraper.
/// </summary>
public sealed class WebscanHtmlReport
{
    public string SourceFilePath { get; init; } = string.Empty;
    public string SourceFileName => Path.GetFileName(SourceFilePath);
    public string? SourceImagePath { get; init; }
    public WebscanImageProvenance SourceImageProvenance { get; init; }
    public string? SourceImageMimeType { get; init; }
    public string? SourceImageBase64 { get; init; }
    public string RawHtml { get; init; } = string.Empty;
    public bool ParseSucceeded { get; init; }
    public string? ParseError { get; init; }

    // Webscan report header and summary.
    public string? VerifiedDisplay { get; init; }
    public DateTime? VerifiedDateTime { get; init; }
    public string? SoftwareVersion { get; init; }
    public string? DeviceSerial { get; init; }
    public string? Data { get; init; }
    public string? Symbology { get; init; }
    public string? VerifiedBy { get; init; }
    public string? CompanyName { get; init; }
    public string? ProductName { get; init; }
    public string? JobNumber { get; init; }

    // Webscan Verification Grades row — retained verbatim.
    public string? Standard { get; init; }
    public string? OverallGradeDisplay { get; init; }
    public string? ApertureDisplay { get; init; }
    public string? ApertureUnit { get; init; }
    public int? Aperture { get; init; }
    public string? WavelengthDisplay { get; init; }
    public int? Wavelength { get; init; }
    public string? Lighting { get; init; }
    public string? Notes { get; init; }
    public string? FormalGrade { get; init; }

    // General Characteristics — values remain as displayed where the record
    // model has no dedicated literal string field.
    public string? MatrixSize { get; init; }
    public string? HorizontalBWGDisplay { get; init; }
    public decimal? HorizontalBWG { get; init; }
    public string? VerticalBWGDisplay { get; init; }
    public decimal? VerticalBWG { get; init; }
    public string? EncodedCharactersDisplay { get; init; }
    public int? EncodedCharacters { get; init; }
    public string? TotalCodewordsDisplay { get; init; }
    public int? TotalCodewords { get; init; }
    public string? DataCodewordsDisplay { get; init; }
    public int? DataCodewords { get; init; }
    public string? ImagePolarity { get; init; }
    public string? NominalXDim { get; init; }
    public string? ContrastUniformity { get; init; }

    /// <summary>
    /// Every row from the ISO quality table, including rows that do not have a
    /// corresponding VerificationRecord property. No values are recomputed.
    /// </summary>
    public IReadOnlyList<WebscanQualityParameter> QualityParameters { get; init; } = [];

    /// <summary>
    /// Native Data Format Check evidence copied from the Webscan export. The
    /// outcome is only populated when the export states PASS or FAIL; it is
    /// never recomputed from decoded data or row values.
    /// </summary>
    public DataFormatCheckResult? DataFormatCheck { get; init; }

    public VerificationRecord ToVerificationRecord()
    {
        if (!ParseSucceeded)
            throw new InvalidOperationException(ParseError ?? "Webscan report did not parse.");
        if (string.IsNullOrWhiteSpace(Symbology))
            throw new InvalidDataException("Webscan report is missing its Symbology value.");
        if (VerifiedDateTime is null)
            throw new InvalidDataException("Webscan report is missing its verified timestamp.");

        var overall = WebscanHtmlParser.ParseGrade(OverallGradeDisplay);
        var quality = QualityParameters.ToDictionary(
            p => p.Name,
            StringComparer.OrdinalIgnoreCase);

        WebscanQualityParameter? uec = Find(quality, "Unused Error Correction (UEC)");
        WebscanQualityParameter? sc = Find(quality, "Symbol Contrast (SC)");
        WebscanQualityParameter? mod = Find(quality, "Modulation (MOD)");
        WebscanQualityParameter? rm = Find(quality, "Reflectance Margin (RM)");
        WebscanQualityParameter? anu = Find(quality, "Axial Nonuniformity (ANU)");
        WebscanQualityParameter? gnu = Find(quality, "Grid Nonuniformity (GNU)");
        WebscanQualityParameter? fpd = Find(quality, "Fixed Pattern Damage (FPD)");
        WebscanQualityParameter? ag = Find(quality, "Average Grade (AG)");
        WebscanQualityParameter? decode = Find(quality, "DECODE");

        return new VerificationRecord
        {
            VerificationDateTime = VerifiedDateTime.Value,
            Symbology = Symbology,
            SymbologyFamily = WebscanHtmlParser.MapSymbologyFamily(Symbology),
            DecodedData = Data,
            OperatorId = VerifiedBy,
            JobName = JobNumber,
            CompanyName = CompanyName,
            ProductName = ProductName,
            DeviceSerial = DeviceSerial,
            DeviceName = "Webscan TruCheck",
            DeviceModel = "TC-829",
            VerifierBrand = "WEBSCAN",
            FirmwareVersion = null,
            SoftwareVersion = SoftwareVersion,
            ConnectionMedium = "USB",
            FormalGrade = FormalGrade,
            OverallGrade = overall,
            SymbolAnsiGrade = overall,
            CustomPassFail = OverallPassFail.NotApplicable,
            Aperture = Aperture,
            Wavelength = Wavelength,
            Lighting = Lighting,
            HtmlNotes = Notes,
            Standard = Standard,
            MatrixSize = MatrixSize,
            HorizontalBWG = HorizontalBWG,
            VerticalBWG = VerticalBWG,
            EncodedCharacters = EncodedCharacters,
            TotalCodewords = TotalCodewords,
            DataCodewords = DataCodewords,
            ImagePolarity = WebscanHtmlParser.MapImagePolarity(ImagePolarity),
            NominalXDim_2D = ParseNominalXDim(NominalXDim),
            ContrastUniformity = ContrastUniformity,
            DataFormatCheck = DataFormatCheck,
            UEC_Percent = uec?.MeasuredPercent,
            UEC_Grade = uec?.Grade,
            SC_Percent = sc?.MeasuredPercent,
            SC_RlRd = sc?.SecondaryValue,
            SC_Grade = sc?.Grade,
            MOD_Grade = mod?.Grade,
            RM_Grade = rm?.Grade,
            ANU_Percent = anu?.MeasuredPercent,
            ANU_Grade = anu?.Grade,
            GNU_Percent = gnu?.MeasuredPercent,
            GNU_Grade = gnu?.Grade,
            FPD_Grade = fpd?.Grade,
            AG_Value = ag?.MeasuredNumeric,
            AG_Grade = ag?.Grade,
            DECODE_Grade = decode?.Grade,

            // These existing metadata slots are used by the report pipeline
            // only as source-artifact metadata. They do not imply DataMan HTTP.
            WebscanSourcePath = SourceFilePath,
            HtmlSourceFileName = SourceFileName,
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlVerifiedString = VerifiedDisplay,
            HtmlSymbology = Symbology,
            HtmlDecodedData = Data,
            HtmlStandard = Standard,
            HtmlOverallGradeDisplay = OverallGradeDisplay,
            HtmlAperture = ApertureDisplay,
            HtmlApertureUnit = ApertureUnit,
            HtmlWavelength = WavelengthDisplay,
            HtmlLighting = Lighting,
            HtmlFormalGrade = FormalGrade,
            HtmlBarcodeImageBase64 = SourceImageBase64,
            HtmlBarcodeImageProvenance = SourceImageProvenance.ToString(),
            HtmlBarcodeImageMimeType = SourceImageMimeType,
            HtmlDataFormatCheck = DataFormatCheck,
            HtmlQualityParameters = QualityParameters.Select(p => new NativeQualityParameter
            {
                Number = p.Number, Name = p.Name, MeasuredValue = p.MeasuredValue,
                GradeDisplay = p.GradeDisplay, SecondaryValue = p.SecondaryValue, Result = p.Result,
            }).ToArray(),
            DataSourceExceptions =
                "Webscan:USB HTML export; values are literal report fields" +
                (DataFormatCheck is not null
                    ? "; DataFormatCheck:Webscan HTML native table"
                    : string.Empty),
        };
    }

    internal VerificationRecord ToCompositePrimaryRecord(WebscanHtmlReport linear)
    {
        VerificationRecord primary = ToVerificationRecord();
        string? linearGtin = NormalizeLinearGtin(linear.Symbology, linear.Data);
        string? twoDGtin = ExtractGtin14(Data);
        bool? match = linearGtin is not null && twoDGtin is not null
            ? linearGtin == twoDGtin
            : null;

        var linearGrade = WebscanHtmlParser.ParseGrade(linear.OverallGradeDisplay);
        return primary with
        {
            IsWebscanComposite = true,
            LinearSymbology = linear.Symbology,
            LinearDecodedData = linear.Data,
            LinearOverallGrade = linearGrade,
            LinearFormalGrade = linear.FormalGrade,
            LinearAperture = linear.Aperture,
            LinearWavelength = linear.Wavelength,
            LinearLighting = linear.Lighting,
            LinearStandard = linear.Standard,
            LinearJpegImageBase64 = linear.SourceImageBase64,
            LinearDataFormatCheck = linear.DataFormatCheck,
            LinearQualityParameters = linear.QualityParameters.Select(p => new NativeQualityParameter
            {
                Number = p.Number, Name = p.Name, MeasuredValue = p.MeasuredValue,
                GradeDisplay = p.GradeDisplay, SecondaryValue = p.SecondaryValue, Result = p.Result,
            }).ToArray(),
            HtmlLinearStandard = linear.Standard,
            HtmlLinearGradeDisplay = linear.OverallGradeDisplay,
            HtmlLinearAperture = linear.ApertureDisplay,
            HtmlLinearWavelength = linear.WavelengthDisplay,
            HtmlLinearLighting = linear.Lighting,
            HtmlLinearFormalGrade = linear.FormalGrade,
            LinearTwoDMatch = match,
            LinearTwoDComparisonDetail = match is true
                ? $"GTIN-14 {linearGtin} matches 2D"
                : match is false
                    ? $"GTIN-14 mismatch: linear {linearGtin ?? "unavailable"}, 2D {twoDGtin ?? "unavailable"}"
                    : "GTIN-14 comparison unavailable",
            DataSourceExceptions = primary.DataSourceExceptions + "; Webscan dual-symbology: linear + 2D native reports",
        };
    }

    private static string? ExtractGtin14(string? data)
    {
        if (string.IsNullOrWhiteSpace(data)) return null;
        Match ai = Regex.Match(data, @"(?:^|[/\x1d|])01/?(\d{14})(?:[/\x1d|]|$)");
        if (!ai.Success)
            ai = Regex.Match(data, @"(?:^|[^0-9])01(\d{14})(?:[^0-9]|$)");
        return ai.Success ? ai.Groups[1].Value : null;
    }

    private static string? NormalizeLinearGtin(string? symbology, string? data)
    {
        if (string.IsNullOrWhiteSpace(data)) return null;
        string digits = new(data.Where(char.IsDigit).ToArray());
        if (symbology?.Contains("UPC", StringComparison.OrdinalIgnoreCase) == true &&
            digits.Length == 12) return "00" + digits;
        if (symbology?.Contains("EAN-13", StringComparison.OrdinalIgnoreCase) == true &&
            digits.Length == 13) return "0" + digits;
        if (symbology?.Contains("EAN-8", StringComparison.OrdinalIgnoreCase) == true &&
            digits.Length == 8) return "000000" + digits;
        return digits.Length == 14 ? digits : null;
    }

    private static WebscanQualityParameter? Find(
        IReadOnlyDictionary<string, WebscanQualityParameter> quality,
        string name)
        => quality.TryGetValue(name, out var result) ? result : null;

    private static decimal? ParseNominalXDim(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string numeric = new string(value
            .TakeWhile(c => char.IsDigit(c) || c is '.' or ',')
            .ToArray());
        return decimal.TryParse(
            numeric.Replace(',', '.'),
            System.Globalization.NumberStyles.Number,
            System.Globalization.CultureInfo.InvariantCulture,
            out var result)
            ? result
            : null;
    }
}

/// <summary>
/// One Webscan HTML export containing the native reports belonging to one
/// multi-symbol verification event. A two-symbol linear-plus-2D instance is
/// specifically a dual-symbology report; it is not a GS1 Composite Component.
/// </summary>
public class WebscanHtmlCompositeReport
{
    public string SourceFilePath { get; init; } = string.Empty;
    public string RawHtml { get; init; } = string.Empty;
    public bool ParseSucceeded { get; init; }
    public string? ParseError { get; init; }
    /// <summary>Every structurally valid native report, in source order.</summary>
    public IReadOnlyList<WebscanHtmlReport> SymbolReports { get; init; } = [];
    public WebscanHtmlReport? LinearReport { get; init; }
    public WebscanHtmlReport? TwoDReport { get; init; }

    public MultiSymbolQualification Qualify(string? rfidGtin14)
    {
        var evidence = SymbolReports.Select((report, index) =>
            new MultiSymbolIdentityEvidence
            {
                Ordinal = index + 1,
                Symbology = report.Symbology,
                Family = WebscanHtmlParser.MapSymbologyFamily(report.Symbology ?? string.Empty).ToString(),
                Gtin14 = ExtractIdentityGtin(report),
            }).ToArray();
        var reasons = new List<string>();
        var unsupported = evidence.Where(e => e.Family == SymbologyFamily.Unknown.ToString()).ToArray();
        var missing = evidence.Where(e => e.Gtin14 is null).ToArray();
        var gtins = evidence.Where(e => e.Gtin14 is not null).Select(e => e.Gtin14!).Distinct().ToArray();
        bool duplicateFamily = evidence.GroupBy(e => e.Family)
            .Any(g => g.Key != SymbologyFamily.Linear1D.ToString() &&
                      g.Key != SymbologyFamily.Unknown.ToString() && g.Count() > 1);

        if (unsupported.Length > 0)
            reasons.Add($"unsupported symbols: {string.Join(", ", unsupported.Select(e => $"#{e.Ordinal} {e.Symbology}"))}");
        if (missing.Length > 0)
            reasons.Add($"missing GTIN-14 (AI (01) or linear identity): {string.Join(", ", missing.Select(e => $"symbol #{e.Ordinal}"))}");
        if (gtins.Length > 1)
            reasons.Add($"GTIN mismatch: {string.Join(", ", gtins)}");
        if (duplicateFamily)
            reasons.Add("same-family conflict: more than one recognized 2D symbol family member");
        if (rfidGtin14 is null)
            reasons.Add("RFID GTIN unavailable");
        else if (gtins.Length == 1 && rfidGtin14 != gtins[0])
            reasons.Add($"RFID GTIN mismatch: RFID {rfidGtin14}, symbols {gtins[0]}");

        MultiSymbolQualificationStatus status =
            unsupported.Length > 0 || gtins.Length > 1 ||
            (rfidGtin14 is not null && gtins.Length == 1 && rfidGtin14 != gtins[0])
                ? MultiSymbolQualificationStatus.Rejected
                : missing.Length > 0 || duplicateFamily || rfidGtin14 is null
                    ? MultiSymbolQualificationStatus.Unverified
                    : MultiSymbolQualificationStatus.Qualified;
        if (reasons.Count == 0) reasons.Add("all recognized symbol identities agree with RFID EPC");
        return new MultiSymbolQualification
        {
            Status = status,
            Reasons = reasons,
            Symbols = evidence,
            MatchingSymbols = rfidGtin14 is null ? [] : evidence.Where(e => e.Gtin14 == rfidGtin14).Select(e => e.Ordinal).ToArray(),
            MismatchingSymbols = rfidGtin14 is null ? [] : evidence.Where(e => e.Gtin14 is not null && e.Gtin14 != rfidGtin14).Select(e => e.Ordinal).ToArray(),
        };
    }

    private static string? ExtractIdentityGtin(WebscanHtmlReport report)
    {
        if (WebscanHtmlParser.MapSymbologyFamily(report.Symbology ?? string.Empty) == SymbologyFamily.Linear1D)
            return RfidValidator.NormalizeLinearGtin14(report.Symbology, report.Data);
        return RfidValidator.ExtractAi01(report.Data);
    }

    public VerificationRecord ToVerificationRecord()
    {
        if (!ParseSucceeded || SymbolReports.Count < 2)
            throw new InvalidOperationException(
                ParseError ?? "Webscan multi-symbol report did not parse.");

        string? linearIdentity = LinearReport is null
            ? null
            : ExtractIdentityGtin(LinearReport);
        // For multi-2D exports, prefer the unique 2D report whose identity agrees
        // with the linear symbol. This is an identity decision, not a source-order
        // decision; source order remains preserved in MultiSymbolReports.
        WebscanHtmlReport primaryReport =
            (linearIdentity is not null
                ? SymbolReports.Where(report =>
                    WebscanHtmlParser.MapSymbologyFamily(report.Symbology ?? string.Empty) != SymbologyFamily.Linear1D &&
                    ExtractIdentityGtin(report) == linearIdentity).FirstOrDefault()
                : null)
            ?? TwoDReport
            ?? SymbolReports[0];
        VerificationRecord primary = primaryReport.ToVerificationRecord();
        VerificationRecord? linear = LinearReport?.ToVerificationRecord();
        bool isDualSymbology = LinearReport is not null && TwoDReport is not null &&
                               SymbolReports.Count == 2;

        return primary with
        {
            IsWebscanComposite = isDualSymbology,
            MultiSymbolReports = SymbolReports.Select((report, index) =>
                ToNativeSummary(report, index + 1)).ToArray(),
            MultiSymbolQualificationStatus = Qualify(null).Status.ToString(),
            MultiSymbolQualificationReasons = Qualify(null).Reasons,
            LinearSymbology = linear?.Symbology,
            LinearDecodedData = linear?.DecodedData,
            LinearOverallGrade = linear?.OverallGrade,
            LinearFormalGrade = linear?.FormalGrade,
            LinearAperture = linear?.Aperture,
            LinearWavelength = linear?.Wavelength,
            LinearLighting = linear?.Lighting,
            LinearStandard = linear?.Standard,
            LinearJpegImageBase64 = linear?.HtmlBarcodeImageBase64,
            LinearDataFormatCheck = linear?.HtmlDataFormatCheck,
            LinearQualityParameters = linear?.HtmlQualityParameters ?? [],
            HtmlQualityParameters = primary.HtmlQualityParameters,
            LinearGtin14 = linear is null ? null : RfidValidator.NormalizeLinearGtin14(
                linear.Symbology, linear.DecodedData),
            LinearTwoDMatch = linear is null || TwoDReport is null ? null :
                RfidValidator.NormalizeLinearGtin14(linear.Symbology, linear.DecodedData) is { } linearGtin &&
                RfidValidator.ExtractAi01(primary.DecodedData) is { } twoDGtin &&
                linearGtin == twoDGtin,
            BarcodeSymbolAgreementDetail =
                $"2D GTIN-14: {RfidValidator.ExtractAi01(primary.DecodedData) ?? "missing"}; " +
                $"linear GTIN-14: {(linear is null ? "missing" : RfidValidator.NormalizeLinearGtin14(linear.Symbology, linear.DecodedData) ?? "missing")}",
            HtmlLinearStandard = linear?.HtmlStandard,
            HtmlLinearGradeDisplay = linear?.HtmlOverallGradeDisplay,
            HtmlLinearAperture = linear?.HtmlAperture,
            HtmlLinearWavelength = linear?.HtmlWavelength,
            // The linear Webscan summary's sixth cell is Notes; the legacy
            // multi-mode report slot is named HtmlLinearLighting.
            HtmlLinearLighting = linear?.HtmlNotes ?? linear?.HtmlLighting,
            HtmlLinearFormalGrade = linear?.HtmlFormalGrade,
            IsStandaloneLinear = false,
            DataSourceExceptions =
                $"{primary.DataSourceExceptions}; Webscan multi-symbol: {SymbolReports.Count} independent native reports",
        };
    }

    private static NativeWebscanReportSummary ToNativeSummary(
        WebscanHtmlReport report,
        int ordinal)
        => new()
        {
            Ordinal = ordinal,
            Symbology = report.Symbology,
            SymbologyFamily = WebscanHtmlParser.MapSymbologyFamily(report.Symbology ?? string.Empty).ToString(),
            DecodedData = report.Data,
            Gtin14 = ExtractIdentityGtin(report),
            SourceImagePath = report.SourceImagePath,
            SourceImageProvenance = report.SourceImageProvenance.ToString(),
            SourceImageBase64 = report.SourceImageBase64,
            SourceImageMimeType = report.SourceImageMimeType,
            Standard = report.Standard,
            OverallGradeDisplay = report.OverallGradeDisplay,
            ApertureDisplay = report.ApertureDisplay,
            ApertureUnit = report.ApertureUnit,
            WavelengthDisplay = report.WavelengthDisplay,
            Lighting = report.Lighting,
            Notes = report.Notes,
            FormalGrade = report.FormalGrade,
            QualityParameters = report.QualityParameters.Select(p => new NativeQualityParameter
            {
                Number = p.Number, Name = p.Name, MeasuredValue = p.MeasuredValue,
                GradeDisplay = p.GradeDisplay, SecondaryValue = p.SecondaryValue, Result = p.Result,
            }).ToArray(),
            DataFormatCheck = report.DataFormatCheck,
        };

    internal static WebscanHtmlCompositeReport Failure(
        string rawHtml,
        string sourcePath,
        string error)
        => new()
        {
            SourceFilePath = sourcePath,
            RawHtml = rawHtml,
            ParseSucceeded = false,
            ParseError = error,
        };
}

/// <summary>
/// Canonical name for a Webscan export containing independent native reports.
/// The older WebscanHtmlCompositeReport type remains as a compatibility
/// boundary for callers that persisted that type name.
/// </summary>
public sealed class WebscanHtmlMultiSymbolReport
{
    private readonly WebscanHtmlCompositeReport _inner;
    private WebscanHtmlMultiSymbolReport(WebscanHtmlCompositeReport inner) => _inner = inner;

    internal static WebscanHtmlMultiSymbolReport From(WebscanHtmlCompositeReport report)
        => new(report);

    public string SourceFilePath => _inner.SourceFilePath;
    public string RawHtml => _inner.RawHtml;
    public bool ParseSucceeded => _inner.ParseSucceeded;
    public string? ParseError => _inner.ParseError;
    public IReadOnlyList<WebscanHtmlReport> SymbolReports => _inner.SymbolReports;
    public WebscanHtmlReport? LinearReport => _inner.LinearReport;
    public WebscanHtmlReport? TwoDReport => _inner.TwoDReport;

    public MultiSymbolQualification Qualify(string? rfidGtin14)
        => _inner.Qualify(rfidGtin14);

    public VerificationRecord ToVerificationRecord()
        => _inner.ToVerificationRecord();
}

public enum WebscanImageProvenance
{
    None,
    EmbeddedHtml,
    SiblingExport,
}

public sealed class WebscanQualityParameter
{
    public string Number { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string? MeasuredValue { get; init; }
    public decimal? MeasuredNumeric { get; init; }
    public decimal? MeasuredPercent { get; init; }
    public string? GradeDisplay { get; init; }
    public GradingResult? Grade { get; init; }
    public string? SecondaryValue { get; init; }
    public string? Result { get; init; }
}