namespace DeviceInterface.Webscan;

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
            HtmlWavelength = WavelengthDisplay,
            HtmlLighting = Lighting,
            HtmlFormalGrade = FormalGrade,
            HtmlBarcodeImageBase64 = SourceImageBase64,
            HtmlBarcodeImageProvenance = SourceImageProvenance.ToString(),
            HtmlBarcodeImageMimeType = SourceImageMimeType,
            HtmlDataFormatCheck = DataFormatCheck,
            DataSourceExceptions =
                "Webscan:USB HTML export; values are literal report fields" +
                (DataFormatCheck is not null
                    ? "; DataFormatCheck:Webscan HTML native table"
                    : string.Empty),
        };
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
/// One Webscan HTML export containing the two reports that belong to one
/// multi-symbol verification event.
/// </summary>
public sealed class WebscanHtmlCompositeReport
{
    public string SourceFilePath { get; init; } = string.Empty;
    public string RawHtml { get; init; } = string.Empty;
    public bool ParseSucceeded { get; init; }
    public string? ParseError { get; init; }
    public WebscanHtmlReport? LinearReport { get; init; }
    public WebscanHtmlReport? TwoDReport { get; init; }

    public VerificationRecord ToVerificationRecord()
    {
        if (!ParseSucceeded || LinearReport is null || TwoDReport is null)
            throw new InvalidOperationException(
                ParseError ?? "Webscan composite report did not parse.");

        VerificationRecord twoD = TwoDReport.ToVerificationRecord();
        VerificationRecord linear = LinearReport.ToVerificationRecord();

        return twoD with
        {
            LinearSymbology = linear.Symbology,
            LinearDecodedData = linear.DecodedData,
            LinearOverallGrade = linear.OverallGrade,
            LinearFormalGrade = linear.FormalGrade,
            LinearAperture = linear.Aperture,
            LinearWavelength = linear.Wavelength,
            LinearLighting = linear.Lighting,
            LinearStandard = linear.Standard,
            LinearJpegImageBase64 = linear.HtmlBarcodeImageBase64,
            LinearDataFormatCheck = linear.HtmlDataFormatCheck,
            HtmlLinearStandard = linear.HtmlStandard,
            HtmlLinearGradeDisplay = linear.HtmlOverallGradeDisplay,
            HtmlLinearAperture = linear.HtmlAperture,
            HtmlLinearWavelength = linear.HtmlWavelength,
            // The linear Webscan summary's sixth cell is Notes; the legacy
            // multi-mode report slot is named HtmlLinearLighting.
            HtmlLinearLighting = linear.HtmlNotes ?? linear.HtmlLighting,
            HtmlLinearFormalGrade = linear.HtmlFormalGrade,
            IsStandaloneLinear = false,
            DataSourceExceptions =
                $"{twoD.DataSourceExceptions}; Webscan composite: one linear + one 2D report",
        };
    }

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