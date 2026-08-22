namespace DeviceInterface.Webscan;

using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using ExcelEngine.Models;

/// <summary>
/// Parses the Webscan TruCheck HTML export format confirmed from the TC-829
/// controlled report. The parser only copies values present in the report.
/// </summary>
public static partial class WebscanHtmlParser
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    public static WebscanHtmlReport ParseFile(string sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
            throw new ArgumentException("A Webscan report path is required.", nameof(sourcePath));

        string rawHtml = File.ReadAllText(sourcePath);
        return Parse(rawHtml, sourcePath);
    }

    public static WebscanHtmlReport Parse(string rawHtml, string sourcePath)
    {
        if (rawHtml is null) throw new ArgumentNullException(nameof(rawHtml));
        if (sourcePath is null) throw new ArgumentNullException(nameof(sourcePath));

        try
        {
            var text = HtmlText(rawHtml);
            if (!text.Contains("Webscan TruCheck", StringComparison.OrdinalIgnoreCase))
                return Failure(rawHtml, sourcePath, "Webscan TruCheck title was not found.");

            var rows = ExtractRows(rawHtml);
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in rows)
            {
                if (row.Cells.Count == 2 && row.Cells[0].Length > 0)
                    values[row.Cells[0]] = row.Cells[1];
            }

            string? verified = FindHeaderText(rawHtml, @"(?<value>(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\d{1,2}-[A-Za-z]{3}-\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M)");
            DateTime? verifiedDateTime = ParseVerifiedDate(verified);
            string? software = FindHeaderValue(rawHtml, "Software Version");
            string? serial = FindHeaderValue(rawHtml, "Serial Number");

            Row? gradeRow = rows.FirstOrDefault(r =>
                r.Cells.Count >= 6 &&
                IsGradeStandard(r.Cells[0]) &&
                IsGradeDisplay(r.Cells[1]) &&
                int.TryParse(r.Cells[2], NumberStyles.Integer, Invariant, out _) &&
                int.TryParse(r.Cells[3], NumberStyles.Integer, Invariant, out _));

            var quality = rows
                .Where(r => r.Cells.Count >= 6 &&
                            Regex.IsMatch(r.Cells[0], @"^\d+[a-z]?\.$",
                                RegexOptions.IgnoreCase) &&
                            !string.IsNullOrWhiteSpace(r.Cells[1]))
                .Select(ToQualityParameter)
                .Where(p => p is not null)
                .Cast<WebscanQualityParameter>()
                .ToArray();

            string? data = Get(values, "Data");
            string? symbology = Get(values, "Symbology");
            string? validationError = ValidateRequiredStructure(
                verifiedDateTime,
                software,
                serial,
                data,
                symbology,
                gradeRow,
                quality);
            if (validationError is not null)
                return Failure(rawHtml, sourcePath, validationError);

            string? imageReference = FindImageReference(rawHtml);
            string? imagePath = imageReference is null
                ? null
                : Path.GetFullPath(Path.Combine(
                    Path.GetDirectoryName(Path.GetFullPath(sourcePath)) ?? string.Empty,
                    imageReference));

            return new WebscanHtmlReport
            {
                SourceFilePath = sourcePath,
                SourceImagePath = imagePath,
                RawHtml = rawHtml,
                ParseSucceeded = true,
                VerifiedDisplay = verified,
                VerifiedDateTime = verifiedDateTime,
                SoftwareVersion = software,
                DeviceSerial = serial,
                Data = data,
                Symbology = symbology,
                VerifiedBy = Get(values, "Verified By"),
                CompanyName = Get(values, "Company Name"),
                ProductName = Get(values, "Product Name"),
                JobNumber = Get(values, "Job Number"),
                Standard = gradeRow?.Cells[0],
                OverallGradeDisplay = gradeRow?.Cells[1],
                ApertureDisplay = gradeRow?.Cells.ElementAtOrDefault(2),
                Aperture = ParseInt(gradeRow?.Cells.ElementAtOrDefault(2)),
                WavelengthDisplay = gradeRow?.Cells.ElementAtOrDefault(3),
                Wavelength = ParseInt(gradeRow?.Cells.ElementAtOrDefault(3)),
                Lighting = gradeRow?.Cells.ElementAtOrDefault(4),
                FormalGrade = gradeRow?.Cells.ElementAtOrDefault(5),
                MatrixSize = Get(values, "Matrix Size"),
                HorizontalBWGDisplay = Get(values, "Horizontal BWG"),
                HorizontalBWG = ParseDecimal(Get(values, "Horizontal BWG"), true),
                VerticalBWGDisplay = Get(values, "Vertical BWG"),
                VerticalBWG = ParseDecimal(Get(values, "Vertical BWG"), true),
                EncodedCharactersDisplay = Get(values, "Encoded characters"),
                EncodedCharacters = ParseInt(Get(values, "Encoded characters")),
                TotalCodewordsDisplay = Get(values, "Total Codewords"),
                TotalCodewords = ParseInt(Get(values, "Total Codewords")),
                DataCodewordsDisplay = Get(values, "Data Codewords"),
                DataCodewords = ParseInt(Get(values, "Data Codewords")),
                ImagePolarity = Get(values, "Image"),
                NominalXDim = Get(values, "Nominal X Dim"),
                ContrastUniformity = Get(values, "Contrast Uniformity"),
                QualityParameters = quality,
            };
        }
        catch (Exception ex)
        {
            return Failure(rawHtml, sourcePath, ex.Message);
        }
    }

    public static VerificationRecord ParseRecord(string rawHtml, string sourcePath)
        => Parse(rawHtml, sourcePath).ToVerificationRecord();

    public static SymbologyFamily MapSymbologyFamily(string symbology)
    {
        if (symbology.Contains("DataMatrix", StringComparison.OrdinalIgnoreCase))
            return SymbologyFamily.DataMatrix;
        if (symbology.Contains("QR", StringComparison.OrdinalIgnoreCase))
            return SymbologyFamily.QRCode;
        return SymbologyFamily.Unknown;
    }

    public static ImagePolarity MapImagePolarity(string? value)
        => value?.Trim() switch
        {
            "White on Black" => ImagePolarity.WhiteOnBlack,
            "Black on White" => ImagePolarity.BlackOnWhite,
            _ => ImagePolarity.Unknown,
        };

    internal static GradingResult? ParseGrade(string? display, string? measured = null, string? result = null)
    {
        if (string.IsNullOrWhiteSpace(display)) return null;
        Match match = GradeRegex().Match(display.Trim());
        if (!match.Success ||
            !decimal.TryParse(match.Groups["numeric"].Value, NumberStyles.Number, Invariant,
                out decimal numeric))
            return null;

        return GradingResult.FromLetterAndNumeric(
            match.Groups["letter"].Value,
            numeric,
            result ?? string.Empty,
            measured);
    }

    private static WebscanQualityParameter? ToQualityParameter(Row row)
    {
        string? measured = NullIfEmpty(row.Cells.ElementAtOrDefault(2));
        string? gradeDisplay = NullIfEmpty(row.Cells.ElementAtOrDefault(3));
        string? secondary = NullIfEmpty(row.Cells.ElementAtOrDefault(4));
        string? result = NullIfEmpty(row.Cells.ElementAtOrDefault(5));

        return new WebscanQualityParameter
        {
            Number = row.Cells[0],
            Name = row.Cells[1],
            MeasuredValue = measured,
            MeasuredNumeric = ParseDecimal(measured, false),
            MeasuredPercent = ParseDecimal(measured, true),
            GradeDisplay = gradeDisplay,
            Grade = ParseGrade(gradeDisplay, measured, result),
            SecondaryValue = secondary,
            Result = result,
        };
    }

    private static string? ValidateRequiredStructure(
        DateTime? verifiedDateTime,
        string? softwareVersion,
        string? deviceSerial,
        string? data,
        string? symbology,
        Row? gradeRow,
        IReadOnlyCollection<WebscanQualityParameter> quality)
    {
        var missing = new List<string>();
        if (verifiedDateTime is null) missing.Add("verified timestamp");
        if (string.IsNullOrWhiteSpace(softwareVersion)) missing.Add("software version");
        if (string.IsNullOrWhiteSpace(deviceSerial)) missing.Add("serial number");
        if (string.IsNullOrWhiteSpace(data)) missing.Add("summary Data");
        if (string.IsNullOrWhiteSpace(symbology)) missing.Add("summary Symbology");
        if (gradeRow is null) missing.Add("Verification Grades row");
        if (quality.Count == 0) missing.Add("ISO quality-parameter rows");

        string[] requiredQuality =
        [
            "Unused Error Correction (UEC)",
            "Symbol Contrast (SC)",
            "Average Grade (AG)",
            "DECODE",
        ];
        foreach (string name in requiredQuality)
        {
            if (!quality.Any(parameter =>
                parameter.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                missing.Add($"quality row '{name}'");
        }

        return missing.Count == 0
            ? null
            : "Incomplete Webscan TruCheck report: missing " + string.Join(", ", missing) + ".";
    }

    private static List<Row> ExtractRows(string html)
    {
        var rows = new List<Row>();
        foreach (Match match in RowRegex().Matches(html))
        {
            var cells = CellRegex().Matches(match.Groups["body"].Value)
                .Select(m => CleanText(m.Groups["body"].Success
                    ? m.Groups["body"].Value
                    : string.Empty))
                .ToList();

            // A self-closing <td /> has no body capture but still produces a
            // regex match, preserving the six-column quality-row positions.
            if (cells.Count > 0)
                rows.Add(new Row(cells));
        }
        return rows;
    }

    private static string? FindHeaderText(string html, string pattern)
    {
        foreach (Match header in HeaderRegex().Matches(html))
        {
            Match match = Regex.Match(
                CleanText(header.Groups["body"].Value),
                pattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (match.Success)
                return match.Groups["value"].Value.Trim();
        }
        return null;
    }

    private static string? FindHeaderValue(string html, string label)
    {
        foreach (Match header in HeaderRegex().Matches(html))
        {
            string text = CleanText(header.Groups["body"].Value);
            if (text.StartsWith(label + ":", StringComparison.OrdinalIgnoreCase))
                return text[(label.Length + 1)..].Trim();
        }
        return null;
    }

    private static string? FindImageReference(string html)
    {
        foreach (Match match in ImageRegex().Matches(html))
        {
            string? alt = match.Groups["alt"].Value;
            string src = WebUtility.HtmlDecode(match.Groups["src"].Value).Trim();
            if (alt.Contains("Symbol", StringComparison.OrdinalIgnoreCase))
                return src;
        }
        return null;
    }

    private static DateTime? ParseVerifiedDate(string? display)
        => DateTime.TryParseExact(
            display,
            ["ddd dd-MMM-yyyy hh:mm:ss tt", "ddd d-MMM-yyyy hh:mm:ss tt"],
            Invariant,
            DateTimeStyles.None,
            out var value)
            ? value
            : null;

    private static bool IsGradeStandard(string value)
        => value.Contains("15415", StringComparison.OrdinalIgnoreCase);

    private static bool IsGradeDisplay(string value)
        => GradeRegex().IsMatch(value.Trim());

    private static int? ParseInt(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        Match match = Regex.Match(value, @"-?\d+");
        return match.Success && int.TryParse(match.Value, NumberStyles.Integer, Invariant, out int result)
            ? result
            : null;
    }

    private static decimal? ParseDecimal(string? value, bool stripPercent)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string candidate = stripPercent
            ? value.Replace("%", string.Empty, StringComparison.Ordinal)
            : value;
        Match match = Regex.Match(candidate, @"[-+]?\d+(?:[.,]\d+)?");
        return match.Success &&
               decimal.TryParse(match.Value.Replace(',', '.'),
                   NumberStyles.Number, Invariant, out decimal result)
            ? result
            : null;
    }

    private static string? Get(
        IReadOnlyDictionary<string, string> values,
        string key)
        => values.TryGetValue(key, out var value) ? value : null;

    private static string CleanText(string value)
    {
        string withoutTags = Regex.Replace(value, "<[^>]+>", " ");
        return Regex.Replace(WebUtility.HtmlDecode(withoutTags), @"\s+", " ").Trim();
    }

    private static string HtmlText(string html) => CleanText(html);

    private static string? NullIfEmpty(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static WebscanHtmlReport Failure(
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

    private sealed record Row(IReadOnlyList<string> Cells);

    [GeneratedRegex(@"<tr\b[^>]*>(?<body>.*?)</tr\s*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex RowRegex();

    [GeneratedRegex(@"<td\b[^>]*?(?:/\s*>|>(?<body>.*?)</td\s*>)",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex CellRegex();

    [GeneratedRegex(@"<img\b[^>]*src\s*=\s*[""'](?<src>[^""']+)[""'][^>]*alt\s*=\s*[""'](?<alt>[^""']*)[""'][^>]*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex ImageRegex();

    [GeneratedRegex(@"<h2\b[^>]*>(?<body>.*?)</h2\s*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex HeaderRegex();

    [GeneratedRegex(@"(?<letter>[A-F])\s*\((?<numeric>\d+(?:\.\d+)?)\)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex GradeRegex();
}