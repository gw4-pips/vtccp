namespace DeviceInterface.Webscan;

using System.Globalization;
using System.Net;
using System.Text;
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
        => ParseInternal(rawHtml, sourcePath, 1);

    /// <summary>
    /// Parses a Webscan export containing two or more independent symbol reports.
    /// The native reports remain separate and retain their ordinal/image
    /// provenance. Exactly one linear plus one supported 2D report is projected
    /// into the legacy dual-symbology fields.
    /// </summary>
    public static WebscanHtmlCompositeReport ParseComposite(
        string rawHtml,
        string sourcePath)
    {
        if (rawHtml is null) throw new ArgumentNullException(nameof(rawHtml));
        if (sourcePath is null) throw new ArgumentNullException(nameof(sourcePath));

        try
        {
            MatchCollection headers = SymbolReportHeaderRegex().Matches(rawHtml);
            if (headers.Count < 2)
                return WebscanHtmlCompositeReport.Failure(
                    rawHtml,
                    sourcePath,
                    "Webscan multi-symbol report must contain at least two symbol reports.");

            int firstStart = headers[0].Index;
            string sharedHeader = rawHtml[..firstStart];
            var reports = new List<WebscanHtmlReport>(headers.Count);
            for (int index = 0; index < headers.Count; index++)
            {
                int start = headers[index].Index;
                int length = index + 1 < headers.Count
                    ? headers[index + 1].Index - start
                    : rawHtml.Length - start;
                string segment = rawHtml.Substring(start, length);
                // The shared Webscan header contains the timestamp, software
                // version, and serial number; prepend it so the normal
                // single-report parser remains the canonical field mapper.
                reports.Add(ParseInternal(
                    sharedHeader + segment,
                    sourcePath,
                    index + 1));
            }

            if (reports.Any(r => !r.ParseSucceeded))
            {
                string error = string.Join(
                    " | ",
                    reports.Select((r, i) =>
                        $"Symbol {i + 1}: {r.ParseError ?? "parse failed"}"));
                return WebscanHtmlCompositeReport.Failure(rawHtml, sourcePath, error);
            }

            WebscanHtmlReport[] linearReports = reports.Where(
                r => MapSymbologyFamily(r.Symbology ?? string.Empty) == SymbologyFamily.Linear1D).ToArray();
            WebscanHtmlReport[] twoDReports = reports.Where(
                r => r.Symbology is not null &&
                     MapSymbologyFamily(r.Symbology) is
                         SymbologyFamily.DataMatrix or SymbologyFamily.QRCode).ToArray();
            return new WebscanHtmlCompositeReport
            {
                SourceFilePath = sourcePath,
                RawHtml = rawHtml,
                ParseSucceeded = true,
                SymbolReports = reports,
                LinearReport = linearReports[0],
                TwoDReport = twoDReports[0],
            };
        }
        catch (Exception ex)
        {
            return WebscanHtmlCompositeReport.Failure(rawHtml, sourcePath, ex.Message);
        }
    }

    /// <summary>
    /// Canonical name for importing a Webscan multi-symbol export. The legacy
    /// ParseComposite entry point remains for serialized callers and older
    /// integrations; it does not imply a GS1 Composite Component.
    /// </summary>
    public static WebscanHtmlMultiSymbolReport ParseMultiSymbol(
        string rawHtml,
        string sourcePath)
        => WebscanHtmlMultiSymbolReport.From(ParseComposite(rawHtml, sourcePath));

    private static WebscanHtmlReport ParseInternal(
        string rawHtml,
        string sourcePath,
        int imageOrdinal)
    {
        if (rawHtml is null) throw new ArgumentNullException(nameof(rawHtml));
        if (sourcePath is null) throw new ArgumentNullException(nameof(sourcePath));

        try
        {
            var text = HtmlText(rawHtml);
            if (!text.Contains("Webscan TruCheck", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("Symbol Verification Report",
                    StringComparison.OrdinalIgnoreCase))
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
            string? symbology = Get(values, "Symbology");
            bool isLinear = symbology is not null &&
                MapSymbologyFamily(symbology) == SymbologyFamily.Linear1D;

            Row? gradeRow = rows.FirstOrDefault(r =>
                r.Cells.Count >= 5 &&
                IsGradeStandard(r.Cells[0]) &&
                IsGradeDisplay(r.Cells[1]) &&
                int.TryParse(r.Cells[2], NumberStyles.Integer, Invariant, out _) &&
                int.TryParse(r.Cells[3], NumberStyles.Integer, Invariant, out _));

            var quality = rows
                .Where(r => r.Cells.Count >= 5 &&
                            Regex.IsMatch(r.Cells[0], @"^\d+[a-z]?\.$",
                                RegexOptions.IgnoreCase) &&
                            !string.IsNullOrWhiteSpace(r.Cells[1]))
                .Select(ToQualityParameter)
                .Where(p => p is not null)
                .Cast<WebscanQualityParameter>()
                .ToArray();

            string? data = Get(values, "Data");
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

            (string? imagePath, WebscanImageProvenance imageProvenance,
                string? imageMimeType, string? imageBase64) =
                ResolveSourceImage(rawHtml, sourcePath, imageOrdinal);
            DataFormatCheckResult? dataFormatCheck = ExtractDataFormatCheck(rawHtml);

            return new WebscanHtmlReport
            {
                SourceFilePath = sourcePath,
                SourceImagePath = imagePath,
                SourceImageProvenance = imageProvenance,
                SourceImageMimeType = imageMimeType,
                SourceImageBase64 = imageBase64,
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
                Lighting = isLinear ? null : gradeRow?.Cells.ElementAtOrDefault(4),
                Notes = isLinear ? gradeRow?.Cells.ElementAtOrDefault(5) : null,
                FormalGrade = gradeRow?.Cells.ElementAtOrDefault(isLinear ? 4 : 5),
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
                DataFormatCheck = dataFormatCheck,
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
        string compact = Regex.Replace(symbology, @"[\s\-_/]", string.Empty);
        if (compact.Equals("UPCA", StringComparison.OrdinalIgnoreCase) ||
            compact.Equals("UPCE", StringComparison.OrdinalIgnoreCase) ||
            compact.Equals("EAN8", StringComparison.OrdinalIgnoreCase) ||
            compact.Equals("EAN13", StringComparison.OrdinalIgnoreCase))
            return SymbologyFamily.Linear1D;
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
        bool hasSecondaryValue = row.Cells.Count >= 6;
        string? secondary = hasSecondaryValue
            ? NullIfEmpty(row.Cells.ElementAtOrDefault(4))
            : null;
        string? result = NullIfEmpty(row.Cells.ElementAtOrDefault(hasSecondaryValue ? 5 : 4));

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

        bool isLinear = symbology is not null &&
            MapSymbologyFamily(symbology) == SymbologyFamily.Linear1D;
        if (isLinear)
            return missing.Count == 0
                ? null
                : "Incomplete Webscan TruCheck report: missing " + string.Join(", ", missing) + ".";

        // The report-level Verification Grades row is the Webscan-provided
        // overall grade. Some valid QR exports do not include a separate
        // Average Grade (AG) quality row, so AG must remain optional and blank
        // when it is absent rather than causing the whole literal export to be
        // rejected.
        string[] requiredQuality =
        [
            "Unused Error Correction (UEC)",
            "Symbol Contrast (SC)",
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

    private static (string? Path, WebscanImageProvenance Provenance,
        string? MimeType, string? Base64) ResolveSourceImage(
        string html,
        string sourcePath,
        int imageOrdinal)
    {
        string directory = Path.GetDirectoryName(Path.GetFullPath(sourcePath)) ?? string.Empty;
        string? imageReference = FindImageReference(html);
        if (imageReference is not null)
        {
            if (TryResolveReportLocalImagePath(
                    directory, imageReference, out string candidate) &&
                TryReadImage(candidate, out string mimeType, out string base64))
                return (candidate, WebscanImageProvenance.EmbeddedHtml,
                    mimeType, base64);
        }

        foreach (string candidate in SiblingImageCandidates(sourcePath, imageOrdinal))
        {
            if (TryResolveReportLocalImagePath(directory, candidate, out string localCandidate) &&
                TryReadImage(localCandidate, out string mimeType, out string base64))
                return (localCandidate, WebscanImageProvenance.SiblingExport, mimeType, base64);
        }

        return (null, WebscanImageProvenance.None, null, null);
    }

    private static string? FindImageReference(string html)
    {
        foreach (Match match in ImageRegex().Matches(html))
        {
            string alt = match.Groups["alt"].Value;
            string src = WebUtility.HtmlDecode(match.Groups["src"].Value).Trim();
            if (alt.Contains("Symbol", StringComparison.OrdinalIgnoreCase))
                return src;
        }
        return null;
    }

    private static bool TryResolveReportLocalImagePath(
        string reportDirectory,
        string imageReference,
        out string imagePath)
    {
        imagePath = string.Empty;
        if (string.IsNullOrWhiteSpace(imageReference) || Path.IsPathRooted(imageReference))
            return false;

        try
        {
            string fullDirectory = Path.GetFullPath(reportDirectory);
            string candidate = Path.GetFullPath(Path.Combine(fullDirectory, imageReference));
            string relative = Path.GetRelativePath(fullDirectory, candidate);
            if (Path.IsPathRooted(relative) ||
                relative.Equals("..", StringComparison.Ordinal) ||
                relative.StartsWith(".." + Path.DirectorySeparatorChar,
                    StringComparison.Ordinal) ||
                relative.StartsWith(".." + Path.AltDirectorySeparatorChar,
                    StringComparison.Ordinal) ||
                ContainsSymbolicLink(fullDirectory, relative))
            {
                return false;
            }

            imagePath = candidate;
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static bool ContainsSymbolicLink(string fullDirectory, string relativePath)
    {
        var current = new DirectoryInfo(fullDirectory);
        if (current.LinkTarget is not null) return true;

        string[] segments = relativePath.Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < segments.Length; index++)
        {
            string currentPath = Path.Combine(
                current.FullName,
                Path.Combine(segments.Take(index + 1).ToArray()));
            if (index == segments.Length - 1)
            {
                if (new FileInfo(currentPath).LinkTarget is not null)
                    return true;
            }
            else if (new DirectoryInfo(currentPath).LinkTarget is not null)
            {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<string> SiblingImageCandidates(
        string sourcePath,
        int imageOrdinal)
    {
        string directory = Path.GetDirectoryName(Path.GetFullPath(sourcePath)) ?? string.Empty;
        string stem = Path.GetFileNameWithoutExtension(sourcePath);
        string[] extensions = [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"];

        foreach (string extension in extensions)
            yield return stem + $".Image{imageOrdinal}" + extension;

        // Webscan sanitizes the HTML export name differently from the image
        // export. For example:
        //   Report._1787402227622.html
        //   Report.Image1_1787402227622.jpg
        int separator = stem.LastIndexOf("._", StringComparison.Ordinal);
        if (separator > 0 && separator + 2 < stem.Length)
        {
            string prefix = stem[..separator];
            string suffix = stem[(separator + 1)..]; // includes the underscore
            foreach (string extension in extensions)
                yield return prefix + $".Image{imageOrdinal}" + suffix + extension;
        }

        // TC-829 linear exports append a numeric export id after the report
        // stem, while the sibling image inserts Image1 before that suffix:
        //   UPCA-...-696114704318_1787446139035.html
        //   UPCA-...-696114704318Image1_1787446139035.jpg
        Match timestampedStem = Regex.Match(
            stem,
            @"^(?<prefix>.+)_(?<id>\d{10,})$",
            RegexOptions.CultureInvariant);
        if (timestampedStem.Success)
        {
            string prefix = timestampedStem.Groups["prefix"].Value;
            string id = timestampedStem.Groups["id"].Value;
            foreach (string extension in extensions)
                yield return prefix + $"Image{imageOrdinal}_" + id + extension;
        }
    }

    private static bool TryReadImage(string path, out string mimeType, out string base64)
    {
        mimeType = string.Empty;
        base64 = string.Empty;
        try
        {
            if (!File.Exists(path)) return false;
            byte[] bytes = File.ReadAllBytes(path);
            mimeType = DetectImageMimeType(bytes);
            if (mimeType.Length == 0) return false;
            base64 = Convert.ToBase64String(bytes);
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static string DetectImageMimeType(byte[] bytes)
    {
        if (bytes.Length >= 3 &&
            bytes[0] == 0xff && bytes[1] == 0xd8 && bytes[2] == 0xff)
            return "image/jpeg";
        if (bytes.Length >= 8 &&
            bytes.AsSpan(0, 8).SequenceEqual(
                new byte[] { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a }))
            return "image/png";
        if (bytes.Length >= 6 &&
            (Encoding.ASCII.GetString(bytes, 0, 6) == "GIF87a" ||
             Encoding.ASCII.GetString(bytes, 0, 6) == "GIF89a"))
            return "image/gif";
        if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4d)
            return "image/bmp";
        if (bytes.Length >= 12 &&
            Encoding.ASCII.GetString(bytes, 0, 4) == "RIFF" &&
            Encoding.ASCII.GetString(bytes, 8, 4) == "WEBP")
            return "image/webp";
        return string.Empty;
    }

    private static DataFormatCheckResult? ExtractDataFormatCheck(string html)
    {
        foreach (Match tableMatch in TableRegex().Matches(html))
        {
            string table = tableMatch.Value;
            if (!CleanText(table).Contains("Data Format Check",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            string? standard = null;
            OverallPassFail overall = OverallPassFail.NotApplicable;
            foreach (Match header in TableHeaderRegex().Matches(table))
            {
                string text = CleanText(header.Groups["body"].Value);
                if (text.Contains("Data Format Check", StringComparison.OrdinalIgnoreCase))
                    continue;

                Match outcome = DfcOutcomeRegex().Match(text);
                if (outcome.Success)
                {
                    (standard, overall) = ParseDfcOutcome(text, outcome);
                    break;
                }
            }

            // Some Webscan versions use one dedicated <td> rather than a table
            // header for the outcome. Only accept a single-cell table row here:
            // a three-column data row can contain arbitrary literal colons and
            // PASS/FAIL text, none of which is an overall verifier outcome.
            if (overall == OverallPassFail.NotApplicable)
            {
                foreach (Match rowMatch in RowRegex().Matches(table))
                {
                    string rowMarkup = rowMatch.Groups["body"].Value;
                    if (rowMarkup.Contains("<th", StringComparison.OrdinalIgnoreCase))
                        continue;
                    MatchCollection cells = CellRegex().Matches(rowMarkup);
                    if (cells.Count != 1) continue;
                    string text = CleanText(cells[0].Groups["body"].Value);
                    Match outcome = DfcOutcomeRegex().Match(text);
                    if (!outcome.Success) continue;

                    (standard, overall) = ParseDfcOutcome(text, outcome);
                    break;
                }
            }

            var rows = new List<DataFormatCheckRow>();
            foreach (Match rowMatch in RowRegex().Matches(table))
            {
                string rowMarkup = rowMatch.Groups["body"].Value;
                if (rowMarkup.Contains("<th", StringComparison.OrdinalIgnoreCase))
                    continue;

                MatchCollection cells = CellRegex().Matches(rowMarkup);
                if (cells.Count < 3) continue;
                string name = CleanText(cells[0].Groups["body"].Value);
                string data = CleanText(cells[1].Groups["body"].Value);
                string check = CleanText(cells[2].Groups["body"].Value);
                if (name.Equals("Name", StringComparison.OrdinalIgnoreCase) &&
                    data.Equals("Data", StringComparison.OrdinalIgnoreCase) &&
                    check.Equals("Check", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (name.Length > 0)
                    rows.Add(new DataFormatCheckRow
                    {
                        Name = name,
                        Data = data,
                        Check = check,
                    });
            }

            return new DataFormatCheckResult
            {
                Standard = standard,
                Overall = overall,
                Rows = rows,
            };
        }

        return null;
    }

    private static (string Standard, OverallPassFail Overall) ParseDfcOutcome(
        string text,
        Match outcome)
        => (
            CleanDfcStandard(text[..outcome.Groups["colon"].Index]),
            outcome.Groups["outcome"].Value.Equals(
                "PASS", StringComparison.OrdinalIgnoreCase)
                ? OverallPassFail.Pass
                : OverallPassFail.Fail);

    private static string CleanDfcStandard(string value)
    {
        string trimmed = value.Trim();
        foreach (string marker in new[] { "GS1", "HIBCC", "ISO " })
        {
            int markerIndex = trimmed.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (markerIndex >= 0)
                return trimmed[markerIndex..].Trim();
        }
        return trimmed;
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
        => value.Contains("15415", StringComparison.OrdinalIgnoreCase) ||
           value.Contains("15416", StringComparison.OrdinalIgnoreCase) ||
           value.Contains("ANSI/ISO", StringComparison.OrdinalIgnoreCase);

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

    [GeneratedRegex(@"<table\b[^>]*>.*?</table\s*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex TableRegex();

    [GeneratedRegex(@"<th\b[^>]*>(?<body>.*?)</th\s*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex TableHeaderRegex();

    [GeneratedRegex(@"(?<standard>[A-Za-z][^:\r\n]{0,120}?)(?<colon>:)\s*(?<outcome>PASS|FAIL)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex DfcOutcomeRegex();

    [GeneratedRegex(@"<h2\b[^>]*>(?<body>.*?)</h2\s*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex HeaderRegex();

    [GeneratedRegex(@"<h1\b[^>]*>\s*Symbol\s+(?<number>\d+)\s+Verification\s+Report\s*</h1\s*>",
        RegexOptions.Singleline | RegexOptions.IgnoreCase)]
    private static partial Regex SymbolReportHeaderRegex();

    [GeneratedRegex(@"(?<letter>[A-F])\s*\((?<numeric>\d+(?:\.\d+)?)\)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex GradeRegex();
}