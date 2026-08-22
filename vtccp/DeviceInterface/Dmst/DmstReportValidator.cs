namespace DeviceInterface.Dmst;

using System.Globalization;
using System.Text.RegularExpressions;
using DeviceInterface.Rfid;
using ExcelEngine.Models;

/// <summary>
/// Cross-validates a push XML <see cref="VerificationRecord"/> against the
/// corresponding DMST HTML <see cref="DmstHtmlReport"/> for the same scan.
///
/// Two distinct operations per scan:
///
///   1. Supplemental merge — fields absent from push XML but present in the HTML
///      (QR_ECLevel, QR_MaskPattern, QR_ECI, ImagePolarity, DataCodewords,
///      ErrorCorrectionBudget, EncodedCharacters) are merged into the record and
///      catalogued in DataSourceExceptions.
///      When the HTML is multi-mode (IsMultiMode = true), the nine source-backed
///      Linear* fields
///      (LinearSymbology, LinearDecodedData, LinearOverallGrade, LinearFormalGrade,
///      LinearAperture, LinearWavelength, LinearLighting, LinearStandard,
///      LinearJpegImageBase64) are also populated.
///      LinearJpegImageBase64 is populated from record.RoiJpegImageBase64 when
///      the ROI frame was already captured (SDK-triggered scans); null in push-only
///      mode where IMAGE.SEND is not issued before MergeAndValidate runs.
///
///   2. Cross-validation — every field producible from BOTH sources is compared.
///      Mismatches → ValidationDiscrepancies + Debug log.
///      - Agreement: data-integrity credibility.
///      - Discrepancy: parser bug or firmware anomaly — Cognex bug-report candidate.
///
/// Confirmed from 2026-05-25 live HTML sample (QR GUID, fw 6.1.16_sr4):
///   ECLevel="M", DataMaskPattern="2", ECI="000003", ImagePolarity="Black on white"
///   DataCodewords=44, ErrorCorrectionBudget=26, EncodedCharacters=36(vs push 39).
///
/// EncodedCharacters: push XML eaLen is wrong (39 vs HTML authoritative 36).
/// HTML value is always taken as authoritative; discrepancy is always flagged.
/// </summary>
public static class DmstReportValidator
{
    /// <summary>
    /// Absolute tolerance for decimal field comparisons.
    /// Absorbs rounding differences between push XML and HTML rendering.
    /// </summary>
    public const decimal NumericTolerance = 0.05m;

    // ── Public entry point ────────────────────────────────────────────────────

    /// <summary>
    /// Merges supplemental fields from <paramref name="html"/> into
    /// <paramref name="record"/> and runs full cross-validation.
    /// Returns a new record instance with all merged and validated fields set.
    /// </summary>
    public static VerificationRecord MergeAndValidate(
        VerificationRecord record,
        DmstHtmlReport     html)
    {
        var exceptions    = new List<string>();
        var discrepancies = new List<string>();
        // Keep VCCS's Digital Link syntax result separate from the vendor DFC.
        // When HTML is present, use its literal decoded value as the validation input.
        string? gs1ValidationInput = html.HtmlDecodedData ?? record.DecodedData;
        DigitalLinkValidationResult? digitalLinkValidation = null;
        bool truCheckDfcPassed =
            html.ScrapedDataFormatCheck is { Rows.Count: > 0, Overall: OverallPassFail.Pass };
        if (!truCheckDfcPassed)
        {
            digitalLinkValidation =
                DeviceInterface.Validation.VccsDigitalLinkValidationService
                    .Validate(gs1ValidationInput);
            if (digitalLinkValidation.Status == DigitalLinkValidationStatus.NotApplicable &&
                DeviceInterface.Validation.VccsDigitalLinkValidationService
                    .LooksLikeGs1ElementString(gs1ValidationInput))
            {
                digitalLinkValidation =
                    DeviceInterface.Validation.VccsDigitalLinkValidationService
                        .ValidateElementString(gs1ValidationInput);
            }
        }
        bool hasCorrelatedFilesystemHtml =
            html.ParseSucceeded &&
            !html.HasSyntheticSourcePath &&
            !string.IsNullOrWhiteSpace(html.SourceFilePath) &&
            !string.IsNullOrWhiteSpace(html.HtmlSourceFileName) &&
            string.Equals(
                Path.GetFileName(html.SourceFilePath.Replace('\\', '/')),
                html.HtmlSourceFileName,
                StringComparison.Ordinal);

        // Carry forward any existing DataSourceExceptions (e.g. from SYMBOL.RESULT FULL).
        if (!string.IsNullOrEmpty(record.DataSourceExceptions))
            exceptions.Add(record.DataSourceExceptions);

        // ── 1a. Merge: four permanently unresolvable from push XML ────────────

        string?       qrECLevel     = record.QR_ECLevel;
        string?       qrMaskPattern = record.QR_MaskPattern;
        string?       qrECI         = record.QR_ECI;
        ImagePolarity imagePolarity  = record.ImagePolarity;

        if (string.IsNullOrEmpty(qrECLevel) && !string.IsNullOrEmpty(html.ECLevel))
        {
            qrECLevel = html.ECLevel;
            exceptions.Add("QR_ECLevel:HtmlReport");
        }
        if (string.IsNullOrEmpty(qrMaskPattern) && !string.IsNullOrEmpty(html.DataMaskPattern))
        {
            qrMaskPattern = html.DataMaskPattern;
            exceptions.Add("QR_MaskPattern:HtmlReport");
        }
        if (string.IsNullOrEmpty(qrECI) && !string.IsNullOrEmpty(html.ECI))
        {
            qrECI = html.ECI;
            exceptions.Add("QR_ECI:HtmlReport");
        }
        if (imagePolarity == ImagePolarity.Unknown && !string.IsNullOrEmpty(html.ImagePolarity))
        {
            var parsed = ParseImagePolarity(html.ImagePolarity);
            if (parsed != ImagePolarity.Unknown)
            {
                imagePolarity = parsed;
                exceptions.Add("ImagePolarity:HtmlReport");
            }
        }

        // ── 1b. Merge: bonus fields (empty in push XML, present in HTML) ──────

        int? dataCodewords         = record.DataCodewords;
        int? errorCorrectionBudget = record.ErrorCorrectionBudget;
        int? encodedCharacters     = record.EncodedCharacters;

        if ((dataCodewords is null or 0) && html.DataCodewords.HasValue)
        {
            dataCodewords = html.DataCodewords;
            exceptions.Add("DataCodewords:HtmlReport");
        }
        if ((errorCorrectionBudget is null or 0) && html.ErrorCorrectionBudget.HasValue)
        {
            errorCorrectionBudget = html.ErrorCorrectionBudget;
            exceptions.Add("ErrorCorrectionBudget:HtmlReport");
        }
        // EncodedCharacters: push XML eaLen fallback is WRONG. HTML is authoritative.
        // Always take HTML value; always flag discrepancy if push differed.
        if (html.EncodedCharacters.HasValue)
        {
            if (encodedCharacters.HasValue && encodedCharacters != html.EncodedCharacters)
                discrepancies.Add(
                    $"EncodedCharacters:Push={encodedCharacters},Html={html.EncodedCharacters}");
            encodedCharacters = html.EncodedCharacters;
            exceptions.Add("EncodedCharacters:HtmlReport");
        }

        // ── 1c. Merge: linear symbol (multi-mode only) ────────────────────────
        //
        // When the HTML report covers a multi-mode scan (EAN/UPC + 2D), populate
        // the source-backed Linear* fields. Single-mode scans leave these null.
        //
        // LinearJpegImageBase64: HTML reports carry no image data, but the ROI frame
        // captured by IMAGE.SEND (DeviceSession.AttachRoiImageAsync) is already
        // attached to the record as RoiJpegImageBase64 before MergeAndValidate runs
        // in the SDK-triggered flow.  Reuse it as the linear image so the PDF
        // dual-image section renders (ROI frame on the left, 2D crop on the right).
        // In push-only mode (HttpEventSubscriber, no SDK) RoiJpegImageBase64 is null
        // and LinearJpegImageBase64 stays null — the PDF falls back to single-image.
        //
        // Data Format Check values are never reconstructed locally. The report may
        // only display DFC rows scraped from the TruCheck HTML.

        string?              linearSymbology      = null;
        string?              linearDecodedData    = null;
        GradingResult?       linearOverallGrade   = null;
        string?              linearFormalGrade    = null;
        int?                 linearAperture       = null;
        int?                 linearWavelength     = null;
        string?              linearLighting       = null;
        string?              linearStandard       = null;
        string?              linearJpegImageBase64 = null;

        if (html.IsMultiMode && !string.IsNullOrWhiteSpace(html.LinearSymbology))
        {
            linearSymbology    = html.LinearSymbology;
            linearDecodedData  = html.LinearDecodedData;
            linearFormalGrade  = html.LinearFormalGrade;
            linearAperture     = html.LinearAperture;
            linearWavelength   = html.LinearWavelength;
            linearLighting     = html.LinearLighting;
            linearStandard     = html.LinearStandard ?? "ISO/IEC 15416";

            // Use the ROI frame (IMAGE.SEND result, already on the record) as the
            // linear image.  This gives the PDF dual-image section something to
            // render on the left side.  Null in push-only mode — graceful fallback.
            linearJpegImageBase64 = record.RoiJpegImageBase64;

            if (!string.IsNullOrEmpty(html.LinearOverallGrade))
            {
                decimal linearNumeric = html.LinearOverallGradeNumeric
                    ?? LetterToNumericGrade(html.LinearOverallGrade);
                linearOverallGrade = GradingResult.FromLetterAndNumeric(
                    html.LinearOverallGrade,
                    linearNumeric,
                    // Pass/fail determined from the actual numeric — not from the letter
                    // midpoint — so a fractional B/2.5 against a 3.0 threshold correctly
                    // fails rather than being rounded up to B's midpoint 3.0.
                    DeterminePassFailNumeric(linearNumeric, record.MinPassGrade, record.MinPassRaw),
                    value: null);
            }

            exceptions.Add("LinearSymbology:HtmlReport");
            exceptions.Add("LinearOverallGrade:HtmlReport");
            if (!string.IsNullOrEmpty(html.LinearFormalGrade))
                exceptions.Add("LinearFormalGrade:HtmlReport");

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] Multi-mode linear: symb={linearSymbology} " +
                $"grade={linearOverallGrade?.LetterGradeString ?? "null"} " +
                $"img={linearJpegImageBase64?.Length.ToString() ?? "null"} chars " +
                $"dfc={html.ScrapedDataFormatCheck?.Rows.Count ?? 0} HTML rows");
        }

        // ── 2. Cross-validation ───────────────────────────────────────────────

        CompareString ("OverallGrade",   record.OverallGrade?.LetterGradeString,
                                         html.OverallGrade,                      discrepancies);
        CompareString ("MatrixSize",     NormaliseMatrixSize(record.MatrixSize),
                                         NormaliseMatrixSize(html.MatrixSize),   discrepancies);
        // NominalXDim_2D (decimal, e.g. 12.6) vs HTML "12.6 mil" (parse first token).
        CompareDecimal("NomXDim_2D",     record.NominalXDim_2D,
                                         ParseNominalXDimHtml(html.NominalXDim), discrepancies);
        CompareDecimal("ANU%",           record.ANU_Percent,
                                         html.ANUPercent,                        discrepancies);
        CompareDecimal("GNU%",           record.GNU_Percent,
                                         html.GNUPercent,                        discrepancies);
        CompareDecimal("UEC%",           record.UEC_Percent,
                                         html.UECPercent,                        discrepancies);
        CompareDecimal("SC%",            record.SC_Percent,
                                         html.SCPercent,                         discrepancies);
        CompareDecimal("HBwg%",          record.HorizontalBWG,
                                         html.HorizontalBWG,                     discrepancies);
        CompareDecimal("VBwg%",          record.VerticalBWG,
                                         html.VerticalBWG,                       discrepancies);
        CompareInt    ("TotalCodewords", record.TotalCodewords,
                                         html.TotalCodewords,                    discrepancies);
        CompareInt    ("ErrorsCorrected",record.ErrorsCorrected,
                                         html.ErrorsCorrected,                   discrepancies);
        CompareInt    ("ErrCapUsed",     record.ErrorCapacityUsed,
                                         html.ErrorCapacityUsed,                 discrepancies);
        // ImagePolarity: compare push enum → canonical string vs HTML literal.
        CompareString ("ImagePolarity",  ImagePolarityToHtmlString(record.ImagePolarity),
                                         html.ImagePolarity,                     discrepancies);

        // Log.
        if (discrepancies.Count > 0)
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] {discrepancies.Count} discrepancy(s) — " +
                $"{record.VerificationDateTime:HH:mm:ss}:\n  " +
                string.Join("\n  ", discrepancies));
        else
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] All fields match — {record.VerificationDateTime:HH:mm:ss}.");

        // ── 3. Produce merged record ──────────────────────────────────────────
        return record with
        {
            QR_ECLevel            = qrECLevel,
            QR_MaskPattern        = qrMaskPattern,
            QR_ECI                = qrECI,
            ImagePolarity         = imagePolarity,
            DataCodewords         = dataCodewords,
            ErrorCorrectionBudget = errorCorrectionBudget,
            EncodedCharacters     = encodedCharacters,

            // Linear symbol (multi-mode): null for single-mode scans.
            LinearSymbology       = linearSymbology,
            LinearDecodedData     = linearDecodedData,
            LinearOverallGrade    = linearOverallGrade,
            LinearFormalGrade     = linearFormalGrade,
            LinearAperture        = linearAperture,
            LinearWavelength      = linearWavelength,
            LinearLighting        = linearLighting,
            LinearStandard        = linearStandard,
            // Populated from RoiJpegImageBase64 (IMAGE.SEND ROI frame) when
            // the SDK-triggered flow captured it before MergeAndValidate ran.
            // Null in push-only mode — PDF falls back to single-image section.
            LinearJpegImageBase64 = linearJpegImageBase64,
            LinearDataFormatCheck = null,

            // Canonical verifier data only. If the TruCheck HTML does not contain
            // a DFC table, leave it unavailable rather than reconstructing one.
            DataFormatCheck         = html.ScrapedDataFormatCheck,

            DataSourceExceptions    = exceptions.Count > 0
                                      ? string.Join(";", exceptions)
                                      : null,
            ValidationDiscrepancies = discrepancies.Count > 0
                                      ? string.Join(";", discrepancies)
                                      : null,

            // Provenance from the matched HTML file — drives PDF Report Name and Date/Time.
            // SourceFilePath is captured before deletion; Path.GetFileName() works on a
            // deleted path string, so WebscanSourcePath is reliable even after the file is gone.
            WebscanSourcePath  = hasCorrelatedFilesystemHtml ? html.SourceFilePath : null,
            HtmlSourceFileName = hasCorrelatedFilesystemHtml ? html.HtmlSourceFileName : null,
            HtmlReportProvenance = hasCorrelatedFilesystemHtml
                                   ? HtmlReportProvenance.CorrelatedFilesystem
                                   : html.HasSyntheticSourcePath
                                   ? HtmlReportProvenance.HttpStreamOnly
                                   : HtmlReportProvenance.None,
            HtmlSourceProvenance = hasCorrelatedFilesystemHtml
                                   ? "DMST filesystem HTML report"
                                   : html.HasSyntheticSourcePath
                                   ? "HTTP stream placeholder — original DMST filename unavailable"
                                   : null,
            HtmlVerifiedString = html.HtmlVerifiedString,
            HtmlSymbology = html.HtmlSymbology,
            HtmlDecodedData = html.HtmlDecodedData,
            HtmlApplicationStandard = html.HtmlApplicationStandard,
            HtmlLinearStandard = html.HtmlLinearStandard,
            HtmlLinearGradeDisplay = html.HtmlLinearGradeDisplay,
            HtmlLinearAperture = html.HtmlLinearAperture,
            HtmlLinearWavelength = html.HtmlLinearWavelength,
            HtmlLinearLighting = html.HtmlLinearLighting,
            HtmlLinearFormalGrade = html.HtmlLinearFormalGrade,
            HtmlBarcodeImageBase64 = html.HtmlBarcodeImageBase64,
            HtmlDataFormatCheck = html.ScrapedDataFormatCheck,
            VccsDigitalLinkValidation = digitalLinkValidation,

            // Verbatim Verification Grades row — used directly in PDF BVG table.
            // HTML values always win over push-XML when present (push-XML often provides
            // wrong format for FormalGrade; Grade display format differs too).
            HtmlStandard            = html.HtmlStandard,
            HtmlOverallGradeDisplay = html.HtmlOverallGradeDisplay,
            HtmlAperture            = html.HtmlAperture,
            HtmlWavelength          = html.HtmlWavelength,
            HtmlLighting            = html.HtmlLighting,
            HtmlFormalGrade         = html.HtmlFormalGrade,
        };
    }

    // ── Comparison helpers ────────────────────────────────────────────────────

    private static void CompareString(
        string field, string? push, string? html, List<string> d)
    {
        if (push is null && html is null) return;
        if (string.Equals(push?.Trim(), html?.Trim(), StringComparison.OrdinalIgnoreCase)) return;
        d.Add($"{field}:Push={push ?? "null"},Html={html ?? "null"}");
    }

    private static void CompareDecimal(
        string field, decimal? push, decimal? html, List<string> d)
    {
        if (push is null && html is null) return;
        if (push is null || html is null)
        { d.Add($"{field}:Push={push?.ToString() ?? "null"},Html={html?.ToString() ?? "null"}"); return; }
        if (Math.Abs(push.Value - html.Value) <= NumericTolerance) return;
        d.Add($"{field}:Push={push},Html={html}");
    }

    private static void CompareInt(
        string field, int? push, int? html, List<string> d)
    {
        if (push is null && html is null) return;
        if (push == html) return;
        d.Add($"{field}:Push={push?.ToString() ?? "null"},Html={html?.ToString() ?? "null"}");
    }

    // ── Conversion helpers ────────────────────────────────────────────────────

    private static ImagePolarity ParseImagePolarity(string htmlValue)
        => htmlValue.Trim().ToLowerInvariant() switch
        {
            "black on white" => ImagePolarity.BlackOnWhite,
            "white on black" => ImagePolarity.WhiteOnBlack,
            _                => ImagePolarity.Unknown,
        };

    private static string? ImagePolarityToHtmlString(ImagePolarity p)
        => p switch
        {
            ImagePolarity.BlackOnWhite => "Black on white",
            ImagePolarity.WhiteOnBlack => "White on black",
            _                          => null,
        };

    private static string? NormaliseMatrixSize(string? raw)
        => raw?.Replace(" x ", "x", StringComparison.OrdinalIgnoreCase).Trim().ToLowerInvariant();

    /// <summary>
    /// Parses the numeric part of an HTML NominalXDim string such as "12.6 mil".
    /// Splits on the first space and parses the leading token as a decimal.
    /// Returns null if the string is absent or unparseable.
    /// </summary>
    private static decimal? ParseNominalXDimHtml(string? htmlValue)
    {
        if (string.IsNullOrEmpty(htmlValue)) return null;
        var token = htmlValue.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return decimal.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)
            ? v : null;
    }

    // ── Multi-mode linear symbol helpers ──────────────────────────────────────

    /// <summary>
    /// Converts a letter grade string (A/B/C/D/F) to its midpoint numeric ISO 15416 value.
    /// Used when only the letter is available (no numeric from the HTML).
    /// </summary>
    private static decimal LetterToNumericGrade(string letter)
        => letter.Trim().ToUpperInvariant() switch
        {
            "A" => 4.0m,
            "B" => 3.0m,
            "C" => 2.0m,
            "D" => 1.0m,
            "F" => 0.0m,
            _   => 0.0m,
        };

    /// <summary>
    /// Determines PASS/FAIL for a linear grade using the actual parsed numeric value,
    /// not a letter-midpoint approximation.  A fractional "B/2.5" grade against a
    /// 3.0 threshold correctly fails; "B/3.5" against the same threshold passes.
    ///
    /// Threshold resolution order:
    ///   1. <paramref name="minPassRaw"/> — already-parsed decimal (preferred).
    ///   2. <paramref name="minPassGrade"/> string — decimal ("1.5") or letter ("C").
    ///      Letter thresholds use ISO 15416 band lower bounds:
    ///        A ≥ 3.5,  B ≥ 2.5,  C ≥ 1.5,  D ≥ 0.5.
    ///   3. Default — anything above 0.0 passes; 0.0 (F band) fails.
    /// </summary>
    private static string DeterminePassFailNumeric(
        decimal numericGrade,
        string? minPassGrade,
        decimal? minPassRaw)
    {
        // Priority 1: numeric threshold already parsed from the record
        if (minPassRaw.HasValue)
            return numericGrade >= minPassRaw.Value ? "PASS" : "FAIL";

        // Priority 2: parse MinPassGrade string — decimal or letter
        if (!string.IsNullOrWhiteSpace(minPassGrade) && minPassGrade != "NA")
        {
            string raw = minPassGrade.Trim().TrimStart('>');

            // Decimal threshold (e.g. "1.5", "3.0")
            if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture,
                    out decimal minNumeric))
                return numericGrade >= minNumeric ? "PASS" : "FAIL";

            // Letter threshold (e.g. "C", "D") — ISO 15416 band lower bound
            decimal letterFloor = raw.ToUpperInvariant() switch
            {
                "A" => 3.5m,   // A band: 3.5 – 4.0
                "B" => 2.5m,   // B band: 2.5 – 3.4
                "C" => 1.5m,   // C band: 1.5 – 2.4
                "D" => 0.5m,   // D band: 0.5 – 1.4
                _   => -1.0m,
            };
            if (letterFloor >= 0m)
                return numericGrade >= letterFloor ? "PASS" : "FAIL";
        }

        // No threshold: F band (0.0) fails; everything else passes.
        return numericGrade > 0.0m ? "PASS" : "FAIL";
    }

    /// <summary>
    /// Builds a <see cref="DataFormatCheckResult"/> for a 2D GS1-encoded symbol
    /// (QR Code or DataMatrix) by parsing AI (01) GTIN-14 and AI (21) Serial Number
    /// from the barcode's decoded data string.
    ///
    /// Handles both GS1 Digital Link URIs and GS1 Element Strings via
    /// <see cref="RfidValidator.ExtractAi01"/> / <see cref="RfidValidator.ExtractAi21"/>.
    ///
    /// Returns null when:
    ///   • The symbology is not recognised as a 2D GS1 carrier.
    ///   • AI (01) is absent or not a 14-digit string.
    ///   • DecodedData is null/empty.
    /// </summary>
    internal static DataFormatCheckResult? BuildDataFormatCheck(VerificationRecord record)
    {
        // Only attempt 2D GS1 symbols.
        string? symb = record.Symbology;
        if (string.IsNullOrWhiteSpace(symb)) return null;
        bool is2D = symb.Contains("DataMatrix", StringComparison.OrdinalIgnoreCase)
                 || symb.Contains("QR",         StringComparison.OrdinalIgnoreCase);
        if (!is2D) return null;

        // Extract GTIN-14 (AI 01).
        string? gtin14 = RfidValidator.ExtractAi01(record.DecodedData);
        if (gtin14 is null || gtin14.Length != 14 || !gtin14.All(char.IsAsciiDigit))
            return null;

        bool   checkOk    = ValidateGs1CheckDigit(gtin14);
        string checkResult = checkOk ? "PASS" : "FAIL";
        string checkDigit  = gtin14[^1].ToString();
        string gtinBody    = gtin14[..^1];   // 13 digits without the check digit

        var rows = new List<DataFormatCheckRow>
        {
            new() { Name = "AI (01) GTIN-14", Data = gtinBody, Check = checkResult },
            new() { Name = "Check Digit",      Data = checkDigit, Check = checkResult },
        };

        // Extract Serial Number (AI 21) if present.
        string? serial = RfidValidator.ExtractAi21(record.DecodedData);
        if (!string.IsNullOrWhiteSpace(serial))
            rows.Add(new() { Name = "AI (21) Serial", Data = serial, Check = "PASS" });

        return new DataFormatCheckResult
        {
            Overall  = checkOk ? OverallPassFail.Pass : OverallPassFail.Fail,
            Standard = "GS1 Application Data Format",
            Rows     = rows,
        };
    }

    /// <summary>
    /// Builds a <see cref="DataFormatCheckResult"/> carrying a GTIN row and a
    /// check-digit validation row for EAN/UPC symbologies.
    ///
    /// The GS1 check-digit algorithm (Luhn-mod-10 with alternating weights 1/3)
    /// is inlined here to avoid a dependency on OcrEngine from DeviceInterface.
    ///
    /// Returns null when <paramref name="decodedData"/> is absent or is not a
    /// digit-only string of the expected length for the stated symbology.
    /// </summary>
    internal static DataFormatCheckResult? BuildLinearDataFormatCheck(
        string? decodedData,
        string? symbology)
    {
        if (string.IsNullOrWhiteSpace(decodedData)) return null;

        // Strip any whitespace; the decoded data must be all-numeric.
        string digits = decodedData.Trim();
        if (!digits.All(char.IsDigit)) return null;

        string sym = symbology?.Trim() ?? string.Empty;

        // UPC-E: compressed 6-digit format — check digit must be validated against the
        // expanded UPC-A equivalent, which requires a non-trivial expansion algorithm.
        // Mark N/A until expansion is implemented rather than silently reporting wrong results.
        if (sym.Equals("UPC-E", StringComparison.OrdinalIgnoreCase))
        {
            return new DataFormatCheckResult
            {
                Overall  = OverallPassFail.NotApplicable,
                Standard = "GS1 Linear",
                Rows     = [new() { Name = "GTIN", Data = digits, Check = "\u2014" }],
            };
        }

        int expectedLen = sym switch
        {
            "EAN-13" => 13,
            "UPC-A"  => 12,
            "EAN-8"  => 8,
            _        => 0,
        };

        if (expectedLen == 0 || digits.Length != expectedLen)
        {
            // Unknown symbology or unexpected length — emit a bare GTIN row with no check.
            return new DataFormatCheckResult
            {
                Overall  = OverallPassFail.NotApplicable,
                Standard = "GS1 Linear",
                Rows     = [new() { Name = "GTIN", Data = digits, Check = "\u2014" }],
            };
        }

        bool checkOk     = ValidateGs1CheckDigit(digits);
        string checkData = digits[^1].ToString();
        string checkPass = checkOk ? "PASS" : "FAIL";

        return new DataFormatCheckResult
        {
            Overall  = checkOk ? OverallPassFail.Pass : OverallPassFail.Fail,
            Standard = "GS1 Linear",
            Rows     =
            [
                new() { Name = "GTIN",      Data = digits[..^1], Check = checkPass },
                new() { Name = "Chk Digit", Data = checkData,    Check = checkPass },
            ],
        };
    }

    /// <summary>
    /// GS1 mod-10 check-digit validation (ISO/IEC 15420, Annex A).
    /// Weights alternate between 3 and 1 from right to left, excluding the
    /// check digit itself.  The check digit is valid when
    ///   (sum + check) mod 10 == 0.
    ///
    /// Works for any GS1 digit string: EAN-8 (8), UPC-A (12), EAN-13 (13).
    /// </summary>
    private static bool ValidateGs1CheckDigit(string digits)
    {
        if (digits.Length < 2) return false;

        int sum = 0;
        for (int i = 0; i < digits.Length - 1; i++)
        {
            if (!char.IsDigit(digits[i])) return false;
            int d = digits[i] - '0';
            // Weight is 3 for even positions from the right (0-based from check digit end),
            // 1 for odd.  Position from right of digit i = (digits.Length - 2 - i).
            int weight = (digits.Length - 2 - i) % 2 == 0 ? 3 : 1;
            sum += d * weight;
        }

        int expected = (10 - sum % 10) % 10;
        return char.IsDigit(digits[^1]) && (digits[^1] - '0') == expected;
    }
}
