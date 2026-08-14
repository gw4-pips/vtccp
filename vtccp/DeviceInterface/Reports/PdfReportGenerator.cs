using ExcelEngine.Models;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DeviceInterface.Reports;

/// <summary>
/// Generates a single-page VCCS RFID Validation PDF per accepted scan.
///
/// The Webscan PDF is left completely untouched.  When
/// <see cref="VerificationRecord.WebscanSourcePath"/> points to an existing Webscan PDF,
/// an optional merge step appends the VCCS page to produce one combined document.
///
/// Layout (top to bottom):
///   1. 4-column header: VCCS logo | scan meta | RFID result badge | company
///   2. RFID Validation table (Tag Detected / EPC Hex / GTIN-14 / GCP Valid / Serial / Result)
///   3. Optional barcode image (ROI JPEG preferred; falls back to L1 crop)
///   4. Data Format Check summary (GS1 AI rows when present)
///   5. VCCS footer
///
/// Thread safety: all public methods are stateless; instances are not required.
/// Call <see cref="GenerateAsync"/> as fire-and-forget from the UI thread —
/// failures are caught and logged via Debug output, never surfaced to the operator.
/// </summary>
public static class PdfReportGenerator
{
    // ── Colour palette (matches Webscan TruCheck navy + GS1 traffic-light grades) ─

    private static readonly string NavyHex  = "#1a3a6b";
    private static readonly string PassHex  = "#155724";
    private static readonly string PassBack = "#d4edda";
    private static readonly string FailHex  = "#721c24";
    private static readonly string FailBack = "#f8d7da";
    private static readonly string WarnBack = "#fff3cd";
    private static readonly string WarnHex  = "#856404";
    private static readonly string GrayHex  = "#6c757d";

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Generates the VCCS PDF page and writes it to <paramref name="outputDir"/>.
    /// When <paramref name="webscanPdfPath"/> is not null and the file exists, the
    /// VCCS page is also appended to the Webscan PDF (producing a combined file
    /// alongside the Webscan original — the original is never modified).
    /// </summary>
    public static async Task GenerateAsync(
        VerificationRecord record,
        string             outputDir,
        string?            webscanPdfPath,
        CancellationToken  ct = default)
    {
        if (string.IsNullOrWhiteSpace(outputDir)) return;

        try
        {
            QuestPDF.Settings.License = LicenseType.Community;

            Directory.CreateDirectory(outputDir);
            string ts       = record.VerificationDateTime.ToString("yyyy-MM-dd_HH-mm-ss");
            string vccsPath = Path.Combine(outputDir, $"{ts}_vccs_rfid.pdf");

            // Generate the single VCCS page
            byte[] vccsBytes = await Task.Run(() => BuildPdfBytes(record), ct);
            await File.WriteAllBytesAsync(vccsPath, vccsBytes, ct);

            System.Diagnostics.Debug.WriteLine(
                $"[PDF] VCCS report written: {vccsPath}");

            // Optional merge — append VCCS page to the Webscan PDF
            if (!string.IsNullOrWhiteSpace(webscanPdfPath) && File.Exists(webscanPdfPath))
            {
                string mergedPath = Path.ChangeExtension(webscanPdfPath,
                    $"_vccs{Path.GetExtension(webscanPdfPath)}");
                await Task.Run(() => MergePdfs(webscanPdfPath, vccsBytes, mergedPath), ct);
                System.Diagnostics.Debug.WriteLine(
                    $"[PDF] Merged report (Webscan + VCCS): {mergedPath}");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[PDF] GenerateAsync failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    // ── PDF construction ──────────────────────────────────────────────────────

    private static byte[] BuildPdfBytes(VerificationRecord r)
    {
        var doc = Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.Letter);
                page.Margin(0.5f, Unit.Inch);
                page.DefaultTextStyle(x => x.FontSize(9).FontFamily(Fonts.Arial));

                page.Header().Element(c => BuildHeader(c, r));

                page.Content().Column(col =>
                {
                    col.Spacing(8);

                    col.Item().Element(c => BuildVerificationSummarySection(c, r));
                    col.Item().Element(c => BuildRfidTable(c, r));

                    string? imgB64 = r.RoiJpegImageBase64 ?? r.JpegImageBase64;
                    if (!string.IsNullOrWhiteSpace(imgB64))
                        col.Item().Element(c => BuildImageSection(c, imgB64));

                    if (r.DataFormatCheck is { Rows.Count: > 0 } dfc)
                        col.Item().Element(c => BuildDataFormatSection(c, dfc));
                });

                page.Footer().Element(c => BuildFooter(c, r));
            });
        });

        return doc.GeneratePdf();
    }

    // ── Header — 4 columns ────────────────────────────────────────────────────

    private static void BuildHeader(IContainer c, VerificationRecord r)
    {
        c.Background("#edf1f7").BorderBottom(2).BorderColor(NavyHex).PaddingBottom(6).PaddingTop(4).Row(row =>
        {
            // Col 1: VCCS logo placeholder
            row.ConstantItem(90).Border(1).BorderColor("#999999").BorderStyle(BorderStyle.Dashed)
               .AlignCenter().AlignMiddle().Padding(6).Column(col =>
               {
                   col.Item().AlignCenter().Text("VCCS")
                      .Bold().FontSize(12).LetterSpacing(1);
                   col.Item().AlignCenter().Text("FlexWedge\u2122")
                      .FontSize(7).FontColor(GrayHex);
               });

            // Col 2: scan meta
            row.RelativeItem().PaddingHorizontal(8).AlignMiddle().Column(col =>
            {
                string dt     = r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt");
                string device = r.DeviceModel ?? "Cognex DataMan";
                string serial = r.DeviceSerial ?? "\u2014";
                string fw     = r.FirmwareVersion ?? "\u2014";

                col.Item().Text(dt).FontSize(8).FontColor(GrayHex);
                col.Item().Text($"Device: {device}").Bold().FontSize(9);
                col.Item().Text($"Serial: {serial}").FontSize(8);
                col.Item().Text($"Firmware: {fw}").FontSize(8);
            });

            // Col 3: RFID result badge (centred)
            row.RelativeItem().AlignCenter().AlignMiddle().Column(col =>
            {
                col.Item().AlignCenter().Text("VCCS RFID Validation Report")
                   .Bold().FontSize(11).FontColor(NavyHex);
                col.Item().PaddingTop(4).AlignCenter()
                   .Element(c2 => RfidBadge(c2, r.RfidStatus));
            });

            // Col 4: company placeholder
            row.ConstantItem(90).Border(1).BorderColor("#999999").BorderStyle(BorderStyle.Dashed)
               .AlignCenter().AlignMiddle().Padding(6)
               .Text(r.CompanyName ?? "Company Logo").FontSize(7).FontColor(GrayHex);
        });
    }

    private static void RfidBadge(IContainer c, string? status)
    {
        string label;
        string bgColor;
        string fgColor;

        switch (status)
        {
            case "Pass":
                label   = "\u2713 RFID MATCHED";
                bgColor = PassBack;
                fgColor = PassHex;
                break;
            case "Fail":
                label   = "\u2717 RFID MISMATCH";
                bgColor = FailBack;
                fgColor = FailHex;
                break;
            case "NoTag":
                label   = "\u26a0 NO RFID TAG";
                bgColor = WarnBack;
                fgColor = WarnHex;
                break;
            case "MultipleTagsDetected":
                label   = "\u26a0 MULTIPLE TAGS";
                bgColor = WarnBack;
                fgColor = WarnHex;
                break;
            case "Skipped":
                label   = "\u2014 RFID SKIPPED";
                bgColor = "#f8f9fa";
                fgColor = GrayHex;
                break;
            default:
                label   = string.IsNullOrWhiteSpace(status) ? "\u2014 NO RFID DATA" : status.ToUpperInvariant();
                bgColor = "#f8f9fa";
                fgColor = GrayHex;
                break;
        }

        c.Background(bgColor).Border(1).BorderColor(fgColor)
         .PaddingHorizontal(8).PaddingVertical(3)
         .Text(label).Bold().FontSize(9).FontColor(fgColor);
    }

    // ── Verification Results Summary + Barcode Grades ────────────────────────

    private static void BuildVerificationSummarySection(IContainer c, VerificationRecord r)
    {
        // Title varies by device family
        string verifier = (r.DeviceModel ?? string.Empty)
            .Contains("DataMan", StringComparison.OrdinalIgnoreCase)
            ? "DataMan TruCheck"
            : "Webscan TruCheck";
        string sectionTitle = $"{verifier} Verification Results Summary";

        // Application Specification row  (standard name + PASS/FAIL)
        string appSpec = !string.IsNullOrWhiteSpace(r.ApplicationStandard)
            ? r.ApplicationStandard
            : r.Standard ?? "\u2014";
        string appResult = r.OverallGrade?.PassFail switch
        {
            OverallPassFail.Pass => "PASS",
            OverallPassFail.Fail => "FAIL",
            _                   => "\u2014",
        };

        // Report name: Webscan source file name if known, else timestamped default
        string reportName = !string.IsNullOrWhiteSpace(r.WebscanSourcePath)
            ? Path.GetFileName(r.WebscanSourcePath)
            : $"{r.VerificationDateTime:yyyy-MM-dd_HH-mm-ss}_vccs_rfid.pdf";

        // Grades row values
        string gradeStandard   = r.Standard ?? "\u2014";
        string gradeGrade      = r.OverallGrade is { } og
            ? $"{og.LetterGradeString} ({og.NumericGrade?.ToString("F1") ?? "\u2014"})"
            : "\u2014";
        string gradeAperture   = r.Aperture.HasValue   ? r.Aperture.Value.ToString("D2")  : "\u2014";
        string gradeWavelength = r.Wavelength.HasValue ? r.Wavelength.Value.ToString()     : "\u2014";
        string gradeLighting   = r.Lighting   ?? "\u2014";
        string gradeFormal     = r.FormalGrade ?? "\u2014";

        c.Column(col =>
        {
            // ── Main section header ───────────────────────────────────────────
            col.Item()
               .Background(NavyHex).Padding(4)
               .Text(sectionTitle).Bold().FontSize(10).FontColor(Colors.White);

            // ── Summary rows (2-col label | value) ───────────────────────────
            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.ConstantColumn(160);
                    cols.RelativeColumn();
                });

                void Row(string label, string? value)
                {
                    table.Cell().BorderBottom(1).BorderColor("#dddddd").Padding(4)
                         .Text(label).FontSize(9);
                    table.Cell().BorderBottom(1).BorderColor("#dddddd").Padding(4)
                         .Text(value ?? "\u2014").FontSize(9);
                }

                Row("Symbology",              r.Symbology);
                Row("Encoded Data",           r.DecodedData ?? "NO DECODE");
                Row("Application Specification", $"{appSpec} \u2014 {appResult}");
                Row("Report Name",            reportName);
                Row("Report Timestamp",       r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt"));
            });

            // ── Sub-header: Barcode Verification Grades (smaller banner) ─────
            col.Item().PaddingTop(6)
               .Background("#2c5296").Padding(3)
               .Text("Barcode Verification Grades").Bold().FontSize(8).FontColor(Colors.White);

            // ── 6-column grades table — matching Webscan TruCheck style ──────
            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.RelativeColumn(2.0f);  // Standard
                    cols.RelativeColumn(1.5f);  // Grade
                    cols.RelativeColumn(1.0f);  // Aperture
                    cols.RelativeColumn(1.0f);  // Wavelength
                    cols.RelativeColumn(2.0f);  // Lighting
                    cols.RelativeColumn(2.5f);  // Formal Grade
                });

                // Bold bordered column headers (Webscan style)
                foreach (string h in new[] { "Standard", "Grade", "Aperture", "Wavelength", "Lighting", "Formal Grade" })
                {
                    table.Cell().Border(1).BorderColor("#999999")
                         .Padding(4).AlignCenter()
                         .Text(h).Bold().FontSize(8);
                }

                // Single data row
                foreach (string v in new[] { gradeStandard, gradeGrade, gradeAperture, gradeWavelength, gradeLighting, gradeFormal })
                {
                    table.Cell().Border(1).BorderColor("#dddddd")
                         .Padding(4).AlignCenter()
                         .Text(v).FontSize(8);
                }
            });
        });
    }

    // ── RFID Validation table ─────────────────────────────────────────────────

    private static void BuildRfidTable(IContainer c, VerificationRecord r)
    {
        c.Column(col =>
        {
            // Section header
            col.Item().Background(NavyHex).Padding(4)
               .Text("RFID Validation").Bold().FontSize(10).FontColor(Colors.White);

            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.ConstantColumn(140);   // label
                    cols.RelativeColumn();       // value
                });

                static void HeaderRow(TableDescriptor t, string left, string right)
                {
                    t.Cell().Background("#e8eef5").BorderBottom(1).BorderColor("#aaaaaa")
                     .Padding(4).Text(left).Bold().FontSize(8);
                    t.Cell().Background("#e8eef5").BorderBottom(1).BorderColor("#aaaaaa")
                     .Padding(4).Text(right).Bold().FontSize(8);
                }

                void DataRow(string label, string? value, bool highlight = false)
                {
                    string bg = highlight ? "#f0f7ff" : Colors.White;
                    table.Cell().Background(bg).BorderBottom(1).BorderColor("#dddddd")
                         .Padding(4).Text(label).FontSize(9);
                    table.Cell().Background(bg).BorderBottom(1).BorderColor("#dddddd")
                         .Padding(4).Text(value ?? "\u2014").FontSize(9);
                }

                // Header row
                HeaderRow(table, "Field", "Value");

                // Tag Detected
                bool tagDetected = r.RfidStatus is "Pass" or "Fail" or "MultipleTagsDetected";
                string tagLabel = r.RfidStatus switch
                {
                    "Pass" or "Fail"          => "Yes",
                    "NoTag"                   => "No",
                    "MultipleTagsDetected"    => "Yes (multiple)",
                    "Skipped"                 => "Skipped",
                    _                         => r.RfidStatus ?? "\u2014",
                };
                DataRow("Tag Detected",  tagLabel);
                DataRow("EPC Hex",       r.RfidEpcHex,   tagDetected);
                DataRow("GTIN-14",       r.RfidGtin14,   tagDetected);
                DataRow("GCP Valid",     r.RfidGcpValid switch
                {
                    true  => "Yes \u2713",
                    false => "No \u2717",
                    null  => "\u2014",
                });
                DataRow("Serial",        r.RfidSerial,   tagDetected);

                // Result row with colour
                string resultVal = r.RfidStatus ?? "\u2014";
                if (!string.IsNullOrWhiteSpace(r.RfidMismatchDetail))
                    resultVal += $" \u2014 {r.RfidMismatchDetail}";

                string resultBg = r.RfidStatus switch
                {
                    "Pass" => PassBack,
                    "Fail" => FailBack,
                    "NoTag" or "MultipleTagsDetected" => WarnBack,
                    _ => Colors.White,
                };
                string resultFg = r.RfidStatus switch
                {
                    "Pass" => PassHex,
                    "Fail" => FailHex,
                    "NoTag" or "MultipleTagsDetected" => WarnHex,
                    _ => Colors.Black,
                };

                table.Cell().Background(resultBg).Padding(4)
                     .Text("Result").Bold().FontSize(9).FontColor(resultFg);
                table.Cell().Background(resultBg).Padding(4)
                     .Text(resultVal).Bold().FontSize(9).FontColor(resultFg);
            });

            // Scan window duration (secondary metadata)
            if (r.RfidScanWindowMs.HasValue)
            {
                col.Item().PaddingTop(2).AlignRight()
                   .Text($"Scan window: {r.RfidScanWindowMs.Value} ms")
                   .FontSize(7).FontColor(GrayHex).Italic();
            }
        });
    }

    // ── Barcode image ─────────────────────────────────────────────────────────

    private static void BuildImageSection(IContainer c, string base64Jpeg)
    {
        try
        {
            byte[] imgBytes = Convert.FromBase64String(base64Jpeg);
            c.Column(col =>
            {
                col.Item().Background(NavyHex).Padding(4)
                   .Text("Barcode Image").Bold().FontSize(10).FontColor(Colors.White);
                col.Item().Border(1).BorderColor(NavyHex).AlignCenter()
                   .Padding(4).MaxHeight(2.5f * 72)  // 2.5 inches at 72pt/in
                   .Image(imgBytes).FitArea();
            });
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[PDF] BuildImageSection: image decode failed: {ex.Message}");
        }
    }

    // ── Data Format Check summary ─────────────────────────────────────────────

    private static void BuildDataFormatSection(IContainer c, DataFormatCheckResult dfc)
    {
        c.Column(col =>
        {
            string title = string.IsNullOrWhiteSpace(dfc.Standard)
                ? "Data Format Check"
                : $"Data Format Check \u2014 {dfc.Standard}";

            col.Item().Background(NavyHex).Padding(4)
               .Text(title).Bold().FontSize(10).FontColor(Colors.White);

            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.RelativeColumn(2);   // Name
                    cols.RelativeColumn(3);   // Data
                    cols.ConstantColumn(50);  // Check
                });

                // Header
                foreach (string h in new[] { "Field", "Data", "Check" })
                {
                    table.Cell().Background("#e8eef5").BorderBottom(1).BorderColor("#aaaaaa")
                         .Padding(4).Text(h).Bold().FontSize(8);
                }

                // Rows
                foreach (var row in dfc.Rows)
                {
                    bool pass = string.Equals(row.Check, "PASS", StringComparison.OrdinalIgnoreCase);
                    string rowBg = pass ? Colors.White : FailBack;
                    string checkFg = pass ? PassHex : FailHex;

                    table.Cell().Background(rowBg).BorderBottom(1).BorderColor("#dddddd")
                         .Padding(3).Text(row.Name).FontSize(8);
                    table.Cell().Background(rowBg).BorderBottom(1).BorderColor("#dddddd")
                         .Padding(3).Text(row.Data).FontSize(8);
                    table.Cell().Background(rowBg).BorderBottom(1).BorderColor("#dddddd")
                         .Padding(3).Text(row.Check).Bold().FontSize(8).FontColor(checkFg);
                }
            });

            // Overall result pill
            string overallText = dfc.Overall switch
            {
                OverallPassFail.Pass           => "OVERALL: PASS",
                OverallPassFail.Fail           => "OVERALL: FAIL",
                OverallPassFail.NotApplicable  => string.Empty,
                _                              => string.Empty,
            };
            if (!string.IsNullOrEmpty(overallText))
            {
                string bg = dfc.Overall switch
                {
                    OverallPassFail.Pass    => PassBack,
                    OverallPassFail.Fail    => FailBack,
                    _                      => Colors.White,
                };
                string fg = dfc.Overall switch
                {
                    OverallPassFail.Pass    => PassHex,
                    OverallPassFail.Fail    => FailHex,
                    _                      => GrayHex,
                };
                col.Item().PaddingTop(2).AlignRight()
                   .Background(bg).Border(1).BorderColor(fg).PaddingHorizontal(8).PaddingVertical(3)
                   .Text(overallText).Bold().FontSize(8).FontColor(fg);
            }
        });
    }

    // ── Footer ────────────────────────────────────────────────────────────────

    private static void BuildFooter(IContainer c, VerificationRecord r)
    {
        c.BorderTop(1).BorderColor("#cccccc").PaddingTop(4).Row(row =>
        {
            row.RelativeItem().Text(txt =>
            {
                txt.Span("VCCS FlexWedge\u2122 RFID Validation Report").Bold().FontSize(7);
                txt.Span($"  \u2014  Generated {r.VerificationDateTime:yyyy-MM-dd HH:mm:ss}").FontSize(7).FontColor(GrayHex);
            });
            row.RelativeItem().AlignRight().Text(txt =>
            {
                txt.Span("Verification Command \u0026 Control System (VCCS)").FontSize(7).FontColor(GrayHex);
            });
        });
    }

    // ── PdfSharp merge ────────────────────────────────────────────────────────

    /// <summary>
    /// Appends the VCCS PDF page(s) to the Webscan PDF and writes the combined
    /// document to <paramref name="mergedPath"/>.  The Webscan source file is never
    /// modified.  Failures are logged and swallowed — the VCCS-only PDF already exists.
    /// </summary>
    private static void MergePdfs(string webscanPath, byte[] vccsBytes, string mergedPath)
    {
        try
        {
            // PdfSharp requires a password parameter when opening for import.
            using var webscanDoc = PdfReader.Open(webscanPath, PdfDocumentOpenMode.Import);
            using var vccsDoc    = PdfReader.Open(
                new MemoryStream(vccsBytes), PdfDocumentOpenMode.Import);

            using var output = new PdfDocument();

            // Copy all Webscan pages first
            for (int i = 0; i < webscanDoc.PageCount; i++)
                output.AddPage(webscanDoc.Pages[i]);

            // Append all VCCS pages
            for (int i = 0; i < vccsDoc.PageCount; i++)
                output.AddPage(vccsDoc.Pages[i]);

            output.Save(mergedPath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[PDF] MergePdfs failed: {ex.GetType().Name}: {ex.Message}");
        }
    }
}
