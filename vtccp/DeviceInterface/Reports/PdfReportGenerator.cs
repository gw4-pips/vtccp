using System.Reflection;
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
///   2. RFID Validation table (Tag Detected / Lock Status / EPC Hex / EPC Tag URI / GCP Length / GTIN-14 / Serial / Result)
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

    // ── Brand lookup table ────────────────────────────────────────────────────
    // Maps known model substrings (case-insensitive) to all-caps brand names.
    // Order matters: more-specific patterns must appear before broader ones.
    // Extend this table when new verifier hardware is added.
    private static readonly (string Substring, string Brand)[] BrandPatterns =
    [
        // Cognex DataMan family (DataMan 100, 260, 370, 390, 395V, 470, 475V, …)
        ("DataMan",  "COGNEX"),
        // Axicon verifiers (Axicon 6000, 7000, 9000, etc.)
        ("Axicon",   "AXICON"),
        // Omron Microscan LVS verifiers (LVS-9510, LVS-7510, etc.)
        ("LVS",      "OMRON/LVS"),
        ("Omron",    "OMRON/LVS"),
        ("Microscan","OMRON/LVS"),
        // Webscan TruCheck verifiers
        ("Webscan",  "WEBSCAN"),
        ("TruCheck", "WEBSCAN"),
    ];

    /// <summary>
    /// Resolves the all-caps verifier brand name from an arbitrary string using the
    /// static <see cref="BrandPatterns"/> lookup table.
    /// Returns <see langword="null"/> when the input is empty or matches no known pattern.
    /// </summary>
    private static string? ResolveBrand(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;

        foreach (var (substring, brand) in BrandPatterns)
        {
            if (text.Contains(substring, StringComparison.OrdinalIgnoreCase))
                return brand;
        }

        return null;
    }

    /// <summary>
    /// Resolves the verifier brand for PDF labelling using four sources in priority order:
    /// <list type="number">
    ///   <item><see cref="VerificationRecord.VerifierBrand"/> — explicit adapter-supplied override.</item>
    ///   <item><see cref="VerificationRecord.DeviceModel"/> — SDK-connected DataMan devices always populate this.</item>
    ///   <item>Source PDF metadata — when <see cref="VerificationRecord.WebscanSourcePath"/> is a .pdf,
    ///         reads document info fields (Title, Creator, Author) via PdfSharp.</item>
    ///   <item>Raw file byte scan — scans the first 64 KB of the PDF file for brand keyword strings.</item>
    /// </list>
    /// Returns <see langword="null"/> when brand cannot be determined, so the caller can
    /// omit the brand prefix rather than incorrectly defaulting to COGNEX.
    /// </summary>
    private static string? ResolveEffectiveBrand(VerificationRecord r)
    {
        // 1. Explicit adapter-supplied override (file-export adapters: Webscan, Axicon, LVS)
        if (!string.IsNullOrWhiteSpace(r.VerifierBrand))
            return r.VerifierBrand;

        // 2. Device model string (SDK-connected DataMan devices always have this)
        string? fromModel = ResolveBrand(r.DeviceModel);
        if (fromModel != null)
            return fromModel;

        // 3 & 4. PDF source file — metadata fields then raw byte scan
        if (!string.IsNullOrWhiteSpace(r.WebscanSourcePath)
            && r.WebscanSourcePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)
            && File.Exists(r.WebscanSourcePath))
        {
            string? fromPdf = ExtractBrandFromPdf(r.WebscanSourcePath);
            if (fromPdf != null)
                return fromPdf;
        }

        return null;
    }

    /// <summary>
    /// Reads a PDF file and scans it for known verifier brand tokens.
    /// Primary path: reads document info fields (Title, Creator, Author, Subject) via
    /// PdfSharp — cheap and sufficient for any Webscan/Axicon/LVS report PDF.
    /// Secondary path: scans the first 64 KB of raw file bytes for brand keyword strings
    /// (covers PDFs where metadata is absent but brand text appears in the XMP stream or
    /// an uncompressed content fragment).
    /// Always safe — all exceptions are caught and logged; never throws.
    /// </summary>
    private static string? ExtractBrandFromPdf(string pdfPath)
    {
        // Primary: PdfSharp document info dictionary
        try
        {
            using var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
            string metaText = string.Concat(
                doc.Info.Title   ?? string.Empty, " ",
                doc.Info.Creator ?? string.Empty, " ",
                doc.Info.Author  ?? string.Empty, " ",
                doc.Info.Subject ?? string.Empty);
            string? fromMeta = ResolveBrand(metaText);
            if (fromMeta != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PDF] Brand '{fromMeta}' resolved from PDF metadata of '{Path.GetFileName(pdfPath)}'");
                return fromMeta;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[PDF] ExtractBrandFromPdf (metadata): {ex.GetType().Name}: {ex.Message}");
        }

        // Secondary: raw byte scan — covers XMP streams and uncompressed content fragments
        try
        {
            const int MaxScanBytes = 65536;
            using var fs = File.OpenRead(pdfPath);
            int readLen = (int)Math.Min(MaxScanBytes, fs.Length);
            byte[] buf = new byte[readLen];
            _ = fs.Read(buf, 0, readLen);
            // Latin-1 preserves every byte value; brand keywords are ASCII so encoding is irrelevant.
            string rawText = System.Text.Encoding.Latin1.GetString(buf);
            string? fromRaw = ResolveBrand(rawText);
            if (fromRaw != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PDF] Brand '{fromRaw}' resolved from raw byte scan of '{Path.GetFileName(pdfPath)}'");
            }
            return fromRaw;
        }
        catch (Exception ex2)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[PDF] ExtractBrandFromPdf (raw scan): {ex2.GetType().Name}: {ex2.Message}");
            return null;
        }
    }

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
                    col.Spacing(6);

                    col.Item().Element(c => BuildVerificationSummarySection(c, r));
                    col.Item().Element(c => BuildRfidTable(c, r));

                    // Barcode Image section — side-by-side when multi-mode
                    string? img2D     = r.RoiJpegImageBase64 ?? r.JpegImageBase64;
                    string? imgLinear = r.LinearJpegImageBase64;
                    bool multiMode    = !string.IsNullOrWhiteSpace(r.LinearSymbology);
                    if (multiMode && !string.IsNullOrWhiteSpace(img2D) && !string.IsNullOrWhiteSpace(imgLinear))
                        col.Item().Element(c => BuildDualImageSection(c, imgLinear, img2D, r.LinearSymbology!, r.Symbology));
                    else if (!string.IsNullOrWhiteSpace(img2D))
                        col.Item().Element(c => BuildImageSection(c, img2D));

                    // Data Format Check — with 2D + linear sub-sections when multi-mode
                    bool has2DdFc     = r.DataFormatCheck       is { Rows.Count: > 0 };
                    bool hasLinearDfc = r.LinearDataFormatCheck  is { Rows.Count: > 0 };
                    if (has2DdFc || hasLinearDfc)
                        col.Item().Element(c => BuildDataFormatSection(c, r.DataFormatCheck, r.LinearDataFormatCheck));
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
            // TODO Task #89: QuestPDF 2024.x has no BorderStyle API; dashed border requires a custom workaround.
            row.ConstantItem(90).Border(1).BorderColor("#999999")
               .AlignCenter().AlignMiddle().Padding(6).Column(col =>
               {
                   col.Item().AlignCenter().Text("VCCS")
                      .Bold().FontSize(12).LetterSpacing(1);
                   col.Item().AlignCenter().Text(txt =>
                   {
                       txt.Span("FlexWedge\u2122 Pro").Italic().FontSize(7).FontColor(GrayHex);
                   });
               });

            // Col 2: scan meta
            row.RelativeItem().PaddingHorizontal(8).AlignMiddle().Column(col =>
            {
                string dt     = r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt");
                // Priority: DeviceModel (SDK-connected DataMan) → VerifierBrand (file-export adapters)
                // → neutral placeholder.  Never fall back to "Cognex DataMan" for non-Cognex hardware.
                string device = r.DeviceModel ?? r.VerifierBrand ?? "\u2014";
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
                col.Item().AlignCenter().Column(titleCol =>
                {
                    titleCol.Item().AlignCenter().Text(txt =>
                    {
                        txt.Span("VCCS ").Bold().FontSize(11).FontColor(NavyHex);
                        txt.Span("FlexWedge\u2122 Pro").Bold().Italic().FontSize(11).FontColor(NavyHex);
                    });
                    titleCol.Item().AlignCenter().Text("RFID Validation Report")
                       .Bold().FontSize(11).FontColor(NavyHex);
                });
                col.Item().PaddingTop(4).AlignCenter()
                   .Element(c2 => RfidBadge(c2, r.RfidStatus));
            });

            // Col 4: company placeholder
            // TODO Task #89: dashed border pending QuestPDF workaround
            row.ConstantItem(90).Border(1).BorderColor("#999999")
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
        // Brand resolved from four sources in priority order (see ResolveEffectiveBrand):
        // VerifierBrand field → DeviceModel → source PDF metadata → raw PDF byte scan.
        // Returns null when brand is genuinely unknown — in that case we omit the brand
        // prefix rather than incorrectly labelling a Webscan/Axicon/LVS report as COGNEX.
        string? brand = ResolveEffectiveBrand(r);
        // Both this header and the "Barcode Verification Grades" sub-header use the
        // same subdued style (#2c5296, 8pt) to visually defer to the RFID section.
        // For Cognex DataMan hardware, insert "DataMan" between brand and "TruCheck"
        // since TruCheck on DataMan is a distinct product from the Webscan/Axicon TruCheck.
        string? dataManInfix = (brand == "COGNEX" &&
            r.DeviceModel?.Contains("DataMan", StringComparison.OrdinalIgnoreCase) == true)
            ? "DataMan " : null;
        string sectionTitle = brand != null
            ? $"{brand} {dataManInfix}TruCheck Barcode Verification Results Summary"
            : "TruCheck Barcode Verification Results Summary";

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
            // ── Section header — same subdued style as sub-header ─────────────
            // Intentionally smaller/lighter than the RFID section to show visual
            // hierarchy: VCCS RFID validation is the primary content.
            // Section title + italic "See separate report for details" annotation inline.
            col.Item()
               .Background("#2c5296").Padding(3)
               .Text(t =>
               {
                   t.Span(sectionTitle).Bold().FontSize(8).FontColor(Colors.White);
                   t.Span(" \u2014 ").Bold().FontSize(8).FontColor(Colors.White);
                   t.Span("See associated TruCheck report for additional details")
                    .Italic().FontSize(7.5f).FontColor(Colors.White);
               });

            // ── Summary table: 3-column (Symbology | Encoded Data | App Specification) ──
            // Symbol rows: 1 in single-mode, 2 in multi-mode (EAN/UPC first, 2D second).
            // DecodedData / LinearDecodedData exclude symbology-identifier prefixes
            // (e.g. ]d2) — those live in the DFC section.
            // A heavier bottom border on the last symbol row separates it from the
            // report-metadata rows (Report Name, Report Timestamp) that follow with
            // the value cell spanning the Encoded Data + App Spec columns.
            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.ConstantColumn(76);   // Symbology — snug for "GS1 DataMatrix" at 8pt
                    cols.RelativeColumn();      // Encoded Data — widest
                    cols.ConstantColumn(88);   // Application Spec. — snug for header text
                });

                // Column headers — same internal-border style as grades table header
                string[] sumHdrs = { "Symbology", "Encoded Data", "Application Spec." };
                for (int i = 0; i < sumHdrs.Length; i++)
                {
                    IContainer hc = table.Cell().Background(Colors.White)
                                         .BorderBottom(1).BorderColor("#999999");
                    if (i < sumHdrs.Length - 1) hc = hc.BorderRight(1).BorderColor("#999999");
                    hc.Padding(3).Text(sumHdrs[i]).Bold().FontSize(8);
                }

                // One symbol row. isSeparatorRow = true on the last symbol row;
                // produces a 1.5 pt bottom border to visually separate symbols from metadata.
                // BorderColor applies to all set borders on the cell, so both the bottom
                // and right-divider borders share the same color per row.
                void SymbolRow(string symb, string? encoded, string appSpecStr, bool isSeparatorRow)
                {
                    float bt = isSeparatorRow ? 1.5f : 1f;
                    string bc = isSeparatorRow ? "#888888" : "#aaaaaa";

                    table.Cell().BorderBottom(bt).BorderColor(bc).BorderRight(1)
                         .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                         .Text(symb).FontSize(8);
                    table.Cell().BorderBottom(bt).BorderColor(bc).BorderRight(1)
                         .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                         .Text(encoded ?? "\u2014").FontSize(9);

                    // App spec: for 2D symbols, appSpecStr is "Element\nString" or "Digital\nLink".
                    // Render as a horizontal Row: "GS1 —" vertically centred on the left,
                    // then the two-line stack (7pt, no gap) flush left beside it.
                    // For linear symbols the string has no \n — render single-line at 8pt.
                    var appCell = table.Cell().BorderBottom(bt).BorderColor(bc)
                                       .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4);
                    var appLines = appSpecStr.Split('\n');
                    if (appLines.Length > 1)
                        appCell.Row(r2 =>
                        {
                            r2.AutoItem().AlignMiddle()
                              .Text("GS1 \u2014").FontSize(8);
                            r2.ConstantItem(3); // gap between prefix and stack
                            r2.RelativeItem().AlignMiddle().Column(c2 =>
                            {
                                c2.Item().Text(appLines[0]).FontSize(7);
                                c2.Item().Text(appLines[1]).FontSize(7);
                            });
                        });
                    else
                        appCell.Text(appSpecStr).FontSize(8);
                }

                bool hasLinearSum = !string.IsNullOrWhiteSpace(r.LinearSymbology);
                string linAppResult = r.LinearOverallGrade?.PassFail switch
                {
                    OverallPassFail.Pass => "PASS",
                    OverallPassFail.Fail => "FAIL",
                    _                   => "\u2014",
                };

                // Derive Application Specification from symbology name:
                //   GS1 DataMatrix → "GS1 Element Strings"
                //   QR Code        → "GS1 Digital Link"
                //   EAN/UPC        → "GS1"
                //   Other          → ApplicationStandard field (appSpec)
                // AppSpecFor returns a \n-separated two-liner for 2D symbols so SymbolRow
                // can render line 1 at 8pt and line 2 at 7pt (smaller secondary text).
                static string AppSpecFor(string? symbology, string fallback) => symbology switch
                {
                    var s when s?.Contains("DataMatrix", StringComparison.OrdinalIgnoreCase) == true
                        => "Element\nString",
                    var s when s?.Contains("QR",         StringComparison.OrdinalIgnoreCase) == true
                        => "Digital\nLink",
                    _   => fallback,
                };

                // LinearAppSpecFor derives the GS1 GTIN qualifier for linear symbols.
                //   UPC-A / UPC-E  → GS1 — GTIN-12
                //   EAN-8          → GS1 — GTIN-8
                //   EAN-13         → GS1 — GTIN-13   (default EAN)
                //   GS1 DataBar    → GS1 — GTIN-14
                //   Other          → GS1
                static string LinearAppSpecFor(string? symbology) => symbology switch
                {
                    var s when s?.Contains("UPC",     StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-12",
                    var s when s?.Contains("EAN-8",   StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-8",
                    var s when s?.Contains("EAN",     StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-13",
                    var s when s?.Contains("DataBar", StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-14",
                    _   => "GS1",
                };

                if (hasLinearSum)
                {
                    // Row 1: linear (EAN/UPC) — not the separator row.
                    // App spec derived from symbology: UPC→GTIN-12, EAN-13→GTIN-13, etc.
                    SymbolRow(r.LinearSymbology!, r.LinearDecodedData,
                              LinearAppSpecFor(r.LinearSymbology), isSeparatorRow: false);
                    // Row 2: 2D symbol — separator row (heavier bottom border).
                    SymbolRow(r.Symbology ?? "\u2014", r.DecodedData,
                              AppSpecFor(r.Symbology, appSpec), isSeparatorRow: true);
                }
                else
                {
                    // Single-symbol mode — one row, separator row (heavier bottom border)
                    SymbolRow(r.Symbology ?? "\u2014", r.DecodedData,
                              AppSpecFor(r.Symbology, appSpec), isSeparatorRow: true);
                }

                // Metadata rows: label (col 1) + value spanning cols 2–3 (ColumnSpan 2).
                // isFirstMeta = true adds a 2pt navy top border (same weight as exterior)
                // to create the heavier separator between symbol rows and metadata rows.
                void MetaRow(string label, string? value, bool isLast, bool isFirstMeta = false)
                {
                    IContainer lc = table.Cell().BorderRight(1).BorderColor("#aaaaaa");
                    if (isFirstMeta) lc = lc.BorderTop(1).BorderColor(NavyHex);
                    if (!isLast) lc = lc.BorderBottom(1).BorderColor("#aaaaaa");
                    lc.PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                      .Text(label).FontSize(8.5f).FontColor(GrayHex);

                    IContainer vc = table.Cell().ColumnSpan(2);
                    if (isFirstMeta) vc = vc.BorderTop(1).BorderColor(NavyHex);
                    if (!isLast)
                        vc.BorderBottom(1).BorderColor("#aaaaaa")
                          .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                          .Text(value ?? "\u2014").FontSize(9);
                    else
                        vc.PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                          .Text(value ?? "\u2014").FontSize(9);
                }

                MetaRow("Report Name",
                        reportName,
                        isLast: false, isFirstMeta: true);
                MetaRow("Report Date/Time",
                        r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt"),
                        isLast: true);
            });

            // ── Sub-header: Barcode Verification Grades ───────────────────────
            // PaddingTop matches the outer column gap so all inter-section spacing is uniform.
            col.Item().PaddingTop(6)
               .Background("#2c5296").Padding(3)
               .Text("Barcode Verification Grades").Bold().FontSize(8).FontColor(Colors.White);

            // ── 7-column grades table — matching Webscan TruCheck style ──────
            // In multi-mode the linear (EAN/UPC) symbol occupies row 1; 2D symbol row 2.
            // In single-mode there is only one row (whichever symbol was graded).
            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.RelativeColumn(1.5f);  // Symbology
                    cols.RelativeColumn(2.0f);  // Standard
                    cols.RelativeColumn(1.5f);  // Grade
                    cols.RelativeColumn(1.0f);  // Aperture
                    cols.RelativeColumn(1.0f);  // Wavelength
                    cols.RelativeColumn(2.0f);  // Lighting
                    cols.RelativeColumn(2.5f);  // Formal Grade
                });

                // Column headers
                var hdrs = new[] { "Symbology", "Standard", "Grade", "Aperture", "Wavelength", "Lighting", "Formal Grade" };
                for (int i = 0; i < hdrs.Length; i++)
                {
                    IContainer hCell = table.Cell().Background(Colors.White)
                                            .BorderBottom(1).BorderColor("#999999");
                    if (i < hdrs.Length - 1)
                        hCell = hCell.BorderRight(1).BorderColor("#999999");
                    hCell.Padding(3).AlignCenter().Text(hdrs[i]).Bold().FontSize(8);
                }

                // Helper: emit one data row with internal right-border column dividers.
                // The last row in the table has no bottom border (outer navy closes it).
                void GradeRow(TableDescriptor tbl, string[] vals, bool addBottomBorder)
                {
                    for (int i = 0; i < vals.Length; i++)
                    {
                        IContainer dc = table.Cell();
                        if (addBottomBorder)
                            dc = dc.BorderBottom(1).BorderColor("#aaaaaa");
                        if (i < vals.Length - 1)
                            dc = dc.BorderRight(1).BorderColor("#aaaaaa");
                        dc.PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                          .AlignCenter().Text(vals[i]).FontSize(8);
                    }
                }

                bool hasLinear = !string.IsNullOrWhiteSpace(r.LinearSymbology);

                if (hasLinear)
                {
                    // Row 1 — linear (EAN/UPC) symbol
                    string linGrade = r.LinearOverallGrade is { } lg
                        ? $"{lg.LetterGradeString} ({lg.NumericGrade?.ToString("F1") ?? "\u2014"})"
                        : "\u2014";
                    string linAperture   = r.LinearAperture.HasValue   ? r.LinearAperture.Value.ToString("D2") : "\u2014";
                    string linWavelength = r.LinearWavelength.HasValue  ? r.LinearWavelength.Value.ToString()  : "\u2014";
                    GradeRow(table,
                        new[] { r.LinearSymbology!, r.LinearStandard ?? "ISO/IEC 15416",
                                linGrade, linAperture, linWavelength,
                                r.LinearLighting ?? "\u2014", r.LinearFormalGrade ?? "\u2014" },
                        addBottomBorder: true);

                    // Row 2 — 2D symbol (no bottom border — outer navy closes)
                    GradeRow(table,
                        new[] { r.Symbology ?? "\u2014", gradeStandard,
                                gradeGrade, gradeAperture, gradeWavelength,
                                gradeLighting, gradeFormal },
                        addBottomBorder: false);
                }
                else
                {
                    // Single-symbol mode — one data row, no bottom border
                    GradeRow(table,
                        new[] { r.Symbology ?? "\u2014", gradeStandard,
                                gradeGrade, gradeAperture, gradeWavelength,
                                gradeLighting, gradeFormal },
                        addBottomBorder: false);
                }
            });
        });
    }

    // ── RFID Validation table ─────────────────────────────────────────────────

    private static void BuildRfidTable(IContainer c, VerificationRecord r)
    {
        c.Column(col =>
        {
            // Section header — full navy, prominent: this is the primary VCCS section.
            // "EPC" qualifier included when GS1 encodation (>99% of cases).
            // TODO: define adjectives for non-GS1 schemes (DOD-96, ISO 17367, etc.)
            // "EPC" = GS1 encodation (SGTIN, SSCC, etc.) — >99% of cases.
            // "UHF" = non-GS1 scheme (DOD-96, ISO 17367, custom, etc.).
            bool isGS1Rfid = r.ApplicationStandard?
                .StartsWith("GS1", StringComparison.OrdinalIgnoreCase) == true;
            string rfidAdj   = isGS1Rfid ? "EPC " : "UHF ";
            col.Item().Background(NavyHex).Padding(3).Text(txt =>
            {
                txt.Span("VCCS ").Bold().FontSize(10).FontColor(Colors.White);
                txt.Span("FlexWedge\u2122 Pro ").Bold().Italic().FontSize(10).FontColor(Colors.White);
                txt.Span($"{rfidAdj}RFID Validation Summary").Bold().FontSize(10).FontColor(Colors.White);
            });

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
                     .Padding(3).Text(left).Bold().FontSize(8);
                    t.Cell().Background("#e8eef5").BorderBottom(1).BorderColor("#aaaaaa")
                     .Padding(3).Text(right).Bold().FontSize(8);
                }

                void DataRow(string label, string? value, bool highlight = false)
                {
                    string bg = highlight ? "#f0f7ff" : Colors.White;
                    table.Cell().Background(bg).BorderBottom(1).BorderColor("#aaaaaa")
                         .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                         .Text(label).FontSize(9);
                    table.Cell().Background(bg).BorderBottom(1).BorderColor("#aaaaaa")
                         .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                         .Text(value ?? "\u2014").FontSize(9);
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
                DataRow("Tag Detected",    tagLabel);

                // Tag lock status (queried separately from inventory — often null)
                // Values derive from ASR-P35U CheckTagStatus: 40=PermaLock, 41=Lock, 42=Unlock
                string lockDisplay = r.RfidTagLockStatus switch
                {
                    "Locked"       => "Locked",
                    "PermaLocked"  => "Permanently Locked",
                    "Unlocked"     => "Unlocked",
                    "Unknown"      => "Unknown",
                    null           => "\u2014",
                    var other      => other,
                };
                // Append all possible lock-status values as a parenthetical legend for auditors.
                const string LockOpts = " (Permalocked / Locked / Unlocked)";
                DataRow("Tag Lock Status", lockDisplay + LockOpts);

                DataRow("EPC Hex",         r.RfidEpcHex,      tagDetected);

                // EPC Tag URI — urn:epc:tag:... form as used by RFID middleware.
                // Distinct from the GS1 Digital Link URI that appears in QR code payloads.
                DataRow("EPC Tag URI",     r.RfidEpcTagUri,   tagDetected);

                // GCP Length: "Valid (N)" or "Invalid (N)" where N = GCP digit count.
                // Shown above GTIN-14 per v6 layout.
                // When GcpTableDate is available, the value cell uses RichText to append
                // an italic provenance annotation: "— From GCP prefix table as of yyyy-MM-dd".
                string gcpLenPart = r.RfidGcpLength.HasValue
                    ? $" ({r.RfidGcpLength.Value})"
                    : string.Empty;
                string gcpDisplay = r.RfidGcpValid switch
                {
                    true  => $"Valid{gcpLenPart}",
                    false => $"Invalid{gcpLenPart}",
                    null  => "\u2014",
                };

                // Emit GCP Length row — inline (bypasses DataRow helper) to allow RichText.
                table.Cell().Background(Colors.White).BorderBottom(1).BorderColor("#aaaaaa")
                     .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                     .Text("GCP Length").FontSize(9);
                if (!string.IsNullOrWhiteSpace(r.RfidGcpTableDate))
                {
                    table.Cell().Background(Colors.White).BorderBottom(1).BorderColor("#aaaaaa")
                         .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                         .Text(txt =>
                         {
                             txt.Span(gcpDisplay).FontSize(9);
                             txt.Span($" \u2014 From GCP prefix table as of {r.RfidGcpTableDate}")
                                .Italic().FontSize(7.5f).FontColor(GrayHex);
                         });
                }
                else
                {
                    table.Cell().Background(Colors.White).BorderBottom(1).BorderColor("#aaaaaa")
                         .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                         .Text(gcpDisplay).FontSize(9);
                }

                DataRow("GTIN-14",         r.RfidGtin14,      tagDetected);
                DataRow("Serial",          r.RfidSerial,      tagDetected);

                // Result row with colour.
                // When only a linear (EAN/UPC) symbol is present the RFID cross-validation
                // can only compare GTIN (the barcode carries no serial); the result wording
                // reflects this.  Multi-mode records (LinearSymbology set) match against the
                // 2D symbol, so they use the same wording as a 2D-only single-mode scan.
                bool eanOnly = r.Is1D && string.IsNullOrWhiteSpace(r.LinearSymbology);
                string resultVal = r.RfidStatus switch
                {
                    "Pass" when eanOnly => "Pass \u2014 GTIN match only (serial not in barcode)",
                    "Fail" when eanOnly => "Fail \u2014 GTIN mismatch",
                    var s               => s ?? "\u2014",
                };
                if (!string.IsNullOrWhiteSpace(r.RfidMismatchDetail) && !eanOnly)
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

                table.Cell().Background(resultBg)
                     .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                     .Text("Result").Bold().FontSize(9).FontColor(resultFg);
                table.Cell().Background(resultBg)
                     .PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4)
                     .Text(resultVal).Bold().FontSize(9).FontColor(resultFg);
            });

        });
    }

    // ── Barcode image ─────────────────────────────────────────────────────────

    /// <summary>Single-symbol image section (single-mode).</summary>
    private static void BuildImageSection(IContainer c, string base64Jpeg)
    {
        try
        {
            byte[] imgBytes = Convert.FromBase64String(base64Jpeg);
            c.Column(col =>
            {
                col.Item().Background("#2c5296").Padding(3)
                   .Text("Barcode Image").Bold().FontSize(8).FontColor(Colors.White);
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

    /// <summary>
    /// Two-symbol image section (multi-mode) — linear crop on the left, 2D crop on the right.
    /// Each image is capped to ~40 pt wide so both fit comfortably on one Letter page.
    /// Falls back gracefully when either image fails to decode.
    /// </summary>
    private static void BuildDualImageSection(
        IContainer c,
        string     linearBase64,
        string     twoDBase64,
        string     linearLabel,
        string?    twoDLabel)
    {
        c.Column(col =>
        {
            col.Item().Background("#2c5296").Padding(3)
               .Text("Barcode Images").Bold().FontSize(8).FontColor(Colors.White);

            col.Item().Border(1).BorderColor(NavyHex).Padding(4).Row(row =>
            {
                // Left: linear symbol
                row.RelativeItem().Column(imgCol =>
                {
                    imgCol.Item().AlignCenter().Text(linearLabel)
                          .FontSize(7).FontColor(GrayHex).Italic();
                    try
                    {
                        byte[] lb = Convert.FromBase64String(linearBase64);
                        imgCol.Item().AlignCenter()
                              .MaxWidth(40 * 3)   // ~3 in; linear barcodes are wide
                              .MaxHeight(40 * 72f / 96f * 72f)  // ~1 in height cap
                              .Image(lb).FitArea();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[PDF] BuildDualImageSection linear decode failed: {ex.Message}");
                        imgCol.Item().AlignCenter().Text("[image unavailable]")
                              .FontSize(8).FontColor(GrayHex);
                    }
                });

                // Divider
                row.ConstantItem(1).Background("#cccccc");

                // Right: 2D symbol
                row.RelativeItem().Column(imgCol =>
                {
                    imgCol.Item().AlignCenter().Text(twoDLabel ?? "2D Symbol")
                          .FontSize(7).FontColor(GrayHex).Italic();
                    try
                    {
                        byte[] tb = Convert.FromBase64String(twoDBase64);
                        imgCol.Item().AlignCenter()
                              .MaxWidth(40 * 72f / 96f * 72f)  // ~1 in square crop
                              .MaxHeight(40 * 72f / 96f * 72f)
                              .Image(tb).FitArea();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[PDF] BuildDualImageSection 2D decode failed: {ex.Message}");
                        imgCol.Item().AlignCenter().Text("[image unavailable]")
                              .FontSize(8).FontColor(GrayHex);
                    }
                });
            });
        });
    }

    // ── Data Format Check summary ─────────────────────────────────────────────

    /// <summary>
    /// Renders the Data Format Check section.
    /// When <paramref name="linearDfc"/> is also supplied (multi-mode), the table shows
    /// a "2D Symbol" sub-header above the 2D rows and a "Linear Symbol (EAN/UPC)" sub-header
    /// above the EAN rows — all within one bordered table.
    /// The overall PASS/FAIL pill reflects <paramref name="mainDfc"/> (the 2D symbol's result)
    /// when both are present; falls back to <paramref name="linearDfc"/> when mainDfc is null.
    /// </summary>
    private static void BuildDataFormatSection(
        IContainer             c,
        DataFormatCheckResult? mainDfc,
        DataFormatCheckResult? linearDfc = null)
    {
        // Guard: nothing to render
        if ((mainDfc is null || mainDfc.Rows.Count == 0) &&
            (linearDfc is null || linearDfc.Rows.Count == 0))
            return;

        bool multiMode = mainDfc is { Rows.Count: > 0 } && linearDfc is { Rows.Count: > 0 };

        // Title uses the 2D standard when available; generic fallback otherwise.
        string standard = mainDfc?.Standard ?? linearDfc?.Standard ?? string.Empty;
        string title = string.IsNullOrWhiteSpace(standard)
            ? "Data Format Check"
            : $"Data Format Check \u2014 {standard}";

        c.Column(col =>
        {
            col.Item().Background("#2c5296").Padding(3)
               .Text(t =>
               {
                   t.Span(title).Bold().FontSize(8).FontColor(Colors.White);
                   t.Span(" \u2014 ").Bold().FontSize(8).FontColor(Colors.White);
                   t.Span("See associated TruCheck report for additional details")
                    .Italic().FontSize(7.5f).FontColor(Colors.White);
               });

            col.Item().Border(1).BorderColor(NavyHex).Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.RelativeColumn(2);   // Name
                    cols.RelativeColumn(3);   // Data
                    cols.ConstantColumn(50);  // Check
                });

                // Column headers
                foreach (string h in new[] { "Field", "Data", "Check" })
                {
                    table.Cell().Background("#e8eef5").BorderBottom(1).BorderColor("#aaaaaa")
                         .Padding(3).Text(h).Bold().FontSize(8);
                }

                // Emit a full-width sub-header row spanning all 3 columns.
                void SubHeader(string label)
                {
                    table.Cell().ColumnSpan(3)
                         .Background("#dce6f1").BorderBottom(1).BorderColor("#aaaaaa")
                         .PaddingTop(2).PaddingBottom(2).PaddingLeft(4)
                         .Text(label).Bold().FontSize(7.5f).FontColor(NavyHex);
                }

                // Emit a single DFC data row (3 cells).
                void DfcRow(DataFormatCheckRow row, bool isLast)
                {
                    bool pass      = string.Equals(row.Check, "PASS", StringComparison.OrdinalIgnoreCase);
                    string rowBg   = pass ? Colors.White : FailBack;
                    string checkFg = pass ? PassHex : FailHex;

                    // Returns a styled IContainer ready for a Text() call.
                    Func<string, IContainer> cell = bg =>
                    {
                        IContainer c = table.Cell().Background(bg);
                        if (!isLast) c = c.BorderBottom(1).BorderColor("#aaaaaa");
                        return c.PaddingTop(2.5f).PaddingBottom(2).PaddingLeft(4).PaddingRight(4);
                    };

                    cell(rowBg).Text(row.Name).FontSize(8);
                    cell(rowBg).Text(row.Data).FontSize(8);
                    cell(rowBg).Text(row.Check).Bold().FontSize(8).FontColor(checkFg);
                }

                if (multiMode)
                {
                    // 2D symbol block
                    SubHeader("2D Symbol");
                    var rows2D = mainDfc!.Rows;
                    for (int i = 0; i < rows2D.Count; i++)
                        DfcRow(rows2D[i], isLast: false);

                    // Linear symbol block (last row truly last in table)
                    SubHeader("Linear Symbol (EAN/UPC)");
                    var rowsLin = linearDfc!.Rows;
                    for (int i = 0; i < rowsLin.Count; i++)
                        DfcRow(rowsLin[i], isLast: i == rowsLin.Count - 1);
                }
                else
                {
                    // Single-symbol: whichever DFC is non-empty
                    var activeDfc  = (mainDfc?.Rows.Count ?? 0) > 0 ? mainDfc! : linearDfc!;
                    var activeRows = activeDfc.Rows;
                    for (int i = 0; i < activeRows.Count; i++)
                        DfcRow(activeRows[i], isLast: i == activeRows.Count - 1);
                }
            });

            // Overall result pill — prefer 2D result in multi-mode
            DataFormatCheckResult? pillSource = (mainDfc?.Rows.Count ?? 0) > 0 ? mainDfc : linearDfc;
            string overallText = pillSource?.Overall switch
            {
                OverallPassFail.Pass          => "OVERALL: PASS",
                OverallPassFail.Fail          => "OVERALL: FAIL",
                _                             => string.Empty,
            };
            if (!string.IsNullOrEmpty(overallText))
            {
                string bg = pillSource?.Overall switch
                {
                    OverallPassFail.Pass    => PassBack,
                    OverallPassFail.Fail    => FailBack,
                    _                      => Colors.White,
                };
                string fg = pillSource?.Overall switch
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
                txt.Span("VCCS ").Bold().FontSize(7);
                txt.Span("FlexWedge\u2122 Pro").Bold().Italic().FontSize(7);
                // Version read from the entry assembly at runtime so it always matches the installed build.
                string ver = Assembly.GetEntryAssembly()?.GetName().Version?.ToString(3) ?? string.Empty;
                string verSuffix = ver.Length > 0 ? $" v{ver}" : string.Empty;
                txt.Span($" RFID Validation Report{verSuffix}").Bold().FontSize(7);
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
