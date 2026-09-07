using DeviceInterface.Reports;
using ExcelEngine.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using System.Text;
using Xunit;

namespace DeviceInterface.Tests;

public sealed class Gs1CanonicalReportTests
{
    [Fact]
    public void Validation_lists_all_mandatory_missing_fields()
    {
        Gs1ReportValidationResult result = Gs1LinearHtmlReportGenerator.Validate(new Gs1ReportData());

        Assert.False(result.IsValid);
        Assert.Contains("Print method", result.MissingFields);
        Assert.Contains("GS1 symbol specification table", result.MissingFields);
    }

    [Fact]
    public void Linear_generator_uses_canonical_labels_and_makes_missing_visible()
    {
        string html = Gs1LinearHtmlReportGenerator.Generate(new Gs1ReportData());

        Assert.Contains("Testing summary of the linear barcode", html);
        Assert.Contains("ISO/IEC 15416 print quality grade", html);
        Assert.Contains("[MISSING: Print method]", html);
        Assert.Contains("Release 26.0, Ratified, Jan 26", html);
    }

    [Fact]
    public void Two_dimensional_generator_keeps_the_canonical_pages_free_of_rfid_content()
    {
        string html = Gs1TwoDimensionalHtmlReportGenerator.Generate(Complete());

        Assert.Contains("Testing summary of the two-dimensional barcode", html);
        Assert.Contains("Overall ISO/IEC 15415 print quality grade", html);
        Assert.DoesNotContain("RFID supplemental information", html);
        Assert.DoesNotContain("EPC Tag URI", html);
        Assert.Contains("@page { size:A4", html);
    }

    [Theory]
    [InlineData(true, "Pass", true)]
    [InlineData(true, "NoTag", true)]
    [InlineData(true, "Skipped", false)]
    [InlineData(false, "Pass", false)]
    [InlineData(null, null, false)]
    public void VeriWedge_addendum_requires_real_rfid_evidence(
        bool? connected, string? status, bool expected)
    {
        var record = new ExcelEngine.Models.VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidReaderConnected = connected,
            RfidStatus = status
        };

        Assert.Equal(expected, VccsPdfRenderer.HasVeriWedgeEvidence(record));
    }

    [Fact]
    public void Canonical_html_is_identical_with_and_without_rfid_evidence()
    {
        var barcodeOnly = new VerificationRecord
        {
            VerificationDateTime = new DateTime(2026, 9, 7),
            Symbology = "GS1 DataMatrix",
            ProductName = "Product",
            DeviceModel = "Verifier",
            HtmlDecodedData = "(01)09506000134352",
            HtmlOverallGradeDisplay = "4.0 (A)",
            HtmlStandard = "ISO/IEC 15415",
        };
        VerificationRecord withRfid = barcodeOnly with
        {
            RfidReaderConnected = true,
            RfidStatus = "Pass",
            RfidEpcTagUri = "urn:epc:tag:sgtin-96:1.0612345.012345.1",
            RfidReaderManufacturer = "AsReader",
            RfidReaderModel = "ASR-P35U",
        };

        string withoutRfid = Gs1CanonicalReport.GenerateTwoDimensionalHtml(
            Gs1CanonicalReport.FromVerificationRecord(barcodeOnly));
        string withRfidHtml = Gs1CanonicalReport.GenerateTwoDimensionalHtml(
            Gs1CanonicalReport.FromVerificationRecord(withRfid));

        Assert.Equal(withoutRfid, withRfidHtml);
        Assert.DoesNotContain("RFID", withRfidHtml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pdf_merge_preserves_canonical_pages_before_veriwedge_pages()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        string canonicalPath = root + "-canonical.pdf";
        string veriwedgePath = root + "-veriwedge.pdf";
        string mergedPath = root + "-merged.pdf";
        try
        {
            CreateMarkerPdf(canonicalPath, 501, 502);
            CreateMarkerPdf(veriwedgePath, 601, 602);

            VccsPdfRenderer.MergePdfDocuments(canonicalPath, veriwedgePath, mergedPath);

            using PdfDocument merged = PdfReader.Open(mergedPath, PdfDocumentOpenMode.Import);
            Assert.Equal(4, merged.PageCount);
            Assert.Equal(501, merged.Pages[0].Width.Point, 3);
            Assert.Equal(502, merged.Pages[1].Width.Point, 3);
            Assert.Equal(601, merged.Pages[2].Width.Point, 3);
            Assert.Equal(602, merged.Pages[3].Width.Point, 3);
        }
        finally
        {
            File.Delete(canonicalPath);
            File.Delete(veriwedgePath);
            File.Delete(mergedPath);
        }
    }

    [Fact]
    public void Letter_profile_is_available_without_changing_the_a4_default()
    {
        string html = Gs1LinearHtmlReportGenerator.Generate(Complete(), Gs1PrintProfile.Letter);

        Assert.Contains("@page { size:Letter", html);
    }

    [Theory]
    [InlineData(Gs1PrintProfile.A4, 595.28, 841.89)]
    [InlineData(Gs1PrintProfile.Letter, 612.0, 792.0)]
    public async Task Windows_renderer_preserves_canonical_pages_before_complete_veriwedge_addendum(
        Gs1PrintProfile profile,
        double expectedWidth,
        double expectedHeight)
    {
        if (!OperatingSystem.IsWindows())
            return; // WebView2 and the bundled wkhtmltopdf fallback are Windows-only.

        string root = Path.Combine(Path.GetTempPath(), $"gs1-render-regression-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        string canonicalOnlyPath = Path.Combine(root, $"{profile}-canonical-only.pdf");
        string veriwedgeOnlyPath = Path.Combine(root, $"{profile}-veriwedge-only.pdf");
        string mergedPath = Path.Combine(root, $"{profile}-canonical-with-veriwedge.pdf");
        try
        {
            VerificationRecord withoutRfid = CanonicalRecord();
            VerificationRecord withRfid = withoutRfid with
            {
                RfidReaderConnected = true,
                RfidStatus = "Pass",
                RfidEpcHex = "3034257BF7194E4000000001",
                RfidEpcTagUri = "urn:epc:tag:sgtin-96:1.0612345.012345.1",
                RfidTid = "E28068940000502F3D5A1C2B",
                RfidReaderManufacturer = "AsReader",
                RfidReaderModel = "ASR-P35U",
            };
            string canonicalHtml = Gs1CanonicalReport.GenerateTwoDimensionalHtml(
                Complete(), profile);

            await VccsPdfRenderer.RenderGs1WithVeriWedgeAddendumAsync(
                canonicalHtml, withoutRfid, canonicalOnlyPath, profile);
            await VccsPdfRenderer.RenderAsync(
                VccsHtmlReportGenerator.Generate(withRfid),
                veriwedgeOnlyPath,
                printProfile: profile,
                addPageNumbers: false);
            await VccsPdfRenderer.RenderGs1WithVeriWedgeAddendumAsync(
                canonicalHtml, withRfid, mergedPath, profile);

            using PdfDocument canonicalOnly =
                PdfReader.Open(canonicalOnlyPath, PdfDocumentOpenMode.Import);
            using PdfDocument veriwedgeOnly =
                PdfReader.Open(veriwedgeOnlyPath, PdfDocumentOpenMode.Import);
            using PdfDocument merged =
                PdfReader.Open(mergedPath, PdfDocumentOpenMode.Import);

            Assert.Equal(2, canonicalOnly.PageCount);
            Assert.True(veriwedgeOnly.PageCount > 0);
            Assert.Equal(
                canonicalOnly.PageCount + veriwedgeOnly.PageCount,
                merged.PageCount);

            for (int index = 0; index < canonicalOnly.PageCount; index++)
            {
                Assert.Equal(
                    FirstContentStream(canonicalOnly.Pages[index]),
                    FirstContentStream(merged.Pages[index]));
            }

            for (int index = 0; index < veriwedgeOnly.PageCount; index++)
            {
                Assert.Equal(
                    FirstContentStream(veriwedgeOnly.Pages[index]),
                    FirstContentStream(merged.Pages[canonicalOnly.PageCount + index]));
            }

            AssertPagesArePrintable(canonicalOnly, expectedWidth, expectedHeight);
            AssertPagesArePrintable(veriwedgeOnly, expectedWidth, expectedHeight);
            AssertPagesArePrintable(merged, expectedWidth, expectedHeight);
            AssertPageNumbering(merged);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Gs1_compliance_and_technical_values_are_only_the_literal_plugin_evidence()
    {
        var gs1 = new Dictionary<string, Gs1ParameterAssessment>
        {
            ["Data structure¹ (syntax)"] = new()
            {
                NormativeRequirement = "Required",
                ComplianceResult = "Axicon: NOT COMPLIANT",
                LiteralAssessedValue = "]C10109506000134352",
                CommentReference = "GS1 Table 5-1",
                Provenance = "Axicon export"
            }
        };
        var iso = new Dictionary<string, Gs1IsoAssessment>
        {
            ["Rmin (minimum reflectance)"] = new() { LiteralValue = "6.0%", CommentReference = "ISO/IEC 15416" },
            ["Print growth (+/- %) – Process control parameter"] = new() { LiteralValue = "-2.15%" }
        };
        string html = Gs1LinearHtmlReportGenerator.Generate(Complete(gs1Parameters: gs1, isoParameters: iso));

        Assert.Contains("Axicon: NOT COMPLIANT", html);
        Assert.Contains("]C10109506000134352", html);
        Assert.Contains("6.0%", html);
        Assert.Contains("-2.15%", html);
        Assert.DoesNotContain("<td>Within standard range</td>", html);
        Assert.DoesNotContain("<td>Compliant to standard</td>", html);
    }

    [Fact]
    public void Linear_validation_requires_the_core_linear_evidence()
    {
        Gs1ReportValidationResult result = Gs1LinearHtmlReportGenerator.Validate(Complete());

        Assert.False(result.IsValid);
        Assert.Contains("ISO/IEC evidence: Rmin (minimum reflectance)", result.MissingFields);
        Assert.Contains("GS1 compliance evidence: Data structure¹ (syntax)", result.MissingFields);
        Assert.Contains("GS1 compliance evidence: Validity of GS1 Company Prefix", result.MissingFields);
        Assert.Contains("GS1 compliance evidence: Human readable", result.MissingFields);
    }

    [Fact]
    public void Two_dimensional_validation_requires_core_2d_evidence_but_not_optional_subcomponents()
    {
        Gs1ReportValidationResult result = Gs1TwoDimensionalHtmlReportGenerator.Validate(Complete());

        Assert.False(result.IsValid);
        Assert.Contains("ISO/IEC evidence: Axial nonuniformity", result.MissingFields);
        Assert.Contains("ISO/IEC evidence: Grid nonuniformity", result.MissingFields);
        Assert.Contains("ISO/IEC evidence: Unused Error Correction", result.MissingFields);
        Assert.Contains("ISO/IEC evidence: Fixed pattern damage", result.MissingFields);
        Assert.DoesNotContain("ISO/IEC evidence: • Format information⁴", result.MissingFields);
    }

    private static Gs1ReportData Complete(
        IReadOnlyDictionary<string, Gs1ParameterAssessment>? gs1Parameters = null,
        IReadOnlyDictionary<string, Gs1IsoAssessment>? isoParameters = null) => new()
    {
        IssueDate = "1 January 2026", TestingAgencyName = "Agency",
        ProductDescription = "Product", BarcodeType = "GS1 DataMatrix",
        PrintMethod = "Thermal", NumberOfBarcodesOnProduct = "1",
        VerifierDevice = "Verifier", VerificationSoftwareVersion = "1.0",
        LastVerifierCalibrationDate = "1 January 2026", SymbolSpecificationTable = "Table 5-1",
        TestedEnvironments = "Retail POS", PlacementResult = "Complies",
        TwoDimensionalProximity = "Yes", OverallIsoIecGrade = "4.0 (A)",
        DecodedText = "(01)09506000134352",
        Gs1Parameters = gs1Parameters ?? new Dictionary<string, Gs1ParameterAssessment>(),
        IsoParameters = isoParameters ?? new Dictionary<string, Gs1IsoAssessment>()
    };

    private static VerificationRecord CanonicalRecord() => new()
    {
        VerificationDateTime = new DateTime(2026, 9, 7, 10, 30, 0),
        Symbology = "GS1 DataMatrix",
        ProductName = "Windows PDF regression fixture",
        DeviceName = "DM475V",
        DeviceModel = "DM475V",
        DeviceSerial = "FIXTURE-001",
        FirmwareVersion = "6.1.16",
        ApplicationStandard = "GS1",
        HtmlDecodedData = "(01)09506000134352",
        HtmlOverallGradeDisplay = "4.0 (A)",
        HtmlStandard = "ISO/IEC 15415",
        HtmlSourceFileName = "canonical-fixture.html",
        HtmlVerifiedString = "9/7/2026 10:30:00 AM",
        HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
        RfidReaderConnected = false,
        RfidStatus = "Skipped",
    };

    private static byte[] FirstContentStream(PdfPage page)
    {
        Assert.True(page.Contents.Elements.Count > 0, "Page has no PDF content streams.");
        PdfDictionary content = page.Contents.Elements.GetDictionary(0)
            ?? throw new InvalidDataException("Page content stream is not a PDF dictionary.");
        Assert.NotNull(content.Stream);
        Assert.NotEmpty(content.Stream!.Value);
        return content.Stream.Value;
    }

    private static void AssertPagesArePrintable(
        PdfDocument document,
        double expectedWidth,
        double expectedHeight)
    {
        for (int index = 0; index < document.PageCount; index++)
        {
            PdfPage page = document.Pages[index];
            Assert.Equal(expectedWidth, page.Width.Point, 1);
            Assert.Equal(expectedHeight, page.Height.Point, 1);

            int contentBytes = 0;
            for (int streamIndex = 0; streamIndex < page.Contents.Elements.Count; streamIndex++)
            {
                PdfDictionary content = page.Contents.Elements.GetDictionary(streamIndex)
                    ?? throw new InvalidDataException(
                        $"Page {index + 1} content stream is not a PDF dictionary.");
                contentBytes += content.Stream?.Value.Length ?? 0;
            }
            Assert.True(contentBytes > 100, $"Page {index + 1} is blank or unexpectedly empty.");
        }
    }

    private static void AssertPageNumbering(PdfDocument document)
    {
        for (int index = 0; index < document.PageCount; index++)
        {
            var content = new StringBuilder();
            for (int streamIndex = 0; streamIndex < document.Pages[index].Contents.Elements.Count; streamIndex++)
            {
                PdfDictionary stream =
                    document.Pages[index].Contents.Elements.GetDictionary(streamIndex)
                    ?? throw new InvalidDataException(
                        $"Page {index + 1} content stream is not a PDF dictionary.");
                content.Append(Encoding.ASCII.GetString(stream.Stream?.Value ?? []));
            }

            Assert.Contains(
                $"Page {index + 1} of {document.PageCount}",
                content.ToString(),
                StringComparison.Ordinal);
        }
    }

    private static void CreateMarkerPdf(string path, params int[] widths)
    {
        using var document = new PdfDocument();
        foreach (int width in widths)
        {
            PdfPage page = document.AddPage();
            page.Width = PdfSharp.Drawing.XUnit.FromPoint(width);
            page.Height = PdfSharp.Drawing.XUnit.FromPoint(700);
        }
        document.Save(path);
    }
}