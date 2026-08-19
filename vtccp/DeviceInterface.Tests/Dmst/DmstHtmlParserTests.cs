namespace DeviceInterface.Tests.Dmst;

using DeviceInterface.Dmst;
using ExcelEngine.Models;
using Xunit;

/// <summary>
/// Fixture-driven tests for the Webscan TruCheck HTML multi-mode parser path and
/// for the GS1 check-digit validator used in LinearDataFormatCheck construction.
///
/// Fixtures are synthetic minimal HTML documents whose cell sequences match
/// the expected Webscan TruCheck multi-mode layout:
///   • Header table — up to two "D.D (L)" grade summaries (one per symbol).
///   • Linear (EAN/UPC) characteristics section — symbology cell, digit data, formal grade.
///   • 2D (DataMatrix) characteristics section — matrix size, error correction fields.
///
/// All timestamps use "2026-01-15_10-30-45" prefix so filename-based DateTime
/// parsing succeeds and ParseSucceeded = true.
/// </summary>
public sealed class DmstHtmlParserTests
{
    // ── Fixture source path (valid timestamp prefix required for ParseSucceeded) ──

    private const string FixturePath = @"C:\fake\2026-01-15_10-30-45-000_fixture.html";
    private const string DataPrefixedFixturePath =
        @"C:\CodeQuality\_F1_01006961147042882172803282009_2026-08-18_20-04-21-314.html";

    // ── Multi-mode fixture: EAN-13 (grade A, 4.0) + DataMatrix (grade B, 3.5) ──
    //
    // Cell sequence (flat, in document order):
    //   [0..3]  header metadata
    //   [4]     "4.0 (A)"  — linear overall grade (grade display 1)
    //   [5]     "3.5 (B)"  — 2D overall grade    (grade display 2)
    //   [6]     "EAN-13"   — linear symbology marker  (linearSymbIdx = 6)
    //   [7]     "5901234123457" — decoded data (13 digits, within 10 of idx 6)
    //   [8]     "A/06/660/Diffuse" — linear formal grade (within 80 of idx 6)
    //   [9..N]  2D characteristics (Matrix Size, ECLevel, DataMaskPattern, …)
    //
    // 5901234123457 has valid EAN-13 check digit 7.
    private const string MultiModeEan13Html = """
        <html><body>
        <table>
          <tr><td>Report</td><td>TruCheck Multi-Mode</td><td>Standard</td><td>ISO</td></tr>
          <tr><td>Symbol 1: EAN-13</td><td>4.0 (A)</td></tr>
          <tr><td>Symbol 2: Data Matrix</td><td>3.5 (B)</td></tr>
        </table>
        <table>
          <tr><td>EAN-13</td><td>5901234123457</td></tr>
          <tr><td>A/06/660/Diffuse</td></tr>
          <tr><td>Matrix Size</td><td>22x22 (Data: 20x20)</td></tr>
          <tr><td>Nominal X Dim</td><td>12.6 mil</td></tr>
          <tr><td>Horizontal BWG</td><td>-3%</td></tr>
          <tr><td>Vertical BWG</td><td>-4%</td></tr>
          <tr><td>Error Correction Level</td><td>M</td></tr>
          <tr><td>Data Mask Pattern</td><td>2</td></tr>
          <tr><td>Image</td><td>Black on white</td></tr>
          <tr><td>ECI</td><td>000003</td></tr>
          <tr><td>Data Codewords</td><td>44</td></tr>
          <tr><td>Error Correction Budget</td><td>26</td></tr>
          <tr><td>Encoded characters</td><td>36</td></tr>
          <tr><td>Errors Corrected</td><td>0</td></tr>
          <tr><td>Error Capacity Used</td><td>0</td></tr>
          <tr><td>Total Codewords</td><td>70</td></tr>
        </table>
        </body></html>
        """;

    // ── Single-mode EAN-13 fixture — only one grade pattern ──────────────────────
    //
    // Single-mode linear scan: isMultiMode must be false even though an EAN-13
    // cell is present, because there is only one "D.D (L)" grade display.
    // The one grade maps to OverallGrade (not LinearOverallGrade).
    private const string SingleModeEan13Html = """
        <html><body>
        <table>
          <tr><td>Report</td><td>TruCheck Single</td></tr>
          <tr><td>Overall Grade</td><td>4.0 (A)</td></tr>
        </table>
        <table>
          <tr><td>EAN-13</td><td>5901234123457</td></tr>
          <tr><td>A/06/660/Diffuse</td></tr>
        </table>
        </body></html>
        """;

    // ── Single-mode 2D fixture — no linear symbology cell ───────────────────────
    //
    // Standard DataMatrix scan: no EAN/UPC cell, one grade → not multi-mode.
    private const string SingleMode2DHtml = """
        <html><body>
        <table>
          <tr><td>Report</td><td>TruCheck Single</td></tr>
          <tr><td>Overall Grade</td><td>4.0 (A)</td></tr>
        </table>
        <table>
          <tr><td>Matrix Size</td><td>22x22 (Data: 20x20)</td></tr>
          <tr><td>Nominal X Dim</td><td>12.6 mil</td></tr>
          <tr><td>Error Correction Level</td><td>M</td></tr>
          <tr><td>Data Mask Pattern</td><td>2</td></tr>
          <tr><td>Image</td><td>Black on white</td></tr>
          <tr><td>ECI</td><td>000003</td></tr>
          <tr><td>Data Codewords</td><td>44</td></tr>
          <tr><td>Error Correction Budget</td><td>26</td></tr>
          <tr><td>Encoded characters</td><td>36</td></tr>
          <tr><td>Total Codewords</td><td>70</td></tr>
        </table>
        </body></html>
        """;

    // ══ ParseHtml — multi-mode detection ════════════════════════════════════════

    [Fact]
    public void ParseHtml_MultiMode_IsMultiMode_True()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.True(report.IsMultiMode);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearSymbology_IsEan13()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("EAN-13", report.LinearSymbology);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearDecodedData_IsDigitString()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("5901234123457", report.LinearDecodedData);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearOverallGrade_IsLetterA()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("A", report.LinearOverallGrade);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearOverallGradeNumeric_Is4Point0()
    {
        // Numeric grade must be parsed from the "4.0 (A)" display — not inferred
        // from the letter. This is the core precision requirement.
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal(4.0m, report.LinearOverallGradeNumeric);
    }

    [Fact]
    public void ParseHtml_MultiMode_OverallGrade_Is2DGrade_B()
    {
        // The 2D symbol's grade ("3.5 (B)") must map to the primary OverallGrade
        // field — NOT to LinearOverallGrade — because it is the second pattern.
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("B", report.OverallGrade);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearFormalGrade_Parsed()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("A/06/660/Diffuse", report.LinearFormalGrade);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearAperture_Is6()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal(6, report.LinearAperture);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearWavelength_Is660()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal(660, report.LinearWavelength);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearLighting_IsDiffuse()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("Diffuse", report.LinearLighting);
    }

    [Fact]
    public void ParseHtml_MultiMode_LinearStandard_IsIso15416()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("ISO/IEC 15416", report.LinearStandard);
    }

    // ── 2D supplemental fields still extracted in multi-mode ───────────────────

    [Fact]
    public void ParseHtml_MultiMode_EcLevel_StillExtracted()
    {
        // The 2D ECLevel field must not be lost due to multi-mode parsing.
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("M", report.ECLevel);
    }

    [Fact]
    public void ParseHtml_MultiMode_DataMaskPattern_StillExtracted()
    {
        var report = DmstHtmlScraper.ParseHtml(MultiModeEan13Html, FixturePath);

        Assert.Equal("2", report.DataMaskPattern);
    }

    // ══ ParseHtml — single-mode linear must NOT be misclassified as multi-mode ══

    [Fact]
    public void ParseHtml_SingleModeLinear_IsMultiMode_False()
    {
        // One grade display present, linear symbology cell present → NOT multi-mode.
        var report = DmstHtmlScraper.ParseHtml(SingleModeEan13Html, FixturePath);

        Assert.False(report.IsMultiMode);
    }

    [Fact]
    public void ParseHtml_SingleModeLinear_LinearSymbology_IsNull()
    {
        var report = DmstHtmlScraper.ParseHtml(SingleModeEan13Html, FixturePath);

        Assert.Null(report.LinearSymbology);
    }

    [Fact]
    public void ParseHtml_SingleModeLinear_OverallGrade_IsA()
    {
        // The one grade must map to OverallGrade, not be lost or misrouted.
        var report = DmstHtmlScraper.ParseHtml(SingleModeEan13Html, FixturePath);

        Assert.Equal("A", report.OverallGrade);
    }

    [Fact]
    public void ParseHtml_SingleModeLinear_LinearOverallGrade_IsNull()
    {
        var report = DmstHtmlScraper.ParseHtml(SingleModeEan13Html, FixturePath);

        Assert.Null(report.LinearOverallGrade);
    }

    // ══ ParseHtml — single-mode 2D ══════════════════════════════════════════════

    [Fact]
    public void ParseHtml_SingleMode2D_IsMultiMode_False()
    {
        var report = DmstHtmlScraper.ParseHtml(SingleMode2DHtml, FixturePath);

        Assert.False(report.IsMultiMode);
    }

    [Fact]
    public void ParseHtml_SingleMode2D_OverallGrade_IsA()
    {
        var report = DmstHtmlScraper.ParseHtml(SingleMode2DHtml, FixturePath);

        Assert.Equal("A", report.OverallGrade);
    }

    // ══ BuildLinearDataFormatCheck ═══════════════════════════════════════════════

    [Fact]
    public void BuildLinearDfc_ValidEan13_ReturnsPass()
    {
        // 5901234123457 — valid EAN-13 check digit 7
        var result = DmstReportValidator.BuildLinearDataFormatCheck(
            "5901234123457", "EAN-13");

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.Pass, result.Overall);
        Assert.Equal(2, result.Rows.Count);
        Assert.Equal("PASS", result.Rows[0].Check);  // GTIN row
        Assert.Equal("PASS", result.Rows[1].Check);  // Chk Digit row
        Assert.Equal("7", result.Rows[1].Data);       // check digit value
    }

    [Fact]
    public void BuildLinearDfc_InvalidEan13CheckDigit_ReturnsFail()
    {
        // 5901234123458 — wrong check digit (should be 7)
        var result = DmstReportValidator.BuildLinearDataFormatCheck(
            "5901234123458", "EAN-13");

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.Fail, result.Overall);
        Assert.Equal("FAIL", result.Rows[0].Check);
        Assert.Equal("FAIL", result.Rows[1].Check);
    }

    [Fact]
    public void BuildLinearDfc_ValidUpcA_ReturnsPass()
    {
        // 012345678905 — valid UPC-A check digit 5
        var result = DmstReportValidator.BuildLinearDataFormatCheck(
            "012345678905", "UPC-A");

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.Pass, result.Overall);
        Assert.Equal(2, result.Rows.Count);
        Assert.Equal("PASS", result.Rows[1].Check);
    }

    [Fact]
    public void BuildLinearDfc_ValidEan8_ReturnsPass()
    {
        // 40170725 — valid EAN-8 check digit 5
        var result = DmstReportValidator.BuildLinearDataFormatCheck(
            "40170725", "EAN-8");

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.Pass, result.Overall);
        Assert.Equal("5", result.Rows[1].Data);
    }

    [Fact]
    public void BuildLinearDfc_UpcE_ReturnsNotApplicable()
    {
        // UPC-E check digit requires UPC-A expansion — must not run the standard
        // EAN/UPC check and risk a false FAIL. Must return N/A.
        var result = DmstReportValidator.BuildLinearDataFormatCheck(
            "01234565", "UPC-E");

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.NotApplicable, result.Overall);
        Assert.Single(result.Rows);
        Assert.Equal("\u2014", result.Rows[0].Check);
    }

    [Fact]
    public void BuildLinearDfc_NullData_ReturnsNull()
    {
        var result = DmstReportValidator.BuildLinearDataFormatCheck(null, "EAN-13");

        Assert.Null(result);
    }

    [Fact]
    public void BuildLinearDfc_EmptyData_ReturnsNull()
    {
        var result = DmstReportValidator.BuildLinearDataFormatCheck("   ", "EAN-13");

        Assert.Null(result);
    }

    [Fact]
    public void BuildLinearDfc_WrongLength_ReturnsNotApplicable()
    {
        // 12 digits provided for EAN-13 (expects 13) — unknown-length fallback.
        var result = DmstReportValidator.BuildLinearDataFormatCheck(
            "590123412345", "EAN-13");

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.NotApplicable, result.Overall);
    }

    [Fact]
    public void BuildDataFormatCheck_LiteralF1ElementString_ProducesGtinAndSerialRows()
    {
        var record = new VerificationRecord
        {
            Symbology  = "GS1 DataMatrix",
            DecodedData = "<F1>01006961147042882172803282009",
        };

        var result = DmstReportValidator.BuildDataFormatCheck(record);

        Assert.NotNull(result);
        Assert.Equal(OverallPassFail.Pass, result.Overall);
        Assert.Contains(result.Rows, row =>
            row.Name == "AI (01) GTIN-14" && row.Data == "0069611470428");
        Assert.Contains(result.Rows, row =>
            row.Name == "AI (21) Serial" && row.Data == "72803282009");
    }

    [Fact]
    public void ParseHtml_DataPrefixedDmstFilename_ExtractsTimestampAndRetainsFilename()
    {
        var report = DmstHtmlScraper.ParseHtml(
            "<html><body><p>Verified: Mon 18-Aug-2026 08:04:21 PM</p></body></html>",
            DataPrefixedFixturePath);

        Assert.True(report.ParseSucceeded);
        Assert.Equal(new DateTime(2026, 8, 18, 20, 4, 21), report.ScanDateTime);
        Assert.Equal(
            "_F1_01006961147042882172803282009_2026-08-18_20-04-21-314.html",
            report.HtmlSourceFileName);
    }

    [Fact]
    public void ParseHtml_TimestamplessFilename_WithVerifiedHeader_IsAccepted()
    {
        var report = DmstHtmlScraper.ParseHtml(
            "<html><body><p>Verified: Mon 18-Aug-2026 08:04:21 PM</p></body></html>",
            @"C:\CodeQuality\barcode-report.html");

        Assert.True(report.ParseSucceeded);
        Assert.Null(report.ScanDateTime);
        Assert.Equal("Mon 18-Aug-2026 08:04:21 PM", report.HtmlVerifiedString);
    }

    [Fact]
    public void ConfiguredReportDirectory_IsExplicitInstallationPath()
    {
        Assert.Equal(
            @"C:\Users\Administrator\Documents\DM Reports & Decoded Images\DM475-866D76\CodeQuality",
            DmstHtmlScraper.ConfiguredReportDirectory);
    }

    [Fact]
    public void ParseHtml_RealDmstReportSummary_UsesDataAndDfcHeadingVerbatim()
    {
        const string html = """
            <html><body>
              <p>Verified: Tue 18-Aug-2026 09:51:39(375ms) PM</p>
              <table>
                <tr><td class="gc"><strong>Data</strong></td><td class="gc">&lt;F1&gt;01006961147042882172803282009</td></tr>
                <tr><td class="gc"><strong>Symbology</strong></td><td class="gc">GS1 DataMatrix</td></tr>
              </table>
              <table>
                <tr><th colspan="3">Data Format Check</th></tr>
                <tr><th colspan="3">GS1 Application Data Format: PASS</th></tr>
                <tr><td><strong>Name</strong></td><td><strong>Data</strong></td><td><strong>Check</strong></td></tr>
                <tr><td>GS1 Header</td><td>&lt;F1&gt;</td><td>PASS</td></tr>
              </table>
            </body></html>
            """;

        var report = DmstHtmlScraper.ParseHtml(html, FixturePath);

        Assert.Equal("<F1>01006961147042882172803282009", report.HtmlDecodedData);
        Assert.Equal("GS1 DataMatrix", report.HtmlSymbology);
        Assert.Equal("GS1 Application Data Format", report.HtmlApplicationStandard);
        Assert.NotNull(report.ScrapedDataFormatCheck);
        Assert.Single(report.ScrapedDataFormatCheck!.Rows);
    }

    [Fact]
    public void VerifiedStringsEquivalent_NormalizesOnlyHarmlessFormatting()
    {
        Assert.True(DmstHtmlScraper.VerifiedStringsEquivalent(
            "Tue 18-Aug-2026 05:10:32 ( 520 ms ) PM",
            "tue\u00a018-Aug-2026   05:10:32(520ms) pm"));
        Assert.False(DmstHtmlScraper.VerifiedStringsEquivalent(
            "Tue 18-Aug-2026 05:10:32(520ms) PM",
            "Tue 18-Aug-2026 05:10:33(520ms) PM"));
    }

    [Fact]
    public async Task TryMergeAsync_RenamedHtmlFile_DoesNotUseFilenameTimestampFallback()
    {
        string dir = Path.Combine(Path.GetTempPath(), $"vtccp-dmst-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);

        try
        {
            using var scraper = new DmstHtmlScraper(dir);
            scraper.Start();

            string tempPath = Path.Combine(dir, "dmst-writing.tmp");
            string htmlPath = Path.Combine(
                dir, "_F1_01006961147042882172803282009_2026-08-18_20-04-21-314.html");
            await File.WriteAllTextAsync(tempPath,
                "<html><body><p>Verified: Tue 18-Aug-2026 08:04:22 PM</p></body></html>");
            File.Move(tempPath, htmlPath);

            var incoming = new VerificationRecord
            {
                Symbology = "GS1 DataMatrix",
                VerificationDateTime = new DateTime(2026, 8, 18, 20, 4, 21),
                HtmlVerifiedString = "Tue 18-Aug-2026 08:04:21 PM",
            };

            var (merged, sourcePath) = await scraper.TryMergeAsync(incoming);

            Assert.Null(sourcePath);
            Assert.Null(merged.HtmlSourceFileName);
            Assert.Equal(HtmlReportProvenance.None, merged.HtmlReportProvenance);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public async Task TryMergeAsync_WaitsForCompleteStableHtmlBeforeConsuming()
    {
        string dir = Path.Combine(Path.GetTempPath(), $"vtccp-dmst-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);

        try
        {
            using var scraper = new DmstHtmlScraper(dir);
            scraper.Start();

            string htmlPath = Path.Combine(
                dir, "2026-08-18_20-04-21-314_partial.html");
            await File.WriteAllTextAsync(htmlPath,
                "<html><body><p>Verified: Tue 18-Aug-2026 08:04:21 PM");

            var incoming = new VerificationRecord
            {
                Symbology = "GS1 DataMatrix",
                VerificationDateTime = new DateTime(2026, 8, 18, 20, 4, 21),
                HtmlVerifiedString = "Tue 18-Aug-2026 08:04:21 PM",
            };

            Task<(VerificationRecord Record, string? SourcePath)> mergeTask =
                scraper.TryMergeAsync(incoming);
            await Task.Delay(350);
            Assert.False(mergeTask.IsCompleted);

            await File.AppendAllTextAsync(htmlPath, "</p></body></html>");
            var (merged, sourcePath) = await mergeTask;

            Assert.Equal(htmlPath, sourcePath);
            Assert.Equal(Path.GetFileName(htmlPath), merged.HtmlSourceFileName);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public async Task TryMergeAsync_HttpIdentityThenLocalHtml_GrantsFilesystemProvenance()
    {
        const string verified = "Tue 18-Aug-2026 08:04:21 PM";
        string dir = Path.Combine(Path.GetTempPath(), $"vtccp-dmst-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);

        try
        {
            var transportRecord = new VerificationRecord { Symbology = "GS1 DataMatrix" };
            var httpReport = DmstHtmlScraper.ParseHtml(
                $"<html><body><p>Verified: {verified}</p></body></html>",
                @"C:\HTTP_STREAM_PLACEHOLDER\2026-08-18_20-04-21-000_http.html",
                hasSyntheticSourcePath: true);
            VerificationRecord httpEnriched =
                DmstReportValidator.MergeAndValidate(transportRecord, httpReport);

            Assert.Equal(HtmlReportProvenance.HttpStreamOnly, httpEnriched.HtmlReportProvenance);
            Assert.Equal(verified, httpEnriched.HtmlVerifiedString);

            using var scraper = new DmstHtmlScraper(dir);
            scraper.Start();
            string htmlPath = Path.Combine(dir, "actual-dmst-report.html");
            await File.WriteAllTextAsync(htmlPath,
                $"<html><body><p>Verified: {verified}</p></body></html>");

            var (merged, sourcePath) = await scraper.TryMergeAsync(httpEnriched);

            Assert.Equal(htmlPath, sourcePath);
            Assert.Equal(HtmlReportProvenance.CorrelatedFilesystem, merged.HtmlReportProvenance);
            Assert.Equal("actual-dmst-report.html", merged.HtmlSourceFileName);
            Assert.Equal(verified, merged.HtmlVerifiedString);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public async Task TryMergeAsync_DefaultMode_PreservesCorrelatedHtmlFile()
    {
        const string verified = "Tue 18-Aug-2026 10:21:42(159ms) PM";
        string dir = Path.Combine(Path.GetTempPath(), $"vtccp-dmst-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);

        try
        {
            using var scraper = new DmstHtmlScraper(dir);
            scraper.Start();
            string htmlPath = Path.Combine(dir, "actual-dmst-report.html");
            await File.WriteAllTextAsync(htmlPath,
                $"<html><body><p>Verified: {verified}</p></body></html>");

            var incoming = new VerificationRecord
            {
                Symbology = "GS1 DataMatrix",
                HtmlVerifiedString = verified,
            };
            var (_, sourcePath) = await scraper.TryMergeAsync(incoming);

            Assert.Equal(htmlPath, sourcePath);
            Assert.True(File.Exists(htmlPath));
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void MergeAndValidate_MismatchedSourceBasename_CannotGrantFilesystemProvenance()
    {
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix" };
        var report = new DmstHtmlReport
        {
            ParseSucceeded = true,
            SourceFilePath = @"C:\CodeQuality\actual-dmst-report.html",
            HtmlSourceFileName = "claimed-other-report.html",
            HtmlVerifiedString = "Tue 18-Aug-2026 08:04:21 PM",
        };

        VerificationRecord merged = DmstReportValidator.MergeAndValidate(record, report);

        Assert.Equal(HtmlReportProvenance.None, merged.HtmlReportProvenance);
        Assert.Null(merged.HtmlSourceFileName);
        Assert.Null(merged.WebscanSourcePath);
    }

    [Fact]
    public async Task Stop_CancelsQueuedPartialReportWithoutLatePendingAdd()
    {
        string dir = Path.Combine(Path.GetTempPath(), $"vtccp-dmst-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);

        try
        {
            using var scraper = new DmstHtmlScraper(dir);
            scraper.Start();

            string htmlPath = Path.Combine(
                dir, "2026-08-18_20-04-21-314_stopping.html");
            await File.WriteAllTextAsync(htmlPath,
                "<html><body><p>Verified: Tue 18-Aug-2026 08:04:21 PM");
            await Task.Delay(175);

            scraper.Stop();
            await File.AppendAllTextAsync(htmlPath, "</p></body></html>");

            using var timeout = new CancellationTokenSource(TimeSpan.FromMilliseconds(350));
            var incoming = new VerificationRecord
            {
                Symbology = "GS1 DataMatrix",
                VerificationDateTime = new DateTime(2026, 8, 18, 20, 4, 21),
                HtmlVerifiedString = "Tue 18-Aug-2026 08:04:21 PM",
            };
            var (_, sourcePath) = await scraper.TryMergeAsync(incoming, timeout.Token);

            Assert.Null(sourcePath);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void MergeAndValidate_DoesNotReconstructMissingHtmlDataFormatCheck()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            DecodedData = "<F1>01006961147042882172803282009",
        };
        var html = new DmstHtmlReport
        {
            ParseSucceeded = true,
            SourceFilePath = FixturePath,
            HtmlSourceFileName = Path.GetFileName(FixturePath),
        };

        VerificationRecord merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Null(merged.DataFormatCheck);
    }

    [Fact]
    public void MergeAndValidate_PreservesLiteralHtmlDataFormatCheckRows()
    {
        var scraped = new DataFormatCheckResult
        {
            Overall = OverallPassFail.Fail,
            Standard = "GS1 Application Data Format",
            Rows =
            [
                new DataFormatCheckRow
                {
                    Name = "AI (21) Serial",
                    Data = "72803282009",
                    Check = "FAIL — TruCheck literal",
                },
            ],
        };
        var html = new DmstHtmlReport
        {
            ParseSucceeded = true,
            SourceFilePath = FixturePath,
            HtmlSourceFileName = Path.GetFileName(FixturePath),
            ScrapedDataFormatCheck = scraped,
        };

        VerificationRecord merged = DmstReportValidator.MergeAndValidate(
            new VerificationRecord { Symbology = "GS1 DataMatrix" }, html);

        Assert.Same(scraped, merged.DataFormatCheck);
        Assert.Equal("FAIL — TruCheck literal", merged.DataFormatCheck!.Rows[0].Check);
    }

    // ══ MergeAndValidate — fractional grade × threshold boundary ════════════════
    //
    // These tests verify that pass/fail is decided from the PARSED decimal grade,
    // not from a letter-midpoint approximation.  The critical regression was:
    //   LetterToNumericGrade("B") = 3.0 → a B/2.5 grade would PASS a 3.0 threshold
    //   using midpoints, when in reality 2.5 < 3.0 and must FAIL.

    // Helper: build a minimal multi-mode DmstHtmlReport with specific linear grade.
    private static DmstHtmlReport MakeMultiModeReport(
        string letter, decimal numeric) => new()
    {
        ParseSucceeded         = true,
        ScanDateTime           = new DateTime(2026, 1, 15, 10, 30, 45),
        SourceFilePath         = FixturePath,
        IsMultiMode            = true,
        LinearSymbology        = "EAN-13",
        LinearDecodedData      = "5901234123457",
        LinearOverallGrade     = letter,
        LinearOverallGradeNumeric = numeric,
        LinearStandard         = "ISO/IEC 15416",
        OverallGrade           = "A",   // 2D symbol — irrelevant to these tests
    };

    [Fact]
    public void MergeAndValidate_BLetter_2point5_Against3point0MinPassRaw_IsFail()
    {
        // B/2.5 < 3.0 threshold → FAIL.
        // Old code: LetterToNumericGrade("B") = 3.0 ≥ 3.0 → PASS (wrong).
        var html   = MakeMultiModeReport("B", 2.5m);
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix", MinPassRaw = 3.0m };

        var merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Equal(OverallPassFail.Fail, merged.LinearOverallGrade?.PassFail);
    }

    [Fact]
    public void MergeAndValidate_BLetter_3point5_Against3point0MinPassRaw_IsPass()
    {
        // B/3.5 ≥ 3.0 threshold → PASS.
        var html   = MakeMultiModeReport("B", 3.5m);
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix", MinPassRaw = 3.0m };

        var merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Equal(OverallPassFail.Pass, merged.LinearOverallGrade?.PassFail);
    }

    [Fact]
    public void MergeAndValidate_BLetter_3point4_Against3point5MinPassRaw_IsFail()
    {
        // B/3.4 < 3.5 threshold → FAIL.
        var html   = MakeMultiModeReport("B", 3.4m);
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix", MinPassRaw = 3.5m };

        var merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Equal(OverallPassFail.Fail, merged.LinearOverallGrade?.PassFail);
    }

    [Fact]
    public void MergeAndValidate_CLetterGrade_2point0_AgainstLetterThresholdC_IsPass()
    {
        // C/2.0 against MinPassGrade="C" (ISO 15416 C-band floor = 1.5) → PASS.
        var html   = MakeMultiModeReport("C", 2.0m);
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix", MinPassGrade = "C" };

        var merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Equal(OverallPassFail.Pass, merged.LinearOverallGrade?.PassFail);
    }

    [Fact]
    public void MergeAndValidate_DLetterGrade_1point4_AgainstLetterThresholdC_IsFail()
    {
        // D/1.4 against MinPassGrade="C" (floor = 1.5) → FAIL (1.4 < 1.5).
        var html   = MakeMultiModeReport("D", 1.4m);
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix", MinPassGrade = "C" };

        var merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Equal(OverallPassFail.Fail, merged.LinearOverallGrade?.PassFail);
    }

    [Fact]
    public void MergeAndValidate_LinearNumeric_PreservedOnMergedRecord()
    {
        // The merged record's LinearOverallGrade.NumericGrade must carry the
        // parsed decimal (2.5m), not a letter midpoint (3.0m for B).
        var html   = MakeMultiModeReport("B", 2.5m);
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix" };

        var merged = DmstReportValidator.MergeAndValidate(record, html);

        Assert.Equal(2.5m, merged.LinearOverallGrade?.NumericGrade);
    }

    // ══ GS1 check-digit boundary cases ══════════════════════════════════════════

    [Theory]
    [InlineData("5901234123457", true)]   // valid EAN-13
    [InlineData("5901234123458", false)]  // wrong check digit
    [InlineData("012345678905",  true)]   // valid UPC-A
    [InlineData("012345678904",  false)]  // wrong check digit
    [InlineData("40170725",      true)]   // valid EAN-8
    [InlineData("40170724",      false)]  // wrong check digit
    public void ValidateGs1CheckDigit_KnownValues(string digits, bool expected)
    {
        // Access via the internal helper through BuildLinearDataFormatCheck —
        // pass a symbology that matches the digit length so the check is reached.
        string sym = digits.Length switch
        {
            13 => "EAN-13",
            12 => "UPC-A",
            8  => "EAN-8",
            _  => "EAN-13"
        };

        var result = DmstReportValidator.BuildLinearDataFormatCheck(digits, sym);

        // When expected is valid → Pass; invalid → Fail
        Assert.NotNull(result);
        Assert.Equal(
            expected ? OverallPassFail.Pass : OverallPassFail.Fail,
            result.Overall);
    }

    [Fact]
    public void ParseHtml_EmbeddedCognexLogoAndBarcode_UsesBarcodeImageOnly()
    {
        string cognexLogo = new('A', 200);
        const string barcodeImage = "YmFyY29kZS1pbWFnZQ==";
        string html = "<html><body>" +
            "<p>Verified: Tue 18-Aug-2026 09:51:39 PM</p>" +
            $"<img src=\"data:image/png;base64,{cognexLogo}\" />" +
            new string(' ', 700) +
            $"<div>Barcode Image — <img src=\"data:image/png;base64,{barcodeImage}\" /></div>" +
            "</body></html>";

        var report = DmstHtmlScraper.ParseHtml(html, FixturePath);

        Assert.Equal(barcodeImage, report.HtmlBarcodeImageBase64);
        Assert.NotEqual(cognexLogo, report.HtmlBarcodeImageBase64);
    }

    [Fact]
    public void ParseHtml_UnknownUnlabelledImage_IsNotPromotedToBarcodeEvidence()
    {
        const string onlyImage = "Y29nbmV4LWxvZ28=";
        string html = $$"""
            <html><body>
              <p>Verified: Tue 18-Aug-2026 09:51:39 PM</p>
              <img src="data:image/png;base64,{{onlyImage}}" />
            </body></html>
            """;

        var report = DmstHtmlScraper.ParseHtml(html, FixturePath);

        Assert.Null(report.HtmlBarcodeImageBase64);
    }
}
