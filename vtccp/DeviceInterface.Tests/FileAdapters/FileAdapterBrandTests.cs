namespace DeviceInterface.Tests.FileAdapters;

using DeviceInterface.FileAdapters;
using ExcelEngine.Models;
using Xunit;

/// <summary>
/// Tests that <see cref="AxiconFileAdapter"/> and <see cref="MicroscanLvsFileAdapter"/>
/// always set the correct <see cref="VerificationRecord.VerifierBrand"/> on every
/// record they produce — preventing the PDF Device header row from showing "—".
///
/// The PDF header renders:
///   string device = r.DeviceModel ?? r.VerifierBrand ?? "—";
///
/// File-export adapters never populate DeviceModel (no SDK connection), so
/// VerifierBrand is the only guard against the fallback "—" placeholder.
/// </summary>
public sealed class FileAdapterBrandTests
{
    // ── AxiconFileAdapter ─────────────────────────────────────────────────────

    [Fact]
    public void Axicon_BuildRecord_SetsVerifierBrand()
    {
        VerificationRecord r = AxiconFileAdapter.BuildRecord();

        Assert.Equal("AXICON", r.VerifierBrand);
    }

    [Fact]
    public void Axicon_BuildRecord_WithContent_SetsVerifierBrand()
    {
        // Simulate a file being read — even with unrecognised content the brand is set.
        const string rawContent = "some unrecognised axicon export content";

        VerificationRecord r = AxiconFileAdapter.BuildRecord(rawContent);

        Assert.Equal("AXICON", r.VerifierBrand);
    }

    [Fact]
    public void Axicon_BuildRecord_WithNullContent_SetsVerifierBrand()
    {
        // Simulate a file that could not be read — brand is still set.
        VerificationRecord r = AxiconFileAdapter.BuildRecord(rawContent: null);

        Assert.Equal("AXICON", r.VerifierBrand);
    }

    [Fact]
    public void Axicon_BuildRecord_DeviceModel_IsNull()
    {
        // DeviceModel is always null for file-export adapters (no SDK connection).
        // PdfReportGenerator falls through to VerifierBrand when DeviceModel is null.
        VerificationRecord r = AxiconFileAdapter.BuildRecord();

        Assert.Null(r.DeviceModel);
    }

    [Fact]
    public void Axicon_BuildRecord_PdfDeviceHeaderValue_IsAxicon()
    {
        // Reproduce the exact expression from PdfReportGenerator.BuildHeader line ~281:
        //   string device = r.DeviceModel ?? r.VerifierBrand ?? "\u2014";
        VerificationRecord r = AxiconFileAdapter.BuildRecord();

        string device = r.DeviceModel ?? r.VerifierBrand ?? "\u2014";

        Assert.Equal("AXICON", device);
    }

    [Fact]
    public void Axicon_BrandConstant_MatchesBrandPatterns()
    {
        // "AXICON" must remain all-caps to match PdfReportGenerator.BrandPatterns
        // ("Axicon" substring → "AXICON" brand).
        Assert.Equal("AXICON", AxiconFileAdapter.Brand);
    }

    // ── MicroscanLvsFileAdapter ───────────────────────────────────────────────

    [Fact]
    public void Lvs_BuildRecord_SetsVerifierBrand()
    {
        VerificationRecord r = MicroscanLvsFileAdapter.BuildRecord();

        Assert.Equal("OMRON/LVS", r.VerifierBrand);
    }

    [Fact]
    public void Lvs_BuildRecord_WithContent_SetsVerifierBrand()
    {
        const string rawContent = "some unrecognised lvs export content";

        VerificationRecord r = MicroscanLvsFileAdapter.BuildRecord(rawContent);

        Assert.Equal("OMRON/LVS", r.VerifierBrand);
    }

    [Fact]
    public void Lvs_BuildRecord_WithNullContent_SetsVerifierBrand()
    {
        VerificationRecord r = MicroscanLvsFileAdapter.BuildRecord(rawContent: null);

        Assert.Equal("OMRON/LVS", r.VerifierBrand);
    }

    [Fact]
    public void Lvs_BuildRecord_DeviceModel_IsNull()
    {
        VerificationRecord r = MicroscanLvsFileAdapter.BuildRecord();

        Assert.Null(r.DeviceModel);
    }

    [Fact]
    public void Lvs_BuildRecord_PdfDeviceHeaderValue_IsOmronLvs()
    {
        // Reproduce the exact expression from PdfReportGenerator.BuildHeader line ~281:
        //   string device = r.DeviceModel ?? r.VerifierBrand ?? "\u2014";
        VerificationRecord r = MicroscanLvsFileAdapter.BuildRecord();

        string device = r.DeviceModel ?? r.VerifierBrand ?? "\u2014";

        Assert.Equal("OMRON/LVS", device);
    }

    [Fact]
    public void Lvs_BrandConstant_MatchesBrandPatterns()
    {
        // "OMRON/LVS" must match the "LVS" / "Omron" / "Microscan" entries
        // in PdfReportGenerator.BrandPatterns.
        Assert.Equal("OMRON/LVS", MicroscanLvsFileAdapter.Brand);
    }
}
