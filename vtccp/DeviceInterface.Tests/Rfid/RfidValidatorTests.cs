using DeviceInterface.Rfid;
using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class RfidValidatorTests
{
    [Fact]
    public void AssessTruCheckValidation_PassingNativeDfc_DoesNotRequireVeriWedge()
    {
        var assessment = RfidValidator.AssessTruCheckValidation(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Rows = [new DataFormatCheckRow { Name = "AI (01)", Data = "09506000134352", Check = "PASS" }],
            },
        });

        Assert.True(assessment.Usable);
        Assert.False(assessment.Failed);
        Assert.False(assessment.RequiresVeriWedge);
    }

    [Fact]
    public void AssessTruCheckValidation_FailedNativeDfc_RequiresVeriWedge()
    {
        var assessment = RfidValidator.AssessTruCheckValidation(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Fail,
                Rows = [new DataFormatCheckRow { Name = "AI (01)", Data = "09506000134352", Check = "FAIL" }],
            },
        });

        Assert.True(assessment.Usable);
        Assert.True(assessment.Failed);
        Assert.True(assessment.RequiresVeriWedge);
    }

    [Fact]
    public void AssessTruCheckValidation_MissingNativeDfc_RequiresVeriWedge()
    {
        var assessment = RfidValidator.AssessTruCheckValidation(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            ApplicationPass = "Pass",
        });

        Assert.False(assessment.Usable);
        Assert.False(assessment.Failed);
        Assert.True(assessment.RequiresVeriWedge);
    }

    [Theory]
    [InlineData("<F1>01006961147042882172803282009")]
    [InlineData("<GS>01006961147042882172803282009")]
    [InlineData("\u001d01006961147042882172803282009")]
    public void ExtractAi01_RecognizesTruCheckAndGsSeparators(string decodedData)
    {
        Assert.Equal("00696114704288", RfidValidator.ExtractAi01(decodedData));
    }

    [Theory]
    [InlineData("<F1>01006961147042882172803282009")]
    [InlineData("<GS>01006961147042882172803282009")]
    [InlineData("\u001d01006961147042882172803282009")]
    public void ExtractAi21_RecognizesTruCheckAndGsSeparators(string decodedData)
    {
        Assert.Equal("72803282009", RfidValidator.ExtractAi21(decodedData));
    }

    [Fact]
    public void ExtractAi21_StopsAtTextualGsSeparator()
    {
        const string data = "<F1>01006961147042882172803282009<GS>10LOT-42";

        Assert.Equal("72803282009", RfidValidator.ExtractAi21(data));
    }

    [Theory]
    [InlineData("UPCA", "696114704318", "00696114704318")]
    [InlineData("UPC-A", "696114704318", "00696114704318")]
    [InlineData("EAN13", "0696114704318", "00696114704318")]
    [InlineData("EAN-13", "0696114704318", "00696114704318")]
    public void NormalizeLinearGtin14_UsesRequiredPadding(
        string symbology,
        string decodedData,
        string expected)
    {
        Assert.Equal(expected,
            RfidValidator.NormalizeLinearGtin14(symbology, decodedData));
    }

    [Theory]
    [InlineData("EAN8", "12345670")]
    [InlineData("UPCE", "01234565")]
    [InlineData("UPCA", "696114704318+12")]
    [InlineData("EAN13", "0696114704318+12345")]
    public void NormalizeLinearGtin14_LeavesUnconfirmedFormsUnchanged(
        string symbology,
        string decodedData)
    {
        Assert.Null(RfidValidator.NormalizeLinearGtin14(symbology, decodedData));
    }

    [Theory]
    [InlineData("UPCA", "696114704318")]
    [InlineData("EAN13", "0696114704318")]
    public void Validate_EquivalentLinearGtinAndEpcGtin14_ReturnsPass(
        string symbology,
        string decodedData)
    {
        var result = new RfidValidator().Validate(
            [new EpcReadResult
            {
                EpcBytes = Convert.FromHexString("30342A7CC844C7D0F36A0676"),
            }],
            new VerificationRecord
            {
                Symbology = symbology,
                DecodedData = decodedData,
            },
            scanWindowMs: 1000);

        Assert.Equal(RfidValidationStatus.Pass, result.Status);
        Assert.Equal("00696114704318", result.RfidGtin14);
        Assert.Equal("00696114704318", result.BarcodeGtin14);
        Assert.Null(result.MismatchDetail);
    }

    [Fact]
    public void Validate_TrueLinearGtinMismatchStillReturnsFail()
    {
        var result = new RfidValidator().Validate(
            [new EpcReadResult
            {
                EpcBytes = Convert.FromHexString("30342A7CC844C7D0F36A0676"),
            }],
            new VerificationRecord
            {
                Symbology = "UPCA",
                DecodedData = "036000291452",
            },
            scanWindowMs: 1000);

        Assert.Equal(RfidValidationStatus.Fail, result.Status);
        Assert.Contains("GTIN14:RFID=00696114704318,BC=00036000291452",
            result.MismatchDetail);
    }
}