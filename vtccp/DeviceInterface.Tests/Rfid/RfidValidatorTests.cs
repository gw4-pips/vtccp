using DeviceInterface.Rfid;
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
}