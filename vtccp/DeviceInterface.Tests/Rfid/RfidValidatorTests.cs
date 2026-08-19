using DeviceInterface.Rfid;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class RfidValidatorTests
{
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