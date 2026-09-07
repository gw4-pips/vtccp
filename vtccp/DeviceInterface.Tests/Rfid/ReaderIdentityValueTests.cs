using DeviceInterface.Rfid;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class ReaderIdentityValueTests
{
    [Theory]
    [InlineData("1.8.0", "1.8.0")]
    [InlineData("  KE00048  ", "KE00048")]
    public void FromSdkResponse_SuccessfulLiteral_ReturnsValue(
        string response,
        string expected)
    {
        Assert.Equal(expected, ReaderIdentityValue.FromSdkResponse(0, response));
    }

    [Theory]
    [InlineData(0U, null)]
    [InlineData(0U, "")]
    [InlineData(0U, "   ")]
    [InlineData(1U, "KE00048")]
    public void FromSdkResponse_AbsentOrFailed_ReturnsNull(
        uint returnCode,
        string? response)
    {
        Assert.Null(ReaderIdentityValue.FromSdkResponse(returnCode, response));
    }

    [Theory]
    [InlineData("KE00048\nspoof")]
    [InlineData("FW\u0000VERSION")]
    [InlineData("bad\uFFFDvalue")]
    public void FromSdkResponse_MalformedValue_ReturnsNull(string response)
    {
        Assert.Null(ReaderIdentityValue.FromSdkResponse(0, response));
    }
}