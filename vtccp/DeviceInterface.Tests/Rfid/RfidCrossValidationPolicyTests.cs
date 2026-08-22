using DeviceInterface.Dmst;
using DeviceInterface.Rfid;
using ExcelEngine.Models;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class RfidCrossValidationPolicyTests
{
    [Fact]
    public void ShouldRun_SkipsGs1DataMatrixAfterTruCheckApplicationPass()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            ApplicationPass = "Pass",
        };

        Assert.False(RfidCrossValidationPolicy.ShouldRun(record));
    }

    [Fact]
    public void ShouldRun_SkipsGs1DataMatrixAfterCorrelatedTruCheckDfcPass()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
            },
        };

        Assert.False(RfidCrossValidationPolicy.ShouldRun(record));
    }

    [Fact]
    public void ShouldRun_KeepsRfidValidationForFailedOrNonGs1Results()
    {
        var failedGs1 = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            ApplicationPass = "Fail (Quality)",
        };
        var passedQr = new VerificationRecord
        {
            Symbology = "QR Code",
            ApplicationPass = "Pass",
        };

        Assert.True(RfidCrossValidationPolicy.ShouldRun(failedGs1));
        Assert.True(RfidCrossValidationPolicy.ShouldRun(passedQr));
    }
}