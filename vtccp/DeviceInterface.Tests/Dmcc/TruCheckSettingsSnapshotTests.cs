namespace DeviceInterface.Tests.Dmcc;

using DeviceInterface.Dmcc;
using ExcelEngine.Models;
using Xunit;

public sealed class TruCheckSettingsSnapshotTests
{
    [Fact]
    public void Apply_PushResultResponses_PopulatesAllThreeReportSettings()
    {
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix" };

        VerificationRecord enriched = TruCheckSettingsSnapshot.Apply(
            record,
            Ok("4"),
            Ok("1"),
            Ok("0"));

        Assert.Equal("Custom", enriched.ApplicationStandardSetting);
        Assert.Equal("GS1", enriched.DataFormatCheckSetting);
        Assert.Equal("User Set", enriched.ApertureSettingMode);
    }

    [Fact]
    public void Apply_UnavailablePushResponses_LeavesSettingsBlank()
    {
        var record = new VerificationRecord { Symbology = "GS1 DataMatrix" };

        VerificationRecord enriched = TruCheckSettingsSnapshot.Apply(
            record,
            DmccResponse.Parse(string.Empty),
            DmccResponse.Parse(string.Empty),
            DmccResponse.Parse(string.Empty));

        Assert.Null(enriched.ApplicationStandardSetting);
        Assert.Null(enriched.DataFormatCheckSetting);
        Assert.Null(enriched.ApertureSettingMode);
    }

    private static DmccResponse Ok(string body)
        => DmccResponse.Parse($"||:::2[0]\r\n||:::2[0]{body}\r\n");
}