using VtccpApp.Services;
using Xunit;

namespace DeviceInterface.Tests.VtccpApp;

public sealed class AppVersionDisplayTests
{
    [Fact]
    public void Current_ReadsTheContainingAssemblyVersion()
    {
        string expected = AppVersionDisplay.Format(
            typeof(AppVersionDisplay).Assembly.GetName().Version);

        Assert.Equal(expected, AppVersionDisplay.Current);
    }

    [Fact]
    public void Format_UsesMajorMinorPatchAndOmitsAssemblyRevision()
    {
        Assert.Equal("v1.5.52", AppVersionDisplay.Format(new Version(1, 5, 52, 0)));
    }

    [Fact]
    public void Format_HandlesVersionsWithoutBuildComponent()
    {
        Assert.Equal("v1.5", AppVersionDisplay.Format(new Version(1, 5)));
    }
}