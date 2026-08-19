using DeviceInterface.Rfid;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class AsReaderLockStatusTests
{
    [Theory]
    [InlineData(0u, "Unlocked")]
    [InlineData(1u, "Locked")]
    [InlineData(2u, "PermaLocked")]
    [InlineData(3u, "Unknown")]
    [InlineData(4u, "Unknown")]
    [InlineData(99u, "Unknown")]
    public void FromCheckTagStatus_MapsTheSdkDirectReturn(uint statusCode, string expected)
    {
        Assert.Equal(expected, AsReaderLockStatus.FromCheckTagStatus(statusCode));
    }
}