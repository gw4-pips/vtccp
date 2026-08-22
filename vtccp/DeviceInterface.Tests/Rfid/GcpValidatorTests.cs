using System.Text;
using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class GcpValidatorTests
{
    [Fact]
    public void Validate_DistinguishesValidInvalidAndNotFoundPrefixes()
    {
        var validator = new GcpValidator(LoadTable(
            "<GCPPrefixFormatList date=\"2026-06-03T11:14:42.028Z\">" +
            "<entry prefix=\"0614141\" gcpLength=\"7\" />" +
            "</GCPPrefixFormatList>"));

        Assert.Equal(
            GcpValidationStatus.Valid,
            validator.Validate(CreateSgtin(partition: 5, companyPrefix: "0614141")));
        Assert.Equal(
            GcpValidationStatus.Invalid,
            validator.Validate(CreateSgtin(partition: 4, companyPrefix: "0614141")));
        Assert.Equal(
            GcpValidationStatus.NotFound,
            validator.Validate(CreateSgtin(partition: 5, companyPrefix: "9999999")));
    }

    [Fact]
    public void Validate_ReturnsNotCheckedForAnEpcWithoutGcpFields()
    {
        var validator = new GcpValidator(LoadTable(
            "<GCPPrefixFormatList><entry prefix=\"0614141\" gcpLength=\"7\" /></GCPPrefixFormatList>"));

        var unknown = new ParsedEpc
        {
            EpcBytes = [],
            Scheme = EpcScheme.Unknown,
        };

        Assert.Equal(GcpValidationStatus.NotChecked, validator.Validate(unknown));
    }

    [Theory]
    [InlineData(0, 12)]
    [InlineData(5, 7)]
    [InlineData(6, 6)]
    public void GetEncodedGcpLength_UsesTheSgtinPartitionTable(int partition, int expectedLength)
    {
        Assert.Equal(
            expectedLength,
            GcpValidator.GetEncodedGcpLength(CreateSgtin(partition, "0614141")));
    }

    private static GcpLengthTable LoadTable(string xml)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        return GcpLengthTable.LoadFromStream(stream);
    }

    private static ParsedEpc CreateSgtin(int partition, string companyPrefix) => new()
    {
        EpcBytes = [],
        Scheme = EpcScheme.Sgtin96,
        Partition = partition,
        CompanyPrefix = companyPrefix,
    };
}