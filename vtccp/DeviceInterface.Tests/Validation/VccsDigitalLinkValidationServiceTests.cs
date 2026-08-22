namespace DeviceInterface.Tests.Validation;

using DeviceInterface.Validation;
using ExcelEngine.Models;
using GS1.Encoders;
using Xunit;

public sealed class VccsDigitalLinkValidationServiceTests
{
    [Fact]
    public void Validate_ValidDigitalLink_ReportsVccsPass()
    {
        var result = VccsDigitalLinkValidationService.Validate(
            "https://id.gs1.org/01/09506000134352/21/1234",
            new AcceptingEngine());

        Assert.Equal(DigitalLinkValidationStatus.Valid, result.Status);
        Assert.Equal(DigitalLinkValidationResult.VccsSource, result.Source);
        Assert.Equal("GS1 Barcode Syntax Engine 1.4.0", result.EngineVersion);
    }

    [Fact]
    public void Validate_EngineProvidesParsedAiData_IncludesItInDetail()
    {
        var result = VccsDigitalLinkValidationService.Validate(
            "https://id.gs1.org/01/09506000134352/21/1234",
            new DetailEngine());

        Assert.Equal(
            "Parsed GS1 AI data: (01)09506000134352(21)1234 Validated with the official GS1 Barcode Syntax Engine.",
            result.Detail);
    }

    [Fact]
    public void Validate_InvalidDigitalLink_ReportsVccsFail()
    {
        var result = VccsDigitalLinkValidationService.Validate(
            "https://id.gs1.org/01/not-a-gtin",
            new RejectingEngine());

        Assert.Equal(DigitalLinkValidationStatus.Invalid, result.Status);
        Assert.Equal("Invalid AI (01) value.", result.Detail);
    }

    [Fact]
    public void Validate_NonDigitalLink_IsExplicitlyNotApplicable()
    {
        var result = VccsDigitalLinkValidationService.Validate(
            "09506000134352",
            new ThrowingEngine());

        Assert.Equal(DigitalLinkValidationStatus.NotApplicable, result.Status);
        Assert.Null(result.EngineVersion);
    }

    [Fact]
    public void ValidateElementString_ValidGs1DataMatrix_ReportsVccsPass()
    {
        var result = VccsDigitalLinkValidationService.ValidateElementString(
            "(01)09506000134352(21)72803288707",
            new AcceptingEngine());

        Assert.Equal(DigitalLinkValidationStatus.Valid, result.Status);
        Assert.Equal(DigitalLinkValidationResult.VccsElementStringSource, result.Source);
        Assert.Equal("GS1 Barcode Syntax Engine 1.4.0", result.EngineVersion);
    }

    [Theory]
    [InlineData("<F1>01095060001343522172803288707")]
    [InlineData("]d201095060001343522172803288707")]
    public void ValidateElementString_DataManGs1DataMatrixForm_ReportsVccsPass(string data)
    {
        var result = VccsDigitalLinkValidationService.ValidateElementString(
            data,
            new AcceptingEngine());

        Assert.Equal(DigitalLinkValidationStatus.Valid, result.Status);
        Assert.Equal(DigitalLinkValidationResult.VccsElementStringSource, result.Source);
    }

    [Fact]
    public void ValidateElementString_NonGs1Data_IsExplicitlyNotApplicable()
    {
        var result = VccsDigitalLinkValidationService.ValidateElementString(
            "plain decoded data",
            new ThrowingEngine());

        Assert.Equal(DigitalLinkValidationStatus.NotApplicable, result.Status);
        Assert.Null(result.EngineVersion);
    }

    [Fact]
    public void Validate_MissingNativeRuntime_IsExplicitlyUnavailable()
    {
        var result = VccsDigitalLinkValidationService.Validate(
            "https://id.gs1.org/01/09506000134352",
            new MissingRuntimeEngine());

        Assert.Equal(DigitalLinkValidationStatus.Unavailable, result.Status);
        Assert.Contains("unavailable", result.Detail!, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class AcceptingEngine : IGs1DigitalLinkSyntaxEngine
    {
        public void Validate(string digitalLinkUri) { }
    }

    private sealed class DetailEngine : IGs1DigitalLinkSyntaxEngine
    {
        public string? ParsedAiData => "(01)09506000134352(21)1234";
        public void Validate(string digitalLinkUri) { }
    }

    private sealed class RejectingEngine : IGs1DigitalLinkSyntaxEngine
    {
        public void Validate(string digitalLinkUri)
            => throw new GS1EncoderParameterException("Invalid AI (01) value.");
    }

    private sealed class ThrowingEngine : IGs1DigitalLinkSyntaxEngine
    {
        public void Validate(string digitalLinkUri)
            => throw new Xunit.Sdk.XunitException("The engine must not run.");
    }

    private sealed class MissingRuntimeEngine : IGs1DigitalLinkSyntaxEngine
    {
        public void Validate(string digitalLinkUri)
            => throw new DllNotFoundException("gs1encoders.dll");
    }
}