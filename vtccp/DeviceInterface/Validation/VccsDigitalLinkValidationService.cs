namespace DeviceInterface.Validation;

using GS1.Encoders;
using ExcelEngine.Models;

/// <summary>
/// VCCS-owned GS1 Digital Link syntax validation.
/// It deliberately validates only a decoded HTTP(S) Digital Link URI and leaves
/// every verifier-provided value untouched.
/// </summary>
public static class VccsDigitalLinkValidationService
{
    public const string EngineVersion = "GS1 Syntax Engine 1.4.0";

    /// <summary>
    /// Validates a decoded GS1 Digital Link URI with the official GS1 engine.
    /// Non-Digital-Link data is explicitly not applicable; a missing native
    /// runtime is explicitly unavailable and is never treated as a pass.
    /// </summary>
    public static DigitalLinkValidationResult Validate(string? decodedData)
        => Validate(decodedData, new Gs1DigitalLinkSyntaxEngine());

    /// <summary>
    /// Injectable overload used by regression tests. Production callers use the
    /// official GS1 engine through <see cref="Validate(string?)"/>.
    /// </summary>
    public static DigitalLinkValidationResult Validate(
        string? decodedData,
        IGs1DigitalLinkSyntaxEngine engine)
    {
        if (!IsHttpDigitalLinkUri(decodedData))
        {
            return new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.NotApplicable,
                Detail = "Decoded verifier data is not a GS1 Digital Link URI.",
            };
        }

        try
        {
            engine.Validate(decodedData!);
            return new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                EngineVersion = EngineVersion,
                Detail = "Validated with the official GS1 Syntax Engine.",
            };
        }
        catch (GS1EncoderParameterException ex)
        {
            return Invalid(ex.Message);
        }
        catch (GS1EncoderDigitalLinkException ex)
        {
            return Invalid(ex.Message);
        }
        catch (GS1EncoderGeneralException ex)
        {
            return Unavailable(ex.Message);
        }
        catch (DllNotFoundException ex)
        {
            return Unavailable(ex.Message);
        }
        catch (BadImageFormatException ex)
        {
            return Unavailable(ex.Message);
        }
        catch (EntryPointNotFoundException ex)
        {
            return Unavailable(ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            return Unavailable(ex.Message);
        }
    }

    private static bool IsHttpDigitalLinkUri(string? value)
        => Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) &&
           (uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase));

    private static DigitalLinkValidationResult Invalid(string? reason)
        => new()
        {
            Status = DigitalLinkValidationStatus.Invalid,
            EngineVersion = EngineVersion,
            Detail = string.IsNullOrWhiteSpace(reason)
                ? "The GS1 Syntax Engine rejected the Digital Link URI."
                : reason,
        };

    private static DigitalLinkValidationResult Unavailable(string? reason)
        => new()
        {
            Status = DigitalLinkValidationStatus.Unavailable,
            Detail = string.IsNullOrWhiteSpace(reason)
                ? "The GS1 Syntax Engine runtime is unavailable."
                : $"The GS1 Syntax Engine runtime is unavailable: {reason}",
        };
}

/// <summary>Small seam so validation outcome tests do not depend on a native DLL.</summary>
public interface IGs1DigitalLinkSyntaxEngine
{
    void Validate(string digitalLinkUri);
}

internal sealed class Gs1DigitalLinkSyntaxEngine : IGs1DigitalLinkSyntaxEngine
{
    public void Validate(string digitalLinkUri)
    {
        using var encoder = new GS1Encoder();
        encoder.DataStr = digitalLinkUri;
    }
}