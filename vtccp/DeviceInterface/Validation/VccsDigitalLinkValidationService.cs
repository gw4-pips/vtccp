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

        return ValidateWithEngine(decodedData!, engine, DigitalLinkValidationResult.VccsSource,
            "The GS1 Syntax Engine rejected the Digital Link URI.",
            engine.Validate);
    }

    /// <summary>
    /// Validates a bracketed GS1 Element String (or an AIM-prefixed GS1 Element
    /// String) with the official GS1 engine. It never treats arbitrary decoded
    /// data as GS1 syntax.
    /// </summary>
    public static DigitalLinkValidationResult ValidateElementString(string? decodedData)
        => ValidateElementString(decodedData, new Gs1DigitalLinkSyntaxEngine());

    /// <summary>Injectable Element String overload used by regression tests.</summary>
    public static DigitalLinkValidationResult ValidateElementString(
        string? decodedData,
        IGs1DigitalLinkSyntaxEngine engine)
    {
        if (!LooksLikeGs1ElementString(decodedData))
        {
            return new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.NotApplicable,
                Detail = "Decoded verifier data is not a GS1 Element String.",
            };
        }

        return ValidateWithEngine(
            decodedData!,
            engine,
            DigitalLinkValidationResult.VccsElementStringSource,
            "The GS1 Syntax Engine rejected the GS1 Element String.",
            engine.ValidateElementString);
    }

    /// <summary>
    /// Recognises the bracketed representation emitted by the report parser and
    /// the raw AIM ]d2 representation emitted by some Data Matrix readers.
    /// </summary>
    public static bool LooksLikeGs1ElementString(string? decodedData)
    {
        string value = decodedData?.Trim() ?? string.Empty;
        return value.StartsWith("(01)", StringComparison.Ordinal) ||
               value.StartsWith("<F1>", StringComparison.OrdinalIgnoreCase) ||
               value.StartsWith("]d2", StringComparison.OrdinalIgnoreCase);
    }

    private static DigitalLinkValidationResult ValidateWithEngine(
        string gs1Data,
        IGs1DigitalLinkSyntaxEngine engine,
        string source,
        string defaultInvalidDetail,
        Action<string> validate)
    {
        try
        {
            validate(gs1Data);
            return new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = source,
                EngineVersion = EngineVersion,
                Detail = BuildValidatedDetail(engine.ParsedAiData),
            };
        }
        catch (GS1EncoderParameterException ex)
        {
            return Invalid(ex.Message, source, defaultInvalidDetail);
        }
        catch (GS1EncoderDigitalLinkException ex)
        {
            return Invalid(ex.Message, source, defaultInvalidDetail);
        }
        catch (GS1EncoderScanDataException ex)
        {
            return Invalid(ex.Message, source, defaultInvalidDetail);
        }
        catch (GS1EncoderGeneralException ex)
        {
            return Unavailable(ex.Message, source);
        }
        catch (DllNotFoundException ex)
        {
            return Unavailable(ex.Message, source);
        }
        catch (BadImageFormatException ex)
        {
            return Unavailable(ex.Message, source);
        }
        catch (EntryPointNotFoundException ex)
        {
            return Unavailable(ex.Message, source);
        }
        catch (FileNotFoundException ex)
        {
            return Unavailable(ex.Message, source);
        }
    }

    private static string BuildValidatedDetail(string? parsedAiData)
        => string.IsNullOrWhiteSpace(parsedAiData)
            ? "Validated with the official GS1 Syntax Engine."
            : $"Parsed GS1 AI data: {parsedAiData} Validated with the official GS1 Syntax Engine.";

    private static bool IsHttpDigitalLinkUri(string? value)
        => Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) &&
           (uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase));

    private static DigitalLinkValidationResult Invalid(
        string? reason,
        string source,
        string defaultDetail)
        => new()
        {
            Status = DigitalLinkValidationStatus.Invalid,
            Source = source,
            EngineVersion = EngineVersion,
            Detail = string.IsNullOrWhiteSpace(reason)
                ? defaultDetail
                : reason,
        };

    private static DigitalLinkValidationResult Unavailable(string? reason, string source)
        => new()
        {
            Status = DigitalLinkValidationStatus.Unavailable,
            Source = source,
            Detail = string.IsNullOrWhiteSpace(reason)
                ? "The GS1 Syntax Engine runtime is unavailable."
                : $"The GS1 Syntax Engine runtime is unavailable: {reason}",
        };
}

/// <summary>Small seam so validation outcome tests do not depend on a native DLL.</summary>
public interface IGs1DigitalLinkSyntaxEngine
{
    void Validate(string gs1Data);

    /// <summary>
    /// Canonical GS1 AI data produced by the official engine after validation,
    /// when that data is available.
    /// </summary>
    string? ParsedAiData => null;

    /// <summary>
    /// Validates a GS1 Element String. Test doubles can use the same simple
    /// implementation as their Digital Link validation.
    /// </summary>
    void ValidateElementString(string elementString) => Validate(elementString);
}

internal sealed class Gs1DigitalLinkSyntaxEngine : IGs1DigitalLinkSyntaxEngine
{
    private string? _parsedAiData;

    public string? ParsedAiData => _parsedAiData;

    public void Validate(string gs1Data)
    {
        _parsedAiData = null;
        using var encoder = new GS1Encoder();
        encoder.DataStr = gs1Data;
        _parsedAiData = encoder.AIdataStr;
    }

    public void ValidateElementString(string elementString)
    {
        string value = elementString.Trim();
        _parsedAiData = null;
        using var encoder = new GS1Encoder();

        if (value.StartsWith("<F1>", StringComparison.OrdinalIgnoreCase))
        {
            string rawData = value[4..].Replace("<F1>", "\x1D",
                StringComparison.OrdinalIgnoreCase);
            encoder.ScanData = "]d2" + rawData;
        }
        else if (value.StartsWith("]d2", StringComparison.OrdinalIgnoreCase))
        {
            string scanData = value.Replace("<F1>", "\x1D",
                StringComparison.OrdinalIgnoreCase);
            encoder.ScanData = scanData;
        }
        else
        {
            encoder.AIdataStr = value;
        }

        _parsedAiData = encoder.AIdataStr;
    }
}