using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeviceInterface.Rfid.FlexWedge;

/// <summary>Portable, UI-free configuration for a FlexWedge deployment.</summary>
public sealed record FlexWedgeConfiguration
{
    public string ConfigurationVersion { get; init; } = "1";
    public string? ReaderName { get; init; }
    public string? ReaderPort { get; init; }
    public string? CustomGcpXmlPath { get; init; }
    public string ActiveProfile { get; init; } = "Raw EPC";
    public IReadOnlyList<FlexWedgeOutputProfile> OutputProfiles { get; init; } =
        [FlexWedgeOutputProfile.RawEpc, FlexWedgeOutputProfile.GtinAndSerial, FlexWedgeOutputProfile.DelimitedMapped];

    /// <summary>A ready-to-save baseline for a single supported reader and focused-cell/text sink.</summary>
    public static FlexWedgeConfiguration SampleReadyDefault { get; } = new()
    {
        ReaderName = "ASR-P35U",
        ActiveProfile = "GTIN + serial",
        OutputProfiles =
        [
            FlexWedgeOutputProfile.RawEpc,
            FlexWedgeOutputProfile.GtinAndSerial with { AppendKey = FlexWedgeAppendKey.Tab },
            FlexWedgeOutputProfile.DelimitedMapped,
            FlexWedgeOutputProfile.CompleteDecoded
        ]
    };

    public IReadOnlyList<string> Validate()
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(ConfigurationVersion)) errors.Add("Configuration version is required.");
        if (OutputProfiles.Count == 0) errors.Add("At least one output profile is required.");
        if (OutputProfiles.GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase).Any(x => x.Count() > 1))
            errors.Add("Output profile names must be unique.");
        if (!OutputProfiles.Any(x => string.Equals(x.Name, ActiveProfile, StringComparison.OrdinalIgnoreCase)))
            errors.Add("Active profile must name an output profile.");
        foreach (var profile in OutputProfiles)
            errors.AddRange(profile.Validate().Select(error => $"Profile '{profile.Name}': {error}"));
        return errors;
    }
}

public static class FlexWedgeConfigurationJson
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    public static FlexWedgeConfiguration Load(string json)
    {
        if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("Configuration JSON is required.", nameof(json));
        var configuration = JsonSerializer.Deserialize<FlexWedgeConfiguration>(json, Options)
            ?? throw new InvalidDataException("Configuration JSON did not contain a configuration object.");
        var errors = configuration.Validate();
        if (errors.Count != 0) throw new InvalidDataException(string.Join(" ", errors));
        return configuration;
    }

    public static string Save(FlexWedgeConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        var errors = configuration.Validate();
        if (errors.Count != 0) throw new InvalidDataException(string.Join(" ", errors));
        return JsonSerializer.Serialize(configuration, Options);
    }

    public static async Task<FlexWedgeConfiguration> LoadAsync(Stream input, CancellationToken cancellationToken = default)
    {
        using var reader = new StreamReader(input ?? throw new ArgumentNullException(nameof(input)), leaveOpen: true);
        return Load(await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false));
    }

    public static async Task SaveAsync(Stream output, FlexWedgeConfiguration configuration, CancellationToken cancellationToken = default)
    {
        using var writer = new StreamWriter(output ?? throw new ArgumentNullException(nameof(output)), leaveOpen: true);
        await writer.WriteAsync(Save(configuration).AsMemory(), cancellationToken).ConfigureAwait(false);
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
    }
}