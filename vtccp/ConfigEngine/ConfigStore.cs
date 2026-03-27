namespace ConfigEngine;

using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// Low-level JSON persistence helper.  Reads/writes files under the VTCCP
/// application-data directory:
///   Windows: %APPDATA%\VTCCP\
///   Other:   ~/.vtccp/
///
/// All I/O is done asynchronously; the caller must provide the filename
/// (leaf name only, e.g. "devices.json"). The directory is created on first use.
/// </summary>
public static class ConfigStore
{
    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        WriteIndented         = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters            = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };

    // ── Directory ─────────────────────────────────────────────────────────────

    private static string? _dataDirectory;

    /// <summary>
    /// Root directory where all VTCCP config files are stored.
    /// Defaults to the platform-appropriate app-data location.
    /// Override via <see cref="SetDataDirectory"/> (e.g. for testing).
    /// </summary>
    public static string DataDirectory
    {
        get
        {
            if (_dataDirectory is not null) return _dataDirectory;
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (string.IsNullOrEmpty(appData))
                appData = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return _dataDirectory = Path.Combine(appData, "VTCCP");
        }
    }

    /// <summary>Overrides the data directory (useful for tests or portable mode).</summary>
    public static void SetDataDirectory(string path)
    {
        _dataDirectory = path ?? throw new ArgumentNullException(nameof(path));
    }

    // ── Core I/O ──────────────────────────────────────────────────────────────

    /// <summary>Returns the full path for <paramref name="filename"/> inside the data directory.</summary>
    public static string PathFor(string filename) => Path.Combine(DataDirectory, filename);

    /// <summary>True if the file exists in the data directory.</summary>
    public static bool Exists(string filename) => File.Exists(PathFor(filename));

    /// <summary>
    /// Serialises <paramref name="value"/> to JSON and writes it atomically
    /// (write to temp file, then rename).
    /// </summary>
    public static async Task SaveAsync<T>(string filename, T value, CancellationToken ct = default)
    {
        Directory.CreateDirectory(DataDirectory);
        string target = PathFor(filename);
        string tmp    = target + ".tmp";

        await using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None,
                                             bufferSize: 4096, useAsync: true))
        {
            await JsonSerializer.SerializeAsync(fs, value, _jsonOpts, ct);
        }

        File.Move(tmp, target, overwrite: true);
    }

    /// <summary>
    /// Reads and deserialises the file, or returns <paramref name="defaultValue"/>
    /// if the file does not exist.
    /// </summary>
    public static async Task<T> LoadAsync<T>(string filename, T defaultValue, CancellationToken ct = default)
    {
        string path = PathFor(filename);
        if (!File.Exists(path)) return defaultValue;

        await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                                            bufferSize: 4096, useAsync: true);
        T? result = await JsonSerializer.DeserializeAsync<T>(fs, _jsonOpts, ct);
        return result ?? defaultValue;
    }

    /// <summary>Deletes the file from the data directory (no-op if absent).</summary>
    public static void Delete(string filename)
    {
        string path = PathFor(filename);
        if (File.Exists(path)) File.Delete(path);
    }
}
