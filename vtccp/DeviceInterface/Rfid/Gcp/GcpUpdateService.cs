using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeviceInterface.Rfid.Gcp;

/// <summary>
/// Metadata document served by the update service (<c>GET /api/gcpMeta</c>).
/// Mirrors <c>gcpMeta.json</c> in the Azure blob container.
/// </summary>
public sealed record GcpMeta
{
    /// <summary>Table date — the "date" attribute of the GCPPrefixFormatList root.</summary>
    [JsonPropertyName("date")]
    public string Date { get; init; } = string.Empty;

    /// <summary>Lower-case hex SHA-256 of the encrypted blob (<c>.enc</c> file).</summary>
    [JsonPropertyName("sha256")]
    public string Sha256 { get; init; } = string.Empty;
}

/// <summary>Result of a version check against the update service.</summary>
public sealed record GcpUpdateCheckResult
{
    public required GcpMeta ServerMeta { get; init; }
    public DateTimeOffset? ServerDate { get; init; }
    public DateTimeOffset? LocalDate { get; init; }

    /// <summary>True when the server table is newer than the local one (or no local table exists).</summary>
    public bool UpdateAvailable =>
        ServerDate is not null && (LocalDate is null || ServerDate > LocalDate);
}

/// <summary>Configuration for <see cref="GcpUpdateService"/>.</summary>
public sealed record GcpUpdateOptions
{
    /// <summary>Base URL of the Azure Function app, e.g. <c>https://vccs-gcp-update.azurewebsites.net</c>.</summary>
    public required string ServiceUrl { get; init; }

    /// <summary>Pre-shared device token sent as <c>X-Device-Token</c> on every request.</summary>
    public required string DeviceToken { get; init; }

    /// <summary>Full path of the local decrypted GCP prefix XML (install target).</summary>
    public required string LocalXmlPath { get; init; }

    /// <summary>Path to <c>gcpKey.bin</c> (32 raw bytes). Defaults to beside the EXE.</summary>
    public string? KeyPath { get; init; }
}

/// <summary>
/// Client for the Azure-gated GCP prefix table update service
/// (see <c>vtccp/architecture/gcp-update-service.md</c>).
///
/// Endpoints (device token in <c>X-Device-Token</c> header):
///   GET {ServiceUrl}/api/gcpMeta   → gcpMeta.json  { "date": "...", "sha256": "..." }
///   GET {ServiceUrl}/api/gcpTable  → AES-256-GCM encrypted table blob (GCP1 envelope)
///
/// All check operations are best-effort and never throw; a null result means
/// "not configured or unreachable" and callers should behave as if no update exists.
/// <see cref="DownloadAndInstallAsync"/> throws on failure so the UI can surface it.
/// </summary>
public sealed class GcpUpdateService
{
    private const string TokenHeader = "X-Device-Token";

    private readonly GcpUpdateOptions _options;
    private readonly HttpClient _http;

    public GcpUpdateService(GcpUpdateOptions options, HttpClient? http = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _http = http ?? new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
    }

    /// <summary>True when a service URL and device token are configured.</summary>
    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(_options.ServiceUrl) &&
        !string.IsNullOrWhiteSpace(_options.DeviceToken);

    // ── Check ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// Fetches server metadata and compares against the local table date.
    /// Returns null when not configured, the network is unreachable, or the
    /// response is malformed. Never throws.
    /// </summary>
    public async Task<GcpUpdateCheckResult?> CheckNowAsync(CancellationToken ct = default)
    {
        if (!IsConfigured) return null;

        try
        {
            using var request = BuildRequest("api/gcpMeta");
            using var response = await _http.SendAsync(request, ct).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) return null;

            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            var meta = await JsonSerializer.DeserializeAsync<GcpMeta>(stream, cancellationToken: ct)
                                           .ConfigureAwait(false);
            if (meta is null || string.IsNullOrWhiteSpace(meta.Date)) return null;

            DateTimeOffset? serverDate =
                DateTimeOffset.TryParse(meta.Date, out var sd) ? sd : null;

            return new GcpUpdateCheckResult
            {
                ServerMeta = meta,
                ServerDate = serverDate,
                LocalDate  = GetLocalTableDate(),
            };
        }
        catch
        {
            return null; // best-effort — offline workstations must start normally
        }
    }

    // ── Download + install ────────────────────────────────────────────────────

    /// <summary>
    /// Downloads the encrypted table, verifies its SHA-256 against server metadata,
    /// decrypts it with <c>gcpKey.bin</c>, sanity-parses the XML, and atomically
    /// installs it at <see cref="GcpUpdateOptions.LocalXmlPath"/>.
    /// Throws on any failure (network, hash mismatch, bad key, malformed XML).
    /// </summary>
    /// <returns>The "date" attribute of the newly installed table, or null when unparseable.</returns>
    public async Task<DateTimeOffset?> DownloadAndInstallAsync(CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("GCP update service is not configured.");

        // 1. Fresh metadata (authoritative hash for the blob we are about to pull)
        var check = await CheckNowAsync(ct).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Update service unreachable.");

        // 2. Encrypted blob
        byte[] envelope;
        using (var request = BuildRequest("api/gcpTable"))
        using (var response = await _http.SendAsync(request, ct).ConfigureAwait(false))
        {
            response.EnsureSuccessStatusCode();
            envelope = await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false);
        }

        // 3. Integrity check — SHA-256 of the encrypted blob must match gcpMeta.json
        string actualHash = Convert.ToHexString(SHA256.HashData(envelope)).ToLowerInvariant();
        string expected   = check.ServerMeta.Sha256.Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(expected) && actualHash != expected)
            throw new InvalidDataException(
                $"GCP table hash mismatch — expected {expected}, got {actualHash}.");

        // 4. Decrypt (throws CryptographicException on wrong key / tampering)
        byte[] key = LoadKey();
        byte[] xmlBytes = GcpCrypto.Decrypt(envelope, key);

        // 5. Sanity parse — reject anything GcpLengthTable cannot load
        GcpLengthTable table;
        using (var ms = new MemoryStream(xmlBytes, writable: false))
            table = GcpLengthTable.LoadFromStream(ms);

        // 6. Atomic install
        string dir = Path.GetDirectoryName(_options.LocalXmlPath) ?? ".";
        Directory.CreateDirectory(dir);
        string tmp = _options.LocalXmlPath + ".tmp";
        await File.WriteAllBytesAsync(tmp, xmlBytes, ct).ConfigureAwait(false);
        File.Move(tmp, _options.LocalXmlPath, overwrite: true);

        System.Diagnostics.Debug.WriteLine(
            $"[GCP] Installed table ({table.EntryCount} entries, date={table.DataDate:yyyy-MM-dd}) → {_options.LocalXmlPath}");
        return table.DataDate;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private HttpRequestMessage BuildRequest(string relativePath)
    {
        string baseUrl = _options.ServiceUrl.TrimEnd('/');
        var request = new HttpRequestMessage(HttpMethod.Get, $"{baseUrl}/{relativePath}");
        request.Headers.TryAddWithoutValidation(TokenHeader, _options.DeviceToken.Trim());
        return request;
    }

    private byte[] LoadKey()
    {
        string keyPath = !string.IsNullOrWhiteSpace(_options.KeyPath)
            ? _options.KeyPath!
            : Path.Combine(AppContext.BaseDirectory, "gcpKey.bin");
        if (!File.Exists(keyPath))
            throw new FileNotFoundException(
                $"GCP decryption key not found: {keyPath}. Deploy gcpKey.bin beside the EXE.", keyPath);
        byte[] key = File.ReadAllBytes(keyPath);
        if (key.Length != GcpCrypto.KeyLength)
            throw new InvalidDataException(
                $"gcpKey.bin must be exactly {GcpCrypto.KeyLength} bytes (found {key.Length}).");
        return key;
    }

    /// <summary>
    /// Parses the "date" attribute from the local table's root element without
    /// loading the full ~10 MB document. Returns null when absent or unreadable.
    /// </summary>
    public DateTimeOffset? GetLocalTableDate()
    {
        if (!File.Exists(_options.LocalXmlPath)) return null;
        try
        {
            using var stream = File.OpenRead(_options.LocalXmlPath);
            using var reader = System.Xml.XmlReader.Create(stream,
                new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });

            while (reader.Read())
            {
                if (reader.NodeType != System.Xml.XmlNodeType.Element) continue;
                string? dateStr = reader.GetAttribute("date");
                return dateStr is not null && DateTimeOffset.TryParse(dateStr, out var d) ? d : null;
            }
        }
        catch { /* best-effort */ }
        return null;
    }
}
