namespace DeviceInterface.Rfid.Gcp;

/// <summary>
/// Checks for and downloads updated GS1 GCP Prefix Format List XML files.
///
/// Update endpoint (confirmed in scope doc Rev 1.1):
///   https://my2dir-resolver-bwa7agd0ctehbqf3.eastus2-01.azurewebsites.net/tools/gcp/interop/current.xml
///   Header: X-GCP-Interop-Key = {GCP_INTEROP_KEY secret}
///
/// Strategy: HEAD request → compare Last-Modified vs stored file date attribute.
/// If key is absent or empty, all operations are silently skipped.
/// </summary>
public sealed class GcpUpdateService
{
    private const string EndpointUrl =
        "https://my2dir-resolver-bwa7agd0ctehbqf3.eastus2-01.azurewebsites.net/tools/gcp/interop/current.xml";
    private const string KeyHeader = "X-GCP-Interop-Key";
    private const string KeyEnvVar = "GCP_INTEROP_KEY";

    private readonly HttpClient _http;
    private readonly string _localXmlPath;

    /// <param name="localXmlPath">
    /// Full path to the local gcp-prefix-format-list.xml file.
    /// Used both as the update target and to read the stored date for comparison.
    /// </param>
    public GcpUpdateService(string localXmlPath, HttpClient? http = null)
    {
        _localXmlPath = localXmlPath ?? throw new ArgumentNullException(nameof(localXmlPath));
        _http = http ?? new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
    }

    /// <summary>
    /// Check whether a newer GCP list is available on the server without downloading it.
    /// Returns null when the key is absent, the network is unreachable, or the server
    /// does not return a Last-Modified header.
    /// Returns true when the server file is newer than the local file.
    /// Returns false when the local file is already current.
    /// </summary>
    public async Task<bool?> IsUpdateAvailableAsync(CancellationToken ct = default)
    {
        string? key = GetKey();
        if (key is null) return null;

        DateTimeOffset? serverDate;
        try { serverDate = await GetServerLastModifiedAsync(key, ct).ConfigureAwait(false); }
        catch { return null; }
        if (serverDate is null) return null;

        DateTimeOffset? localDate = GetLocalFileDate();
        if (localDate is null) return true; // no local file → update available

        return serverDate.Value > localDate.Value;
    }

    /// <summary>
    /// Download the current GCP list from the server and save it to <see cref="_localXmlPath"/>.
    /// Throws on network or disk errors.
    /// No-op (returns null) when the key is absent.
    /// </summary>
    /// <returns>
    /// The "date" attribute value parsed from the downloaded XML root element,
    /// or null if the key is absent or the attribute is missing/unparseable.
    /// Store the returned value in <see cref="ConfigEngine.Models.AppSettings.GcpLastModified"/>
    /// so the provenance annotation in PDF reports reflects the newly downloaded table.
    /// </returns>
    public async Task<DateTimeOffset?> DownloadUpdateAsync(CancellationToken ct = default)
    {
        string? key = GetKey();
        if (key is null) return null;

        using var request = new HttpRequestMessage(HttpMethod.Get, EndpointUrl);
        request.Headers.TryAddWithoutValidation(KeyHeader, key);

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct)
                                        .ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        string dir = Path.GetDirectoryName(_localXmlPath) ?? ".";
        Directory.CreateDirectory(dir);
        string tmp = _localXmlPath + ".tmp";

        await using (var fs = File.Create(tmp))
        await using (var content = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false))
            await content.CopyToAsync(fs, ct).ConfigureAwait(false);

        File.Move(tmp, _localXmlPath, overwrite: true);

        // Read the date from the newly written file and return it so the caller
        // can persist it to AppSettings.GcpLastModified.
        return GetLocalFileDate();
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private static string? GetKey()
    {
        string? key = Environment.GetEnvironmentVariable(KeyEnvVar);
        return string.IsNullOrWhiteSpace(key) ? null : key.Trim();
    }

    private async Task<DateTimeOffset?> GetServerLastModifiedAsync(string key, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Head, EndpointUrl);
        request.Headers.TryAddWithoutValidation(KeyHeader, key);
        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct)
                                        .ConfigureAwait(false);
        if (!response.IsSuccessStatusCode) return null;
        return response.Content.Headers.LastModified;
    }

    private DateTimeOffset? GetLocalFileDate()
    {
        if (!File.Exists(_localXmlPath)) return null;
        try
        {
            // Parse the date attribute from the root element (fast, avoids loading full 8MB file)
            using var stream = File.OpenRead(_localXmlPath);
            using var reader = System.Xml.XmlReader.Create(stream,
                new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });

            while (reader.Read())
            {
                if (reader.NodeType != System.Xml.XmlNodeType.Element) continue;
                string? dateStr = reader.GetAttribute("date");
                if (dateStr is null) return null;
                return DateTimeOffset.TryParse(dateStr, out var d) ? d : null;
            }
        }
        catch { /* best-effort */ }
        return null;
    }
}
