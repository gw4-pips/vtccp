using System.Net;
using Azure.Data.Tables;
using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace Vccs.GcpUpdateService;

/// <summary>
/// Azure-gated GCP prefix table distribution endpoints.
///
///   GET /api/gcpMeta   → gcpMeta.json (tiny; used by clients for version checks)
///   GET /api/gcpTable  → streams gcpPrefixFormatList.xml.enc (AES-256-GCM envelope)
///
/// Both require a pre-shared device token in the X-Device-Token header; valid
/// tokens are configured in the DeviceTokens app setting (semicolon-separated).
/// Every request is audit-logged to the GcpAudit table (device token, IP,
/// endpoint, table version, timestamp).
///
/// Contract documented in vtccp/architecture/gcp-update-service.md.
/// </summary>
public sealed class GcpFunctions
{
    private const string TokenHeader = "X-Device-Token";
    private const string MetaBlobName = "gcpMeta.json";
    private const string TableBlobName = "gcpPrefixFormatList.xml.enc";

    private readonly ILogger<GcpFunctions> _log;

    public GcpFunctions(ILogger<GcpFunctions> log) => _log = log;

    [Function("gcpMeta")]
    public Task<HttpResponseData> GetMeta(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "gcpMeta")] HttpRequestData req)
        => ServeBlobAsync(req, MetaBlobName, "application/json", endpoint: "gcpMeta");

    [Function("gcpTable")]
    public Task<HttpResponseData> GetTable(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "gcpTable")] HttpRequestData req)
        => ServeBlobAsync(req, TableBlobName, "application/octet-stream", endpoint: "gcpTable");

    // ── Core ──────────────────────────────────────────────────────────────────

    private async Task<HttpResponseData> ServeBlobAsync(
        HttpRequestData req, string blobName, string contentType, string endpoint)
    {
        // 1. Device token check
        string? token = req.Headers.TryGetValues(TokenHeader, out var vals)
            ? vals.FirstOrDefault()?.Trim()
            : null;

        if (string.IsNullOrEmpty(token) || !IsValidToken(token))
        {
            _log.LogWarning("Rejected {Endpoint} request — missing/invalid device token.", endpoint);
            var unauthorized = req.CreateResponse(HttpStatusCode.Unauthorized);
            await unauthorized.WriteStringAsync("Invalid or missing device token.");
            return unauthorized;
        }

        // 2. Fetch blob
        string connection = GetSetting("GcpStorageConnection");
        string container  = Environment.GetEnvironmentVariable("GcpBlobContainer") ?? "gcp-prefix-table";
        var blob = new BlobClient(connection, container, blobName);

        if (!await blob.ExistsAsync())
        {
            var notFound = req.CreateResponse(HttpStatusCode.NotFound);
            await notFound.WriteStringAsync($"Blob '{blobName}' not found in container '{container}'.");
            return notFound;
        }

        var download = await blob.DownloadContentAsync();
        byte[] payload = download.Value.Content.ToArray();

        // 3. Audit — best-effort; never blocks delivery
        string version = await TryReadVersionAsync(connection, container);
        await TryAuditAsync(token, req, endpoint, version);

        // 4. Respond
        var response = req.CreateResponse(HttpStatusCode.OK);
        response.Headers.Add("Content-Type", contentType);
        await response.WriteBytesAsync(payload);
        return response;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static bool IsValidToken(string token)
    {
        string? configured = Environment.GetEnvironmentVariable("DeviceTokens");
        if (string.IsNullOrWhiteSpace(configured)) return false;
        return configured
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Any(t => string.Equals(t, token, StringComparison.Ordinal));
    }

    private static string GetSetting(string name) =>
        Environment.GetEnvironmentVariable(name)
        ?? throw new InvalidOperationException($"App setting '{name}' is not configured.");

    /// <summary>Reads the current table date from gcpMeta.json for the audit row.</summary>
    private static async Task<string> TryReadVersionAsync(string connection, string container)
    {
        try
        {
            var metaBlob = new BlobClient(connection, container, MetaBlobName);
            if (!await metaBlob.ExistsAsync()) return "unknown";
            var content = await metaBlob.DownloadContentAsync();
            using var doc = System.Text.Json.JsonDocument.Parse(content.Value.Content.ToString());
            return doc.RootElement.TryGetProperty("date", out var d) ? d.GetString() ?? "unknown" : "unknown";
        }
        catch { return "unknown"; }
    }

    private async Task TryAuditAsync(string token, HttpRequestData req, string endpoint, string version)
    {
        try
        {
            string connection = GetSetting("GcpStorageConnection");
            string tableName  = Environment.GetEnvironmentVariable("GcpAuditTable") ?? "GcpAudit";
            var table = new TableClient(connection, tableName);
            await table.CreateIfNotExistsAsync();

            string ip = req.Headers.TryGetValues("X-Forwarded-For", out var fwd)
                ? fwd.First().Split(',')[0].Trim()
                : "unknown";

            var entity = new TableEntity(
                partitionKey: DateTime.UtcNow.ToString("yyyy-MM-dd"),
                rowKey: $"{DateTime.UtcNow:HHmmss.fffffff}-{Guid.NewGuid():N}")
            {
                ["DeviceToken"] = token,
                ["Endpoint"]    = endpoint,
                ["Ip"]          = ip,
                ["Version"]     = version,
                ["TimestampUtc"] = DateTime.UtcNow.ToString("O"),
            };
            await table.AddEntityAsync(entity);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Audit write failed (non-fatal).");
        }
    }
}
