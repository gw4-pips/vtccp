namespace DeviceInterface;

using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// Connection and behaviour configuration for a single Cognex DataMan DMV device.
/// Serializable to/from JSON for persistence in the VTCCP application settings.
/// </summary>
public sealed class DeviceConfig
{
    // ── Network ───────────────────────────────────────────────────────────────

    /// <summary>Device IP address or hostname.</summary>
    public string Host { get; set; } = "192.168.0.1";

    /// <summary>DMCC TCP port. DataMan DMV default = 44444.</summary>
    public int Port { get; set; } = 44444;

    /// <summary>Milliseconds to wait for a TCP connection to be established.</summary>
    public int ConnectTimeoutMs { get; set; } = 3_000;

    /// <summary>
    /// Milliseconds to wait for the first byte of a DMCC response before giving up.
    /// Should be longer than the device's longest processing time.
    /// </summary>
    public int ResponseTimeoutMs { get; set; } = 5_000;

    /// <summary>
    /// Milliseconds of inactivity (no new bytes) that signals the response is complete.
    /// Used to detect end-of-response for commands that don't have a fixed terminator.
    /// </summary>
    public int IdleGapMs { get; set; } = 150;

    /// <summary>
    /// Maximum milliseconds to wait for the device's TCP welcome banner immediately
    /// after connecting.  Many devices send nothing; 500 ms is enough to capture it
    /// without stalling session start for 5+ seconds.
    /// </summary>
    public int BannerTimeoutMs { get; set; } = 500;

    // ── Result delivery ───────────────────────────────────────────────────────

    /// <summary>
    /// How the VTCCP client receives scan results from the device.
    /// Poll = host sends TRIGGER then polls GET SYMBOL.RESULT.
    /// Push = device pushes DMST XML to a listener socket on the host.
    /// </summary>
    public ResultDeliveryMode ResultMode { get; set; } = ResultDeliveryMode.Poll;

    /// <summary>
    /// TCP port on the host that the device connects to in Push (DMST) mode.
    /// Not used in Poll mode.
    /// </summary>
    public int DmstListenPort { get; set; } = 9004;

    // ── Retry / reconnect ─────────────────────────────────────────────────────

    /// <summary>Number of automatic reconnection attempts before surfacing an error.</summary>
    public int MaxReconnectAttempts { get; set; } = 3;

    /// <summary>Milliseconds to wait between reconnection attempts.</summary>
    public int ReconnectDelayMs { get; set; } = 1_000;

    // ── Serialization helpers ─────────────────────────────────────────────────

    private static readonly JsonSerializerOptions _json = new()
    {
        WriteIndented         = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters            = { new JsonStringEnumConverter() },
    };

    public string ToJson()             => JsonSerializer.Serialize(this, _json);
    public static DeviceConfig FromJson(string json) =>
        JsonSerializer.Deserialize<DeviceConfig>(json, _json) ?? new DeviceConfig();
}

/// <summary>How verification results are delivered to the VTCCP client.</summary>
public enum ResultDeliveryMode
{
    /// <summary>
    /// VTCCP sends TRIGGER, waits, then polls GET SYMBOL.RESULT for the XML response.
    /// Simpler; works on all firmware without additional device configuration.
    /// </summary>
    Poll,

    /// <summary>
    /// Device pushes DMST XML results to a TCP listener socket on the host after each scan.
    /// Lower latency; requires DMST output to be configured on the device.
    /// </summary>
    Push,
}
