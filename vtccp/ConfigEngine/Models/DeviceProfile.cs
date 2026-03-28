namespace ConfigEngine.Models;

using DeviceInterface;

/// <summary>
/// Named, serializable device connection profile. Mirrors <see cref="DeviceConfig"/>
/// with additional identity and display metadata for the GUI layer.
/// </summary>
public sealed class DeviceProfile
{
    // ── Identity ──────────────────────────────────────────────────────────────

    /// <summary>Stable unique identifier (GUID string). Set on creation.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>Operator-facing display name, e.g. "Line-1 DM260Q".</summary>
    public string Name { get; set; } = "New Device";

    /// <summary>Optional free-text notes.</summary>
    public string? Notes { get; set; }

    /// <summary>Whether this profile is the application-wide default.</summary>
    public bool IsDefault { get; set; }

    // ── Connection parameters ─────────────────────────────────────────────────

    /// <summary>Device hostname or IP address.</summary>
    public string Host { get; set; } = "192.168.0.100";

    /// <summary>DMCC TCP port. DataMan DMV default = 44444.</summary>
    public int Port { get; set; } = 44444;

    /// <summary>TCP connect timeout in milliseconds.</summary>
    public int ConnectTimeoutMs { get; set; } = 5_000;

    /// <summary>Overall command response timeout in milliseconds.</summary>
    public int ResponseTimeoutMs { get; set; } = 5_000;

    /// <summary>
    /// Idle-gap in milliseconds used to detect end-of-response.
    /// Should be 2–3× the expected inter-packet gap.
    /// </summary>
    public int IdleGapMs { get; set; } = 150;

    // ── Acquisition mode ──────────────────────────────────────────────────────

    /// <summary>DMST listener port for Push (unsolicited result) mode. 0 = Poll mode only.</summary>
    public int DmstListenPort { get; set; } = 0;

    // ── Derived helpers ───────────────────────────────────────────────────────

    /// <summary>Converts this profile to the <see cref="DeviceConfig"/> used by DeviceInterface.</summary>
    public DeviceConfig ToDeviceConfig() => new()
    {
        Host              = Host,
        Port              = Port,
        ConnectTimeoutMs  = ConnectTimeoutMs,
        ResponseTimeoutMs = ResponseTimeoutMs,
        IdleGapMs         = IdleGapMs,
        DmstListenPort    = DmstListenPort,
    };

    /// <summary>Populates this profile's connection fields from a <see cref="DeviceConfig"/>.</summary>
    public void ApplyDeviceConfig(DeviceConfig cfg)
    {
        Host              = cfg.Host;
        Port              = cfg.Port;
        ConnectTimeoutMs  = cfg.ConnectTimeoutMs;
        ResponseTimeoutMs = cfg.ResponseTimeoutMs;
        IdleGapMs         = cfg.IdleGapMs;
        DmstListenPort    = cfg.DmstListenPort;
    }

    public override string ToString() => $"{Name} ({Host}:{Port})";
}
