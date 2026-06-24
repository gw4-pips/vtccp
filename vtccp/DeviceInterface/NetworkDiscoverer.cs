namespace DeviceInterface;

using System.Net;
using CognexSdk = Cognex.DataMan.SDK;

/// <summary>
/// A DataMan device found on the local subnet by <see cref="NetworkDiscoverer"/>.
/// All fields are populated from the UDP broadcast reply; some may be empty if the
/// device firmware does not include them in the discovery response.
/// </summary>
public sealed record DiscoveredDevice(
    string Name,
    string Host,
    int    Port,
    string DeviceType,
    string FirmwareVersion,
    string Serial,
    string MacAddress);

/// <summary>
/// Wraps the Cognex DataMan SDK <c>EthSystemDiscoverer</c> to locate DataMan
/// devices on the local subnet via UDP broadcast (same mechanism as DMST).
/// </summary>
public static class NetworkDiscoverer
{
    /// <summary>
    /// Broadcasts a UDP discovery probe and collects responding DataMan devices.
    /// Listens for <paramref name="listenMs"/> milliseconds (default 3 000 ms).
    /// Returns a deduplicated, read-only list ordered by response arrival time.
    /// </summary>
    public static async Task<IReadOnlyList<DiscoveredDevice>> DiscoverAsync(
        int               listenMs = 3_000,
        CancellationToken ct       = default)
    {
        var results = new List<DiscoveredDevice>();
        var seen    = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var gate    = new object();

        var discoverer = new CognexSdk.EthSystemDiscoverer();

        discoverer.SystemDiscovered += (_, e) =>
        {
            try
            {
                var connector = e.Connector;
                string host   = connector.Address?.ToString() ?? string.Empty;
                if (string.IsNullOrEmpty(host)) return;

                lock (gate)
                {
                    if (!seen.Add(host)) return;
                    results.Add(new DiscoveredDevice(
                        Name:            connector.Name            ?? host,
                        Host:            host,
                        Port:            44_444,
                        DeviceType:      connector.DeviceType      ?? string.Empty,
                        FirmwareVersion: connector.FirmwareVersion ?? string.Empty,
                        Serial:          connector.SerialNumber    ?? string.Empty,
                        MacAddress:      connector.MACAddress      ?? string.Empty));
                }
            }
            catch { /* ignore malformed discovery responses */ }
        };

        await Task.Run(() =>
        {
            discoverer.StartDiscovery();
            System.Threading.Thread.Sleep(listenMs);
            discoverer.StopDiscovery();
        }, ct).ConfigureAwait(false);

        lock (gate)
            return results.AsReadOnly();
    }
}
