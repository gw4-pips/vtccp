namespace DeviceInterface;

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
/// Locates DataMan devices on the local subnet via UDP broadcast (same mechanism
/// as DMST).
///
/// ⚠ SDK DISCOVERY STUB — 2026-06-24
/// The Cognex DataMan SDK v25.4.1 DLL (Cognex.DataMan.SDK.PC.dll) does NOT expose
/// <c>EthSystemDiscoverer</c> at <c>Cognex.DataMan.SDK.EthSystemDiscoverer</c>.
/// <c>EthSystemConnector</c> and <c>DataManSystem</c> are present and working.
/// The correct discovery class name for this SDK version is unconfirmed.
/// This method currently returns an empty list — operators can add devices manually
/// via the ⊕ Import button in the Devices panel.
/// TODO: Confirm the correct SDK discovery class name from Cognex SDK docs / DLL
///       reflection, then restore the SDK-based implementation.
/// </summary>
public static class NetworkDiscoverer
{
    /// <summary>
    /// Broadcasts a UDP discovery probe and collects responding DataMan devices.
    /// Listens for <paramref name="listenMs"/> milliseconds (default 3 000 ms).
    /// Returns a deduplicated, read-only list ordered by response arrival time.
    ///
    /// Currently stubbed — returns empty list. See class-level XML doc.
    /// </summary>
    public static Task<IReadOnlyList<DiscoveredDevice>> DiscoverAsync(
        int               listenMs = 3_000,
        CancellationToken ct       = default)
    {
        System.Diagnostics.Debug.WriteLine(
            "[VTCCP-NET] NetworkDiscoverer: SDK discovery stubbed — " +
            "EthSystemDiscoverer class name unconfirmed for SDK v25.4.1. " +
            "Manual IP entry via ⊕ Import is available.");

        return Task.FromResult<IReadOnlyList<DiscoveredDevice>>(
            new List<DiscoveredDevice>().AsReadOnly());
    }
}
