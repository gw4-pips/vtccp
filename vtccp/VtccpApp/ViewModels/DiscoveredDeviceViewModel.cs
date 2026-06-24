namespace VtccpApp.ViewModels;

using DeviceInterface;

/// <summary>
/// Thin view-model wrapper around a <see cref="DiscoveredDevice"/> for display in
/// the scan-results panel.  All fields are immutable after construction.
/// </summary>
public sealed class DiscoveredDeviceViewModel
{
    public string Name            { get; }
    public string Host            { get; }
    public int    Port            { get; }
    public string DeviceType      { get; }
    public string FirmwareVersion { get; }
    public string Serial          { get; }
    public string MacAddress      { get; }

    public DiscoveredDeviceViewModel(DiscoveredDevice d)
    {
        Name            = d.Name;
        Host            = d.Host;
        Port            = d.Port;
        DeviceType      = d.DeviceType;
        FirmwareVersion = d.FirmwareVersion;
        Serial          = d.Serial;
        MacAddress      = d.MacAddress;
    }

    public string Summary =>
        string.Join("  ·  ",
            new[] { Name, Host, DeviceType, $"FW {FirmwareVersion}", $"S/N {Serial}" }
            .Where(s => !string.IsNullOrWhiteSpace(s)));
}
