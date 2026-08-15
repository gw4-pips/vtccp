namespace VtccpApp.Services;

using ConfigEngine.Models;
using DeviceInterface.Rfid.Gcp;

/// <summary>
/// Builds a configured <see cref="GcpUpdateService"/> from <see cref="AppSettings"/>.
/// Shared by the startup background check (MainViewModel) and the manual
/// "Check for update" button on the Settings page.
/// </summary>
public static class GcpUpdateServiceFactory
{
    /// <summary>
    /// Default install target for downloaded tables: %AppData%\VCCS\gcpPrefixes.xml.
    /// </summary>
    public static string DefaultInstallPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "VCCS", "gcpPrefixes.xml");

    /// <summary>
    /// Resolves the local table path used for date comparison and install:
    /// the configured <see cref="AppSettings.GcpDataPath"/> when set, otherwise
    /// the %AppData% default.
    /// </summary>
    public static string ResolveLocalXmlPath(AppSettings settings) =>
        string.IsNullOrWhiteSpace(settings.GcpDataPath)
            ? DefaultInstallPath
            : settings.GcpDataPath!;

    /// <summary>
    /// Creates the update client, or null when the service URL / device token
    /// are not configured (auto-update disabled).
    /// </summary>
    public static GcpUpdateService? Create(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.GcpUpdateServiceUrl) ||
            string.IsNullOrWhiteSpace(settings.GcpDeviceToken))
            return null;

        return new GcpUpdateService(new GcpUpdateOptions
        {
            ServiceUrl   = settings.GcpUpdateServiceUrl!.Trim(),
            DeviceToken  = settings.GcpDeviceToken!.Trim(),
            LocalXmlPath = ResolveLocalXmlPath(settings),
            // gcpKey.bin beside the EXE (GcpUpdateService default when null)
            KeyPath      = null,
        });
    }
}
