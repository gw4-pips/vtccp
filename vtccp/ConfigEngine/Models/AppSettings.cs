namespace ConfigEngine.Models;

/// <summary>
/// Application-wide preferences persisted alongside device profiles and job templates.
/// </summary>
public sealed class AppSettings
{
    // ── Defaults ──────────────────────────────────────────────────────────────

    /// <summary>Id of the <see cref="DeviceProfile"/> selected at last exit, or null.</summary>
    public string? LastDeviceProfileId { get; set; }

    /// <summary>Id of the <see cref="JobTemplate"/> selected at last exit, or null.</summary>
    public string? LastJobTemplateId { get; set; }

    // ── Output ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Root output directory for all sessions.
    /// Defaults to <c>%USERPROFILE%\Documents\VTCCP</c> on Windows,
    /// <c>~/VTCCP</c> on other platforms.
    /// </summary>
    public string DefaultOutputDirectory { get; set; } = GetDefaultOutputDirectory();

    // ── UI ────────────────────────────────────────────────────────────────────

    /// <summary>"Light" or "Dark". Applies to the WPF resource dictionary swap.</summary>
    public string Theme { get; set; } = "Light";

    /// <summary>MainWindow width remembered across sessions.</summary>
    public double WindowWidth { get; set; } = 1024;

    /// <summary>MainWindow height remembered across sessions.</summary>
    public double WindowHeight { get; set; } = 680;

    // ── Operator quick-access ─────────────────────────────────────────────────

    /// <summary>Operator ID typed at last session start (pre-fills the session launcher).</summary>
    public string? LastOperatorId { get; set; }

    // ── Schema version awareness ──────────────────────────────────────────────

    /// <summary>App version that wrote this settings file.</summary>
    public string AppVersion { get; set; } = "1.0.0";

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string GetDefaultOutputDirectory()
    {
        string home = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (string.IsNullOrEmpty(home))
            home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return Path.Combine(home, "VTCCP");
    }
}
