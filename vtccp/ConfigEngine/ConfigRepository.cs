namespace ConfigEngine;

using ConfigEngine.Models;

/// <summary>
/// High-level repository for device profiles, job templates, and application settings.
/// All collections are loaded from / saved to <see cref="ConfigStore"/> as JSON files.
///
/// Thread safety: single-threaded (call from the UI thread or under external lock).
/// </summary>
public sealed class ConfigRepository
{
    // ── File names ────────────────────────────────────────────────────────────

    private const string DevicesFile   = "devices.json";
    private const string TemplatesFile = "templates.json";
    private const string SettingsFile  = "appsettings.json";

    // ── State ─────────────────────────────────────────────────────────────────

    public List<DeviceProfile> Devices   { get; private set; } = [];
    public List<JobTemplate>   Templates { get; private set; } = [];
    public AppSettings         Settings  { get; private set; } = new();

    // ── Load / Save ───────────────────────────────────────────────────────────

    /// <summary>Loads all config files from disk. Missing files default to empty collections.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        Devices   = await ConfigStore.LoadAsync(DevicesFile,   new List<DeviceProfile>(), ct);
        Templates = await ConfigStore.LoadAsync(TemplatesFile, new List<JobTemplate>(),   ct);
        Settings  = await ConfigStore.LoadAsync(SettingsFile,  new AppSettings(),         ct);
    }

    /// <summary>Persists all collections to disk atomically (one file at a time).</summary>
    public async Task SaveAsync(CancellationToken ct = default)
    {
        await ConfigStore.SaveAsync(DevicesFile,   Devices,   ct);
        await ConfigStore.SaveAsync(TemplatesFile, Templates, ct);
        await ConfigStore.SaveAsync(SettingsFile,  Settings,  ct);
    }

    /// <summary>Persists only the settings file (avoids rewriting lists on UI preference change).</summary>
    public Task SaveSettingsAsync(CancellationToken ct = default) =>
        ConfigStore.SaveAsync(SettingsFile, Settings, ct);

    // ── Device profiles ───────────────────────────────────────────────────────

    /// <summary>Returns the profile with the given Id, or null.</summary>
    public DeviceProfile? FindDevice(string id) =>
        Devices.FirstOrDefault(d => d.Id == id);

    /// <summary>Returns the default profile, or null if none is marked default.</summary>
    public DeviceProfile? DefaultDevice =>
        Devices.FirstOrDefault(d => d.IsDefault) ?? Devices.FirstOrDefault();

    /// <summary>
    /// Adds <paramref name="profile"/> to the collection.
    /// If <paramref name="profile"/>.IsDefault is true, clears IsDefault on all others.
    /// </summary>
    public void AddDevice(DeviceProfile profile)
    {
        ArgumentNullException.ThrowIfNull(profile);
        if (profile.IsDefault) ClearDeviceDefaults();
        Devices.Add(profile);
    }

    /// <summary>Replaces the stored profile that matches <paramref name="updated"/>.Id.</summary>
    public bool UpdateDevice(DeviceProfile updated)
    {
        int idx = Devices.FindIndex(d => d.Id == updated.Id);
        if (idx < 0) return false;
        if (updated.IsDefault) ClearDeviceDefaults();
        Devices[idx] = updated;
        return true;
    }

    /// <summary>Removes the profile with the given Id. Returns true if found and removed.</summary>
    public bool RemoveDevice(string id)
    {
        int idx = Devices.FindIndex(d => d.Id == id);
        if (idx < 0) return false;
        Devices.RemoveAt(idx);
        return true;
    }

    private void ClearDeviceDefaults()
    {
        foreach (var d in Devices) d.IsDefault = false;
    }

    // ── Job templates ─────────────────────────────────────────────────────────

    /// <summary>Returns the template with the given Id, or null.</summary>
    public JobTemplate? FindTemplate(string id) =>
        Templates.FirstOrDefault(t => t.Id == id);

    /// <summary>Returns the default template, or null if none is marked default.</summary>
    public JobTemplate? DefaultTemplate =>
        Templates.FirstOrDefault(t => t.IsDefault) ?? Templates.FirstOrDefault();

    /// <summary>
    /// Adds <paramref name="template"/> to the collection.
    /// If IsDefault is true, clears IsDefault on all others.
    /// </summary>
    public void AddTemplate(JobTemplate template)
    {
        ArgumentNullException.ThrowIfNull(template);
        if (template.IsDefault) ClearTemplateDefaults();
        Templates.Add(template);
    }

    /// <summary>Replaces the stored template that matches <paramref name="updated"/>.Id.</summary>
    public bool UpdateTemplate(JobTemplate updated)
    {
        int idx = Templates.FindIndex(t => t.Id == updated.Id);
        if (idx < 0) return false;
        if (updated.IsDefault) ClearTemplateDefaults();
        Templates[idx] = updated;
        return true;
    }

    /// <summary>Removes the template with the given Id. Returns true if found and removed.</summary>
    public bool RemoveTemplate(string id)
    {
        int idx = Templates.FindIndex(t => t.Id == id);
        if (idx < 0) return false;
        Templates.RemoveAt(idx);
        return true;
    }

    private void ClearTemplateDefaults()
    {
        foreach (var t in Templates) t.IsDefault = false;
    }
}
