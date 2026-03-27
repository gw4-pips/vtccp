namespace VtccpApp.ViewModels;

using ConfigEngine.Models;

/// <summary>
/// Editable view-model wrapper around a <see cref="DeviceProfile"/>.
/// Exposes all editable fields as bindable properties and supports
/// Apply/Revert semantics.
/// </summary>
public sealed class DeviceProfileViewModel : ViewModelBase
{
    private string _id          = Guid.NewGuid().ToString();
    private string _name        = "New Device";
    private string _host        = "192.168.0.100";
    private int    _port        = 23;
    private int    _connectMs   = 5_000;
    private int    _responseMs  = 5_000;
    private int    _idleGapMs   = 150;
    private int    _dmstPort    = 0;
    private bool   _isDefault;
    private string _notes       = string.Empty;

    public string Id         { get => _id;         set => Set(ref _id,         value); }
    public string Name       { get => _name;       set => Set(ref _name,       value); }
    public string Host       { get => _host;       set => Set(ref _host,       value); }
    public int    Port       { get => _port;       set => Set(ref _port,       value); }
    public int    ConnectMs  { get => _connectMs;  set => Set(ref _connectMs,  value); }
    public int    ResponseMs { get => _responseMs; set => Set(ref _responseMs, value); }
    public int    IdleGapMs  { get => _idleGapMs;  set => Set(ref _idleGapMs,  value); }
    public int    DmstPort   { get => _dmstPort;   set => Set(ref _dmstPort,   value); }
    public bool   IsDefault  { get => _isDefault;  set => Set(ref _isDefault,  value); }
    public string Notes      { get => _notes;      set => Set(ref _notes,      value); }

    public DeviceProfileViewModel() { }

    public DeviceProfileViewModel(DeviceProfile profile) => LoadFrom(profile);

    public void LoadFrom(DeviceProfile p)
    {
        Id         = p.Id;
        Name       = p.Name;
        Host       = p.Host;
        Port       = p.Port;
        ConnectMs  = p.ConnectTimeoutMs;
        ResponseMs = p.ResponseTimeoutMs;
        IdleGapMs  = p.IdleGapMs;
        DmstPort   = p.DmstListenPort;
        IsDefault  = p.IsDefault;
        Notes      = p.Notes ?? string.Empty;
    }

    public DeviceProfile ToModel() => new()
    {
        Id                = Id,
        Name              = Name,
        Host              = Host,
        Port              = Port,
        ConnectTimeoutMs  = ConnectMs,
        ResponseTimeoutMs = ResponseMs,
        IdleGapMs         = IdleGapMs,
        DmstListenPort    = DmstPort,
        IsDefault         = IsDefault,
        Notes             = string.IsNullOrWhiteSpace(Notes) ? null : Notes,
    };

    public override string ToString() => $"{Name} ({Host}:{Port})";
}
