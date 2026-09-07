using System.Runtime.InteropServices;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using DeviceInterface.Rfid;
using DeviceInterface.Rfid.FlexWedge;
using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace FlexWedgeApp;

public partial class MainWindow : Window
{
    private readonly string _configurationPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VCCS", "FlexWedge", "flexwedge.config.json");
    private readonly string _automaticSessionPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VCCS", "FlexWedge", "sessions", "current.jsonl");
    private FlexWedgeConfiguration _configuration = FlexWedgeConfiguration.SampleReadyDefault;
    private FlexWedgeReadResultBuilder _builder = new();
    private GcpValidator? _gcpValidator;
    private readonly FlexWedgeSessionLogger _session = new();
    private FlexWedgeReadResult? _current;
    private IEpcReader? _reader;
    private IntPtr _previousExternalWindow;
    private string? _customGcpXmlPath;
    private readonly DispatcherTimer _foregroundWatcher;

    public MainWindow()
    {
        InitializeComponent();
        AppendBox.ItemsSource = Enum.GetValues<FlexWedgeAppendKey>();
        _foregroundWatcher = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        _foregroundWatcher.Tick += (_, _) => CaptureExternalForegroundWindow();
        _foregroundWatcher.Start();
        LoadConfiguration();
        RecoverSession();
        ReloadGcp();
        RefreshProfiles();
#if ASREADER_SDK
        StatusText.Text = "ASR-P35U support is available when the vendor DLL is deployed.";
#else
        StatusText.Text = "ASR-P35U hardware support unavailable: AsReaderP3xU.dll was not present at build time. Manual EPC input remains available.";
#endif
    }

    private void LoadConfiguration()
    {
        try
        {
            string path = File.Exists(_configurationPath) ? _configurationPath : Path.Combine(AppContext.BaseDirectory, "flexwedge.config.json");
            string json = File.ReadAllText(path);
            _configuration = FlexWedgeConfigurationJson.Load(json);
            _customGcpXmlPath = _configuration.CustomGcpXmlPath;
        }
        catch (Exception ex) { StatusText.Text = "Configuration unavailable; using defaults. " + ex.Message; }
        PortBox.Text = _configuration.ReaderPort ?? "";
        GcpPathBox.Text = _customGcpXmlPath ?? "";
    }

    private void SaveConfiguration()
    {
        _configuration = _configuration with
        {
            ReaderPort = PortBox.Text.Trim(),
            CustomGcpXmlPath = string.IsNullOrWhiteSpace(GcpPathBox.Text) ? null : GcpPathBox.Text.Trim(),
            ActiveProfile = ((FlexWedgeOutputProfile)ProfileBox.SelectedItem).Name,
            OutputProfiles = _configuration.OutputProfiles.Select(p => p.Name == ((FlexWedgeOutputProfile)ProfileBox.SelectedItem).Name ? EditedProfile(p) : p).ToArray()
        };
        Directory.CreateDirectory(Path.GetDirectoryName(_configurationPath)!);
        File.WriteAllText(_configurationPath, FlexWedgeConfigurationJson.Save(_configuration));
        _customGcpXmlPath = GcpPathBox.Text.Trim();
        StatusText.Text = "Configuration saved: " + _configurationPath;
    }

    private void RefreshProfiles()
    {
        ProfileBox.ItemsSource = _configuration.OutputProfiles;
        ProfileBox.SelectedItem = _configuration.OutputProfiles.FirstOrDefault(p => p.Name == _configuration.ActiveProfile) ?? _configuration.OutputProfiles[0];
        LoadProfile((FlexWedgeOutputProfile)ProfileBox.SelectedItem);
    }
    private void LoadProfile(FlexWedgeOutputProfile p) { PrefixBox.Text = p.Prefix; SuffixBox.Text = p.Suffix; DelimiterBox.Text = p.Delimiter; AppendBox.SelectedItem = p.AppendKey; MappingBox.Text = string.Join("; ", p.Fields.Select(x => $"{x.Field}={x.Header}")); }
    private FlexWedgeOutputProfile EditedProfile(FlexWedgeOutputProfile p) => p with { Prefix = PrefixBox.Text, Suffix = SuffixBox.Text, Delimiter = DelimiterBox.Text, AppendKey = (FlexWedgeAppendKey)AppendBox.SelectedItem, Fields = ParseMappings(MappingBox.Text) };
    private FlexWedgeOutputProfile SelectedProfile() => EditedProfile((FlexWedgeOutputProfile)ProfileBox.SelectedItem);
    private void ProfileChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { if (ProfileBox.SelectedItem is FlexWedgeOutputProfile p) LoadProfile(p); }
    private void SaveConfig_Click(object sender, RoutedEventArgs e) { try { SaveConfiguration(); } catch (Exception ex) { StatusText.Text = "Configuration was not saved: " + ex.Message; } }

    private void ReloadGcp()
    {
        string path = string.IsNullOrWhiteSpace(GcpPathBox.Text) ? Path.Combine(AppContext.BaseDirectory, "data", "gcp-prefix-format-list.xml") : GcpPathBox.Text.Trim();
        try { var table = GcpLengthTable.LoadFromFile(path); _gcpValidator = new GcpValidator(table); _builder = new FlexWedgeReadResultBuilder(_gcpValidator); StatusText.Text = $"GCP table loaded: {table.EntryCount} entries ({path})."; }
        catch (Exception ex) { _gcpValidator = null; _builder = new FlexWedgeReadResultBuilder(); StatusText.Text = "GCP table unavailable or invalid; records remain NotChecked. " + ex.Message; }
    }
    private void ReloadGcp_Click(object sender, RoutedEventArgs e) => ReloadGcp();
    private void BrowseGcp_Click(object sender, RoutedEventArgs e) { var d = new OpenFileDialog { Filter = "GCP XML (*.xml)|*.xml|All files (*.*)|*.*" }; if (d.ShowDialog() == true) { GcpPathBox.Text = d.FileName; ReloadGcp(); } }

    private void ProcessManual_Click(object sender, RoutedEventArgs e) => Process(ManualEpcBox.Text);
    private void Process(string epc, FlexWedgeReaderAudit? audit = null)
    {
        try
        {
            _current = _builder.Build(epc, audit); _session.Append(_current);
            Directory.CreateDirectory(Path.GetDirectoryName(_automaticSessionPath)!);
            using (var stream = new FileStream(_automaticSessionPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                FlexWedgeSessionLogger.AppendJsonLine(stream, _current);
            ResultBox.Text = Describe(_current); RecentList.Items.Insert(0, $"{_current.Audit.Timestamp:HH:mm:ss}  {_current.RawEpcHex}  {_current.Gtin14}  {_current.Serial}");
            SessionText.Text = $"{_session.Entries.Count} reads; automatic JSONL: {_automaticSessionPath}";
        }
        catch (Exception ex) { StatusText.Text = "EPC was not processed: " + ex.Message; }
    }
    private static string Describe(FlexWedgeReadResult r) => $"Raw EPC: {r.RawEpcHex}\nScheme: {r.Scheme}\nGTIN: {r.Gtin14}\nSerial: {r.Serial}\nFilter: {r.Filter}  Partition: {r.Partition}\nCompany prefix: {r.CompanyPrefix}\nEPC URI: {r.EpcUri}\nGCP: {r.GcpStatus} / registered length {r.GcpRegisteredLength}\nTID: {r.Audit.Tid}\nLock: {r.Audit.LockStatus}\nBits: {Bits(r.HeaderBits)}; {Bits(r.FilterBits)}; {Bits(r.PartitionBits)}; {Bits(r.GcpBits)}; {Bits(r.ItemReferenceBits)}; {Bits(r.SerialBits)}";
    private static string Bits(EpcBitField? b) => b is null ? "" : $"{b.Name}[{b.StartBit}..{b.EndBitInclusive}]={b.Bits}";

    private void Connect_Click(object sender, RoutedEventArgs e)
    {
#if ASREADER_SDK
        _ = ConnectReaderAsync();
#else
        StatusText.Text = "Hardware support unavailable because AsReaderP3xU.dll was absent at build time.";
#endif
    }
#if ASREADER_SDK
    private async Task ConnectReaderAsync()
    {
        try
        {
            _reader ??= EpcReaderFactory.CreateAsReaderP35U();
            await _reader.ConnectAsync(PortBox.Text.Trim());
            StatusText.Text = "Connected to ASR-P35U on " + PortBox.Text.Trim();
        }
        catch (Exception ex) { StatusText.Text = "ASR-P35U connection failed: " + ex.Message; }
    }
#endif
    private async void Disconnect_Click(object sender, RoutedEventArgs e) { if (_reader is not null) { await _reader.DisconnectAsync(); await _reader.DisposeAsync(); _reader = null; } StatusText.Text = "Reader disconnected."; }
    private async void Trigger_Click(object sender, RoutedEventArgs e)
    {
        if (_reader is null || !_reader.IsConnected) { StatusText.Text = "Connect an ASR-P35U reader first."; return; }
        try
        {
            var reads = await _reader.TriggerInventoryAsync(TimeSpan.FromSeconds(2));
            if (reads.Count != 1) { StatusText.Text = $"Expected one tag; reader returned {reads.Count}."; return; }
            var x = reads[0];
            string? tid = x.Tid ?? await _reader.ReadTidAsync(x.EpcBytes, TimeSpan.FromSeconds(1));
            string? lockStatus = x.LockStatus ?? await _reader.ReadLockStatusAsync(x.EpcBytes, TimeSpan.FromSeconds(1));
            Process(x.EpcHex, new FlexWedgeReaderAudit(x.ReadTime, "ASR-P35U", PortBox.Text, x.Rssi, x.PcWord, tid, lockStatus));
        }
        catch (Exception ex) { StatusText.Text = "Trigger failed: " + ex.Message; }
    }

    private void Copy_Click(object sender, RoutedEventArgs e)
    {
        try { if (_current is null) return; System.Windows.Clipboard.SetText(FlexWedgeOutputFormatter.Format(_current, SelectedProfile())); StatusText.Text = "Output copied to clipboard."; }
        catch (Exception ex) { StatusText.Text = "Output was not copied: " + ex.Message; }
    }
    private void CaptureExternalForegroundWindow()
    {
        IntPtr own = new System.Windows.Interop.WindowInteropHelper(this).Handle;
        IntPtr candidate = GetForegroundWindow();
        GetWindowThreadProcessId(candidate, out uint processId);
        if (candidate != IntPtr.Zero && candidate != own && IsWindow(candidate) && processId != Environment.ProcessId)
            _previousExternalWindow = candidate;
    }
    private async void Send_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_current is null || _previousExternalWindow == IntPtr.Zero || !IsWindow(_previousExternalWindow))
            {
                StatusText.Text = "No valid previously focused external application was captured.";
                return;
            }
            System.Windows.Clipboard.SetText(FlexWedgeOutputFormatter.Format(_current, SelectedProfile()));
            if (!SetForegroundWindow(_previousExternalWindow)) { StatusText.Text = "Could not restore the previously focused application."; return; }
            await Task.Delay(100);
            if (GetForegroundWindow() != _previousExternalWindow) { StatusText.Text = "Destination focus could not be confirmed; nothing was pasted."; return; }
            keybd_event(0x11, 0, 0, UIntPtr.Zero); keybd_event(0x56, 0, 0, UIntPtr.Zero); keybd_event(0x56, 0, 2, UIntPtr.Zero); keybd_event(0x11, 0, 2, UIntPtr.Zero);
            StatusText.Text = "Pasted output into the confirmed previously focused application.";
        }
        catch (Exception ex) { StatusText.Text = "Output was not sent: " + ex.Message; }
    }
    private void Validate_Click(object sender, RoutedEventArgs e)
    {
        if (_current is null) { ValidationText.Text = "Read or enter exactly one EPC before validation."; return; }
        var read = new EpcReadResult { EpcBytes = _current.RawEpcBytes.ToArray() };
        var record = new VerificationRecord { Symbology = "GS1", DecodedData = BarcodeBox.Text };
        var result = new RfidValidator(_gcpValidator).Validate([read], record, 0);
        ValidationText.Text = $"VeriWedge: {result.Status}; RFID GTIN={result.RfidGtin14}; barcode GTIN={result.BarcodeGtin14}; RFID serial={result.RfidSerial}; barcode serial={result.BarcodeSerial}; GCP={result.GcpStatus}/{result.GcpRegisteredLength}; {result.MismatchDetail}";
    }
    private void Export_Click(object sender, RoutedEventArgs e)
    {
        var d = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv|JSON Lines (*.jsonl)|*.jsonl", FileName = "flexwedge-session" }; if (d.ShowDialog() != true) return;
        using var s = File.Create(d.FileName); if (Path.GetExtension(d.FileName).Equals(".jsonl", StringComparison.OrdinalIgnoreCase)) _session.ExportJsonLines(s); else _session.ExportCsv(s); StatusText.Text = "Session exported: " + d.FileName;
    }
    protected override async void OnClosed(EventArgs e)
    {
        _foregroundWatcher.Stop();
        if (_reader is not null)
        {
            try { await _reader.CancelAsync(); } catch { }
            await _reader.DisposeAsync();
        }
        base.OnClosed(e);
    }
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")] private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    private static IReadOnlyList<FlexWedgeFieldMapping> ParseMappings(string text)
    {
        var mappings = new List<FlexWedgeFieldMapping>();
        foreach (string token in text.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string[] parts = token.Split('=', 2, StringSplitOptions.TrimEntries);
            if (parts.Length != 2 || !Enum.TryParse(parts[0], true, out FlexWedgeField field) || string.IsNullOrWhiteSpace(parts[1]))
                throw new FormatException($"Invalid mapping '{token}'. Use Field=Header.");
            mappings.Add(new FlexWedgeFieldMapping(field, parts[1]));
        }
        return mappings;
    }

    private void RecoverSession()
    {
        if (!File.Exists(_automaticSessionPath)) return;
        try
        {
            using var stream = File.Open(_automaticSessionPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            int recovered = _session.ImportJsonLines(stream);
            SessionText.Text = $"{recovered} prior reads recovered from {_automaticSessionPath}";
            foreach (var item in _session.Entries.TakeLast(50).Reverse())
                RecentList.Items.Add($"{item.Audit.Timestamp:yyyy-MM-dd HH:mm:ss}  {item.RawEpcHex}  {item.Gtin14}  {item.Serial}");
        }
        catch (Exception ex) { StatusText.Text = "Prior session journal could not be recovered: " + ex.Message; }
    }
}