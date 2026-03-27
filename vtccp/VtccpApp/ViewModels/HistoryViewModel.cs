namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
using ExcelEngine.Models;
using VtccpApp.Commands;
using VtccpApp.Models;

/// <summary>
/// Backs the Results History page.
///
/// <see cref="AddRecord"/> is called by <see cref="SessionViewModel"/> each time a
/// scan result is received. Records are stored in <see cref="AllRecords"/> and
/// projected into <see cref="FilteredRecords"/> according to the active filter.
///
/// All mutation happens on the UI thread (RelayCommand + WPF data binding ensure this).
/// </summary>
public sealed class HistoryViewModel : ViewModelBase
{
    // ── Storage ───────────────────────────────────────────────────────────────

    /// <summary>Every record received in the current session, in arrival order.</summary>
    public ObservableCollection<ScanResultRow> AllRecords      { get; } = [];

    /// <summary>Filtered subset shown in the DataGrid.</summary>
    public ObservableCollection<ScanResultRow> FilteredRecords { get; } = [];

    // ── Selection / detail strip ──────────────────────────────────────────────

    private ScanResultRow? _selectedRow;

    /// <summary>
    /// The row currently selected in the DataGrid.
    /// Drives the detail strip below the grid.
    /// </summary>
    public ScanResultRow? SelectedRow
    {
        get => _selectedRow;
        set
        {
            Set(ref _selectedRow, value);
            OnPropertyChanged(nameof(HasSelection));
            OnPropertyChanged(nameof(DetailGrade));
            OnPropertyChanged(nameof(DetailSymbology));
            OnPropertyChanged(nameof(DetailFullDecode));
            OnPropertyChanged(nameof(DetailUec));
            OnPropertyChanged(nameof(DetailDateTime));
            OnPropertyChanged(nameof(DetailDevice));
            OnPropertyChanged(nameof(DetailOperator));
        }
    }

    /// <summary>True when a row is selected; controls detail strip visibility.</summary>
    public bool HasSelection => _selectedRow is not null;

    // ── Detail strip computed properties ─────────────────────────────────────

    public string DetailGrade    => _selectedRow is null ? string.Empty
        : $"{_selectedRow.NumericGradeDisplay}  ({_selectedRow.Grade})  —  {_selectedRow.PassFail}";

    public string DetailSymbology => _selectedRow?.Symbology ?? string.Empty;

    public string DetailFullDecode => _selectedRow?.FullDecodedData ?? string.Empty;

    public string DetailUec      => _selectedRow?.UecPercent is { } u
        ? $"UEC  {u:F1}%"
        : string.Empty;

    public string DetailDateTime => _selectedRow is null ? string.Empty
        : _selectedRow.Source.VerificationDateTime.ToString("yyyy-MM-dd  HH:mm:ss");

    public string DetailDevice   => _selectedRow is null ? string.Empty
        : string.Join("  ·  ",
            new[] {
                _selectedRow.Source.DeviceName,
                _selectedRow.Source.DeviceSerial,
                _selectedRow.Source.FirmwareVersion,
            }.Where(s => !string.IsNullOrWhiteSpace(s)));

    public string DetailOperator => _selectedRow is null ? string.Empty
        : string.Join("  ·  ",
            new[] {
                _selectedRow.OperatorId,
                _selectedRow.JobName,
            }.Where(s => !string.IsNullOrWhiteSpace(s)));

    // ── Filter ────────────────────────────────────────────────────────────────

    private readonly HistoryFilter _filter = new();

    private string _gradeFilter     = "All";
    private string _passFailFilter  = "All";
    private string _symbologyFilter = "All";
    private string _statusMessage   = "No records yet.";

    public string GradeFilter
    {
        get => _gradeFilter;
        set { Set(ref _gradeFilter, value); _filter.GradeFilter = value; ApplyFilter(); }
    }

    public string PassFailFilter
    {
        get => _passFailFilter;
        set { Set(ref _passFailFilter, value); _filter.PassFailFilter = value; ApplyFilter(); }
    }

    public string SymbologyFilter
    {
        get => _symbologyFilter;
        set { Set(ref _symbologyFilter, value); _filter.SymbologyFilter = value; ApplyFilter(); }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => Set(ref _statusMessage, value);
    }

    // ── Combo-box option lists ─────────────────────────────────────────────────

    public static IReadOnlyList<string> GradeOptions    { get; } = ["All", "A", "B", "C", "D", "F", "—"];
    public static IReadOnlyList<string> PassFailOptions { get; } = ["All", "Pass", "Fail", "N/A"];

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand ClearCommand        { get; }
    public RelayCommand ClearFiltersCommand { get; }
    public RelayCommand CopyCommand         { get; }

    public HistoryViewModel()
    {
        ClearCommand        = new RelayCommand(OnClear,        () => AllRecords.Count > 0);
        ClearFiltersCommand = new RelayCommand(OnClearFilters, () => !_filter.IsEmpty);
        CopyCommand         = new RelayCommand(OnCopy,         () => FilteredRecords.Count > 0);
    }

    // ── Public API called by SessionViewModel ─────────────────────────────────

    /// <summary>Adds a scan record. Must be called on the UI thread.</summary>
    public void AddRecord(VerificationRecord record)
    {
        var row = ScanResultRow.From(record, AllRecords.Count + 1);
        AllRecords.Add(row);
        if (_filter.Matches(row)) FilteredRecords.Add(row);
        UpdateStatus();
        RelayCommand.Refresh();
    }

    /// <summary>Clears all records. Called when a new session starts.</summary>
    public void ClearHistory()
    {
        SelectedRow = null;
        AllRecords.Clear();
        FilteredRecords.Clear();
        UpdateStatus();
        RelayCommand.Refresh();
    }

    // ── Filter helpers ────────────────────────────────────────────────────────

    private void ApplyFilter()
    {
        FilteredRecords.Clear();
        foreach (var row in AllRecords)
            if (_filter.Matches(row))
                FilteredRecords.Add(row);
        UpdateStatus();
        RelayCommand.Refresh();
    }

    private void UpdateStatus()
    {
        if (AllRecords.Count == 0)
        {
            StatusMessage = "No records yet.";
            return;
        }

        int pass = AllRecords.Count(r => r.IsPass);
        int fail = AllRecords.Count - pass;
        string filterNote = _filter.IsEmpty
            ? string.Empty
            : $"  (filtered: {FilteredRecords.Count} shown)";

        StatusMessage = $"{AllRecords.Count} records  ·  {pass} pass  ·  {fail} fail{filterNote}";
    }

    // ── Command handlers ──────────────────────────────────────────────────────

    private void OnClear() => ClearHistory();

    private void OnClearFilters()
    {
        GradeFilter     = "All";
        PassFailFilter  = "All";
        SymbologyFilter = "All";
    }

    private void OnCopy()
    {
        if (FilteredRecords.Count == 0) return;

        const string sep = "\t";
        var sb = new System.Text.StringBuilder();

        sb.AppendLine(string.Join(sep,
            "#", "Time", "Symbology", "Numeric Grade", "Letter", "Pass/Fail",
            "UEC%", "Full Decoded Data", "Operator", "Job"));

        foreach (var r in FilteredRecords)
        {
            sb.AppendLine(string.Join(sep,
                r.RowNumber,
                r.Time,
                r.Symbology,
                r.NumericGradeDisplay,
                r.Grade,
                r.PassFail,
                r.UecPercent?.ToString("F1") ?? string.Empty,
                r.FullDecodedData,          // full string, never truncated
                r.OperatorId,
                r.JobName));
        }

        try { System.Windows.Clipboard.SetText(sb.ToString()); }
        catch { /* clipboard may be unavailable during testing */ }
    }
}
