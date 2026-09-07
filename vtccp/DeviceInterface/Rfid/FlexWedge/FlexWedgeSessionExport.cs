using System.Text;
using System.Text.Json;

namespace DeviceInterface.Rfid.FlexWedge;

/// <summary>CSV row format suitable for Excel import without Excel automation.</summary>
public static class FlexWedgeCsvFormatter
{
    private static readonly FlexWedgeField[] Fields =
    [
        FlexWedgeField.Timestamp, FlexWedgeField.RawSource, FlexWedgeField.RawEpc, FlexWedgeField.Gtin14, FlexWedgeField.Serial,
        FlexWedgeField.EpcUri, FlexWedgeField.Scheme, FlexWedgeField.Filter, FlexWedgeField.Partition,
        FlexWedgeField.CompanyPrefix, FlexWedgeField.ItemReference, FlexWedgeField.GcpStatus,
        FlexWedgeField.GcpRegisteredLength, FlexWedgeField.Reader, FlexWedgeField.Port, FlexWedgeField.Rssi,
        FlexWedgeField.Pc, FlexWedgeField.Tid, FlexWedgeField.LockStatus, FlexWedgeField.DecodeWarnings,
        FlexWedgeField.HeaderBits, FlexWedgeField.FilterBits, FlexWedgeField.PartitionBits,
        FlexWedgeField.GcpBits, FlexWedgeField.ItemReferenceBits, FlexWedgeField.SerialBits
    ];
    public static string Header => string.Join(",", Fields.Select(x => Escape(x.ToString())));
    public static string FormatRow(FlexWedgeReadResult result) =>
        string.Join(",", Fields.Select(x => Escape(FlexWedgeOutputFormatter.GetValue(result, x))));
    public static string Escape(string value) =>
        value.IndexOfAny([',', '"', '\r', '\n']) >= 0 ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
}

/// <summary>Append-only in-memory session plus CSV and JSON Lines export methods.</summary>
public sealed class FlexWedgeSessionLogger
{
    private readonly List<FlexWedgeReadResult> _entries = [];
    public IReadOnlyList<FlexWedgeReadResult> Entries => _entries.AsReadOnly();
    public void Append(FlexWedgeReadResult result) => _entries.Add(result ?? throw new ArgumentNullException(nameof(result)));

    public void ExportCsv(Stream output)
    {
        ArgumentNullException.ThrowIfNull(output);
        using var writer = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        writer.WriteLine(FlexWedgeCsvFormatter.Header);
        foreach (var entry in _entries) writer.WriteLine(FlexWedgeCsvFormatter.FormatRow(entry));
    }

    public void ExportJsonLines(Stream output)
    {
        ArgumentNullException.ThrowIfNull(output);
        using var writer = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        foreach (var entry in _entries) writer.WriteLine(JsonSerializer.Serialize(entry));
    }

    public static void AppendJsonLine(Stream output, FlexWedgeReadResult result)
    {
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(result);
        using var writer = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        writer.WriteLine(JsonSerializer.Serialize(result));
    }

    public int ImportJsonLines(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);
        int imported = 0;
        using var reader = new StreamReader(input, Encoding.UTF8, true, 1024, leaveOpen: true);
        while (reader.ReadLine() is { } line)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try
            {
                var result = JsonSerializer.Deserialize<FlexWedgeReadResult>(line);
                if (result is null) continue;
                _entries.Add(result);
                imported++;
            }
            catch (JsonException)
            {
                // A partial final line after an interrupted write must not hide prior records.
            }
        }
        return imported;
    }
}