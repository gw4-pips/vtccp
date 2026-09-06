namespace DeviceInterface.FileAdapters;

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ExcelEngine.Models;

/// <summary>Literal row emitted by an Axicon automatic-export CSV template.</summary>
public sealed class AxiconAutomaticExportReport
{
    public string SourceFilePath { get; init; } = string.Empty;
    public string RawCsv { get; init; } = string.Empty;
    public DateTime SourceLastWriteTime { get; init; }
    public bool ParseSucceeded { get; init; }
    public string? ParseError { get; init; }
    public IReadOnlyDictionary<string, string> Values { get; init; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public string? Message => Get("Message");
    public string? Symbology => Get("Symbology") ?? Get("Subsymbology");
    public string? IsoGrade => Get("ISO Grade");
    public string? Overall => Get("Overall");
    public string? SoftwareVersion => Get("Software Version") ?? Get("Verifier Software Version");
    public string? CalibrationDateDisplay => Get("Calibration Date") ?? Get("Last Calibration Date");
    public string? Get(string name) => Values.TryGetValue(name, out string? value) ? value : null;

    public VerificationRecord ToVerificationRecord()
    {
        if (!ParseSucceeded) throw new InvalidDataException(ParseError ?? "Axicon automatic-export report did not parse.");
        string symbology = Symbology!;
        return new VerificationRecord
        {
            VerificationDateTime = SourceLastWriteTime,
            Symbology = symbology,
            SymbologyFamily = AxiconAutomaticExportParser.MapSymbologyFamily(symbology),
            DecodedData = Message,
            VerifierBrand = AxiconFileAdapter.Brand,
            DeviceSerial = Get("Reader Serial"),
            ConnectionMedium = "File export",
            SoftwareVersion = SoftwareVersion,
            CalibrationDate = AxiconAutomaticExportParser.ParseDate(CalibrationDateDisplay),
            SourceArtifactPath = SourceFilePath,
            SourceArtifactType = "Axicon automatic-export CSV",
            VerifierSourceFields = Values,
            Standard = IsoGrade,
            OverallGrade = AxiconAutomaticExportParser.ParseGrade(Overall),
            SymbolAnsiGrade = AxiconAutomaticExportParser.ParseGrade(Overall),
            NominalXDim_1D = AxiconAutomaticExportParser.ParseDecimal(Get("X Dimension")),
            NominalXDim_2D = AxiconAutomaticExportParser.ParseDecimal(Get("X Dimension")),
            SC_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Symbol Contrast")),
            MOD_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Modulation")),
            ANU_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Axial Nonuniformity")),
            GNU_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Grid Nonuniformity")),
            RM_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Reflectance Margin")),
            UEC_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Unused Error Correction")),
            FPD_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Fixed Pattern Damage")),
            DECODE_Grade = AxiconAutomaticExportParser.ParseGrade(Get("Decode")),
            Rmin = AxiconAutomaticExportParser.ParseDecimal(Get("Rmin")),
            Rmax = AxiconAutomaticExportParser.ParseDecimal(Get("Rmax")),
            Avg_Edge = AxiconAutomaticExportParser.ParseDecimal(Get("Edge Contrast")),
            Avg_Defect = AxiconAutomaticExportParser.ParseDecimal(Get("Defects")),
            Avg_DEC = AxiconAutomaticExportParser.ParseDecimal(Get("Decodability")),
            MatrixSize = Get("Matrix Size"),
            QR_Version = Get("QR Version"),
            QR_ECLevel = Get("QR ECC Level"),
        };
    }
}

/// <summary>Parser for the header/data CSV contract of the supplied Axicon Lua templates.</summary>
public static class AxiconAutomaticExportParser
{
    public static AxiconAutomaticExportReport Parse(string rawCsv, string sourcePath, DateTime sourceLastWriteTime)
    {
        try
        {
            List<string[]> rows = ReadCsv(rawCsv);
            if (rows.Count < 2) return Failure(rawCsv, sourcePath, sourceLastWriteTime, "Incomplete Axicon automatic export: expected header and data row.");
            string[] headers = rows[0];
            string[] data = rows[1];
            if (headers.Length != data.Length) return Failure(rawCsv, sourcePath, sourceLastWriteTime, "Incomplete Axicon automatic export: header and data column counts differ.");
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Length; i++)
                if (!string.IsNullOrWhiteSpace(headers[i])) values[headers[i].Trim()] = data[i];
            if (!values.TryGetValue("Message", out string? message) || string.IsNullOrWhiteSpace(message))
                return Failure(rawCsv, sourcePath, sourceLastWriteTime, "Incomplete Axicon automatic export: missing Message.");
            string? symbology = values.GetValueOrDefault("Symbology") ?? values.GetValueOrDefault("Subsymbology");
            if (string.IsNullOrWhiteSpace(symbology))
                return Failure(rawCsv, sourcePath, sourceLastWriteTime, "Incomplete Axicon automatic export: missing Symbology or Subsymbology.");
            string[] requiredVerifierFields =
                ["Reader Serial", "Software Version", "Calibration Date"];
            string[] missing = requiredVerifierFields.Where(field =>
                string.IsNullOrWhiteSpace(values.GetValueOrDefault(field))).ToArray();
            if (missing.Length > 0)
                return Failure(rawCsv, sourcePath, sourceLastWriteTime,
                    "Incomplete Axicon automatic export: missing " + string.Join(", ", missing) + ".");
            if (ParseDate(values["Calibration Date"]) is null)
                return Failure(rawCsv, sourcePath, sourceLastWriteTime,
                    "Invalid Axicon automatic export: Calibration Date is not a date.");
            SymbologyFamily family = MapSymbologyFamily(symbology);
            string[] requiredTechnicalFields = family switch
            {
                SymbologyFamily.GS1DataMatrix or SymbologyFamily.GS1QRCode =>
                    ["Matrix Size", "Data Structure", "GS1 Company Prefix", "Human Readable", "Print Growth"],
                SymbologyFamily.DataMatrix or SymbologyFamily.QRCode =>
                    ["Matrix Size", "Print Growth"],
                _ => ["X Dimension", "Rmin", "Rmax", "Edge Contrast", "Defects", "Decodability"],
            };
            missing = requiredTechnicalFields.Where(field =>
                string.IsNullOrWhiteSpace(values.GetValueOrDefault(field))).ToArray();
            if (missing.Length > 0)
                return Failure(rawCsv, sourcePath, sourceLastWriteTime,
                    "Incomplete Axicon automatic export: missing " + string.Join(", ", missing) + ".");
            return new AxiconAutomaticExportReport { SourceFilePath = sourcePath, RawCsv = rawCsv, SourceLastWriteTime = sourceLastWriteTime, ParseSucceeded = true, Values = values };
        }
        catch (Exception ex) { return Failure(rawCsv, sourcePath, sourceLastWriteTime, $"Invalid Axicon automatic export: {ex.Message}"); }
    }

    public static SymbologyFamily MapSymbologyFamily(string value)
    {
        bool dataMatrix = value.Contains("Data Matrix", StringComparison.OrdinalIgnoreCase) ||
                          value.Contains("DataMatrix", StringComparison.OrdinalIgnoreCase);
        if (value.Contains("GS1", StringComparison.OrdinalIgnoreCase) && dataMatrix) return SymbologyFamily.GS1DataMatrix;
        if (value.Contains("GS1", StringComparison.OrdinalIgnoreCase) && value.Contains("QR", StringComparison.OrdinalIgnoreCase)) return SymbologyFamily.GS1QRCode;
        if (dataMatrix) return SymbologyFamily.DataMatrix;
        if (value.Contains("QR", StringComparison.OrdinalIgnoreCase)) return SymbologyFamily.QRCode;
        return SymbologyFamily.Linear1D;
    }
    internal static decimal? ParseDecimal(string? value)
    {
        Match match = Regex.Match(value ?? "", @"[-+]?\d+(?:[.,]\d+)?");
        return match.Success && decimal.TryParse(match.Value.Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal result) ? result : null;
    }
    internal static GradingResult? ParseGrade(string? value)
    {
        Match match = Regex.Match(value ?? "", @"(?:(?<letter>[A-F])\s*\(\s*)?(?<number>[0-4](?:[.,]\d+)?)\s*\)?", RegexOptions.IgnoreCase);
        if (!match.Success || !decimal.TryParse(match.Groups["number"].Value.Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal numeric)) return null;
        string letter = match.Groups["letter"].Success ? match.Groups["letter"].Value : "";
        return GradingResult.FromLetterAndNumeric(letter, numeric, string.Empty, value);
    }
    internal static DateTime? ParseDate(string? value) =>
        DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed)
            ? parsed : null;
    private static AxiconAutomaticExportReport Failure(string raw, string path, DateTime time, string error)
        => new() { SourceFilePath = path, RawCsv = raw, SourceLastWriteTime = time, ParseError = error };
    private static List<string[]> ReadCsv(string raw)
    {
        var rows = new List<string[]>(); var row = new List<string>(); var cell = new StringBuilder(); bool quoted = false;
        for (int i = 0; i < raw.Length; i++)
        {
            char c = raw[i];
            if (c == '"') { if (quoted && i + 1 < raw.Length && raw[i + 1] == '"') { cell.Append(c); i++; } else quoted = !quoted; }
            else if (c == ',' && !quoted) { row.Add(cell.ToString()); cell.Clear(); }
            else if ((c == '\r' || c == '\n') && !quoted) { if (c == '\r' && i + 1 < raw.Length && raw[i + 1] == '\n') i++; row.Add(cell.ToString()); cell.Clear(); if (row.Any(x => x.Length > 0)) rows.Add(row.ToArray()); row.Clear(); }
            else cell.Append(c);
        }
        if (quoted) throw new FormatException("unterminated quoted CSV value.");
        if (cell.Length > 0 || row.Count > 0) { row.Add(cell.ToString()); rows.Add(row.ToArray()); }
        return rows;
    }
}