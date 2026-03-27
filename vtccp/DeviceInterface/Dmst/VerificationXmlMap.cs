namespace DeviceInterface.Dmst;

using ExcelEngine.Models;

/// <summary>
/// Maps DMST result XML element names to <see cref="VerificationRecord"/> fields.
///
/// The default values correspond to the Cognex DataMan DMV DMST XML format
/// (firmware 5.x, &lt;DMSymVerResponse&gt; container).  Override any property
/// to adapt to a different firmware version or custom XML template without
/// changing parser code.
///
/// XML path convention: simple element name (searched anywhere under the
/// result container), or "Parent/Child" for disambiguation when duplicate
/// element names exist at different levels.
/// </summary>
public sealed class VerificationXmlMap
{
    // ── Result container ──────────────────────────────────────────────────────

    /// <summary>Root element of the DMCC response envelope.</summary>
    public string ResponseRoot    { get; set; } = "DMCCResponse";

    /// <summary>Container element holding the verification result payload.</summary>
    public string ResultContainer { get; set; } = "DMSymVerResponse";

    // ── Identity / timing ─────────────────────────────────────────────────────

    public string DateTime        { get; set; } = "DateTime";
    public string SymbologyName   { get; set; } = "SymbologyName";
    public string DecodedData     { get; set; } = "DecodedData";

    // ── Grading summary ───────────────────────────────────────────────────────

    public string FormalGrade         { get; set; } = "FormalGrade";
    public string OverallGrade        { get; set; } = "OverallGrade";
    public string OverallGradeNumeric { get; set; } = "OverallGradeNumeric";

    // ── Verification settings ─────────────────────────────────────────────────

    public string ApertureRef { get; set; } = "ApertureRef";
    public string Wavelength  { get; set; } = "Wavelength";
    public string Lighting    { get; set; } = "Lighting";
    public string Standard    { get; set; } = "Standard";

    // ── 2D quality parameters ─────────────────────────────────────────────────

    public string UECPercent    { get; set; } = "UECPercent";
    public string UECGrade      { get; set; } = "UECGrade";
    public string SCPercent     { get; set; } = "SCPercent";
    public string SCGrade       { get; set; } = "SCGrade";
    public string MODGrade      { get; set; } = "MODGrade";
    public string RMGrade       { get; set; } = "RMGrade";
    public string ANUPercent    { get; set; } = "ANUPercent";
    public string ANUGrade      { get; set; } = "ANUGrade";
    public string GNUPercent    { get; set; } = "GNUPercent";
    public string GNUGrade      { get; set; } = "GNUGrade";
    public string FPDGrade      { get; set; } = "FPDGrade";
    public string DecodeGrade   { get; set; } = "DecodeGrade";
    public string AGValue       { get; set; } = "AGValue";
    public string AGGrade       { get; set; } = "AGGrade";

    // ── 2D matrix characteristics ─────────────────────────────────────────────

    public string MatrixSize            { get; set; } = "MatrixSize";
    public string HorizontalBWG         { get; set; } = "HorizontalBWG";
    public string VerticalBWG           { get; set; } = "VerticalBWG";
    public string EncodedCharacters     { get; set; } = "EncodedCharacters";
    public string TotalCodewords        { get; set; } = "TotalCodewords";
    public string DataCodewords         { get; set; } = "DataCodewords";
    public string ErrorCorrectionBudget { get; set; } = "ErrorCorrectionBudget";
    public string ErrorsCorrected       { get; set; } = "ErrorsCorrected";
    public string ErrorCapacityUsed     { get; set; } = "ErrorCapacityUsed";
    public string ErrorCorrectionType   { get; set; } = "ErrorCorrectionType";
    public string NominalXDim           { get; set; } = "NominalXDim";
    public string PixelsPerModule       { get; set; } = "PixelsPerModule";
    public string ImagePolarity         { get; set; } = "ImagePolarity";
    public string ContrastUniformity    { get; set; } = "ContrastUniformity";
    public string MRD                   { get; set; } = "MRD";

    // ── 2D quiet zones / borders (single-region ≤26×26) ──────────────────────

    public string LLSGrade { get; set; } = "LLSGrade";
    public string BLSGrade { get; set; } = "BLSGrade";
    public string LQZGrade { get; set; } = "LQZGrade";
    public string BQZGrade { get; set; } = "BQZGrade";
    public string TQZGrade { get; set; } = "TQZGrade";
    public string RQZGrade { get; set; } = "RQZGrade";

    // ── 2D transition ratios / clock tracks ───────────────────────────────────

    public string TTRPercent { get; set; } = "TTRPercent";
    public string TTRGrade   { get; set; } = "TTRGrade";
    public string RTRPercent { get; set; } = "RTRPercent";
    public string RTRGrade   { get; set; } = "RTRGrade";
    public string TCTGrade   { get; set; } = "TCTGrade";
    public string RCTGrade   { get; set; } = "RCTGrade";

    // ── 2D quadrant parameters (≥32×32) ──────────────────────────────────────

    public string ULQZGrade   { get; set; } = "ULQZGrade";
    public string URQZGrade   { get; set; } = "URQZGrade";
    public string RUQZGrade   { get; set; } = "RUQZGrade";
    public string RLQZGrade   { get; set; } = "RLQZGrade";

    public string ULQTTRPercent { get; set; } = "ULQTTRPercent";
    public string ULQTTRGrade  { get; set; } = "ULQTTRGrade";
    public string URQTTRPercent { get; set; } = "URQTTRPercent";
    public string URQTTRGrade  { get; set; } = "URQTTRGrade";
    public string LLQTTRPercent { get; set; } = "LLQTTRPercent";
    public string LLQTTRGrade  { get; set; } = "LLQTTRGrade";
    public string LRQTTRPercent { get; set; } = "LRQTTRPercent";
    public string LRQTTRGrade  { get; set; } = "LRQTTRGrade";

    public string ULQRTRPercent { get; set; } = "ULQRTRPercent";
    public string ULQRTRGrade  { get; set; } = "ULQRTRGrade";
    public string URQRTRPercent { get; set; } = "URQRTRPercent";
    public string URQRTRGrade  { get; set; } = "URQRTRGrade";
    public string LLQRTRPercent { get; set; } = "LLQRTRPercent";
    public string LLQRTRGrade  { get; set; } = "LLQRTRGrade";
    public string LRQRTRPercent { get; set; } = "LRQRTRPercent";
    public string LRQRTRGrade  { get; set; } = "LRQRTRGrade";

    public string ULQTCTGrade { get; set; } = "ULQTCTGrade";
    public string URQTCTGrade { get; set; } = "URQTCTGrade";
    public string LLQTCTGrade { get; set; } = "LLQTCTGrade";
    public string LRQTCTGrade { get; set; } = "LRQTCTGrade";

    public string ULQRCTGrade { get; set; } = "ULQRCTGrade";
    public string URQRCTGrade { get; set; } = "URQRCTGrade";
    public string LLQRCTGrade { get; set; } = "LLQRCTGrade";
    public string LRQRCTGrade { get; set; } = "LRQRCTGrade";

    // ── 1D ISO 15416 parameters ───────────────────────────────────────────────

    public string SymbolAnsiGrade { get; set; } = "SymbolAnsiGrade";
    public string AvgEdge         { get; set; } = "AvgEdge";
    public string AvgRlRd         { get; set; } = "AvgRlRd";
    public string AvgSC           { get; set; } = "AvgSC";
    public string AvgMinEC        { get; set; } = "AvgMinEC";
    public string AvgMOD          { get; set; } = "AvgMOD";
    public string AvgDefect       { get; set; } = "AvgDefect";
    public string AvgDcod         { get; set; } = "AvgDcod";
    public string AvgDEC          { get; set; } = "AvgDEC";
    public string AvgLQZ          { get; set; } = "AvgLQZ";
    public string AvgRQZ          { get; set; } = "AvgRQZ";
    public string AvgHQZ          { get; set; } = "AvgHQZ";
    public string AvgMinQZ        { get; set; } = "AvgMinQZ";
    public string BWGPercent      { get; set; } = "BWGPercent";
    public string Magnification   { get; set; } = "Magnification";
    public string Ratio           { get; set; } = "Ratio";
    public string NominalXDim1D   { get; set; } = "NominalXDim1D";

    // Per-scan sub-results for 1D
    public string ScanResults     { get; set; } = "ScanResults";
    public string ScanElement     { get; set; } = "Scan";
    public string ScanNumber      { get; set; } = "number";
    public string ScanEdge        { get; set; } = "Edge";
    public string ScanSC          { get; set; } = "SC";
    public string ScanMinEC       { get; set; } = "MinEC";
    public string ScanMOD         { get; set; } = "MOD";
    public string ScanDefect      { get; set; } = "Defect";
    public string ScanDEC         { get; set; } = "DEC";
    public string ScanLQZ         { get; set; } = "LQZ";
    public string ScanRQZ         { get; set; } = "RQZ";
    public string ScanHQZ         { get; set; } = "HQZ";

    // ── Symbology family classification ───────────────────────────────────────

    /// <summary>
    /// Maps the SymbologyName strings that the device emits to the VTCCP
    /// <see cref="SymbologyFamily"/> classification.  Keys are case-insensitive prefix
    /// matches — the first matching entry wins.
    /// </summary>
    public IReadOnlyList<(string Prefix, SymbologyFamily Family)> SymbologyFamilyMap { get; set; } =
    [
        ("GS1 Data Matrix",   SymbologyFamily.GS1DataMatrix),
        ("GS1 DataMatrix",    SymbologyFamily.GS1DataMatrix),
        ("GS1-DataMatrix",    SymbologyFamily.GS1DataMatrix),
        ("GS1 QR",            SymbologyFamily.GS1QRCode),
        ("GS1QR",             SymbologyFamily.GS1QRCode),
        ("Data Matrix ECC 200", SymbologyFamily.DataMatrix),
        ("Data Matrix",       SymbologyFamily.DataMatrix),
        ("DataMatrix",        SymbologyFamily.DataMatrix),
        ("DM DMRE",           SymbologyFamily.DMRE),
        ("DMRE",              SymbologyFamily.DMRE),
        ("DotCode",           SymbologyFamily.DotCode),
        ("Dot Code",          SymbologyFamily.DotCode),
        ("QR Code",           SymbologyFamily.QRCode),
        ("QRCode",            SymbologyFamily.QRCode),
        ("UPC",               SymbologyFamily.Linear1D),
        ("EAN",               SymbologyFamily.Linear1D),
        ("Code 128",          SymbologyFamily.Linear1D),
        ("Code128",           SymbologyFamily.Linear1D),
        ("Code 39",           SymbologyFamily.Linear1D),
        ("Code39",            SymbologyFamily.Linear1D),
        ("Interleaved",       SymbologyFamily.Linear1D),
        ("ITF",               SymbologyFamily.Linear1D),
        ("PDF417",            SymbologyFamily.Linear1D),
        ("Codabar",           SymbologyFamily.Linear1D),
        ("MSI",               SymbologyFamily.Linear1D),
    ];

    public SymbologyFamily ClassifySymbology(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return SymbologyFamily.Unknown;
        foreach (var (prefix, family) in SymbologyFamilyMap)
            if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return family;
        return SymbologyFamily.Unknown;
    }
}
