namespace DeviceInterface.Dmcc;

using ExcelEngine.Models;

/// <summary>
/// Maps the three live TruCheck configuration responses into the per-result
/// report fields. Shared by persistent SDK sessions and short-lived raw DMCC
/// Push-mode reads so every delivery path produces the same display values.
/// </summary>
public static class TruCheckSettingsSnapshot
{
    public static VerificationRecord Apply(
        VerificationRecord record,
        DmccResponse applicationStandard,
        DmccResponse dataFormatCheck,
        DmccResponse apertureSetting)
        => record with
        {
            ApplicationStandardSetting = applicationStandard.StatusCode == DmccStatus.Ok
                ? MapApplicationStandard(applicationStandard.Body)
                : null,
            DataFormatCheckSetting = dataFormatCheck.StatusCode == DmccStatus.Ok
                ? MapDataFormatCheckSetting(dataFormatCheck.Body)
                : null,
            ApertureSettingMode = apertureSetting.StatusCode == DmccStatus.Ok
                ? MapApertureSettingMode(apertureSetting.Body)
                : null,
        };

    public static string? MapApplicationStandard(string? rawValue)
        => rawValue?.Trim() switch
        {
            "0" => "GS1",
            "1" => "HIBCC",
            "2" => "UDI (GS1 or HIBCC)",
            "3" => "UID (MIL-STD-130)",
            "4" => "Custom",
            "5" => "Auto",
            "6" => "Cryptocode",
            _   => null,
        };

    public static string? MapDataFormatCheckSetting(string? rawValue)
        => rawValue?.Trim() switch
        {
            "0" => "None",
            "1" => "GS1",
            "2" => "HIBCC",
            "3" => "ISO 15434",
            _   => null,
        };

    public static string? MapApertureSettingMode(string? rawValue)
        => rawValue?.Trim() switch
        {
            "0" => "User Set",
            "1" => "Auto 50%",
            "2" => "Auto Aperture",
            _   => null,
        };
}