namespace DeviceInterface.Reports;

/// <summary>Dedicated entry point for the Release 26.0 linear report (5.12.7.3).</summary>
public static class Gs1LinearHtmlReportGenerator
{
    public static Gs1ReportValidationResult Validate(Gs1ReportData report) =>
        Gs1CanonicalReport.ValidateLinear(report);

    public static string Generate(Gs1ReportData report) =>
        Gs1CanonicalReport.GenerateLinearHtml(report);

    public static string Generate(Gs1ReportData report, Gs1PrintProfile printProfile) =>
        Gs1CanonicalReport.GenerateLinearHtml(report, printProfile);
}

/// <summary>Dedicated entry point for the Release 26.0 2D report (5.12.7.4).</summary>
public static class Gs1TwoDimensionalHtmlReportGenerator
{
    public static Gs1ReportValidationResult Validate(Gs1ReportData report) =>
        Gs1CanonicalReport.ValidateTwoDimensional(report);

    public static string Generate(Gs1ReportData report) =>
        Gs1CanonicalReport.GenerateTwoDimensionalHtml(report);

    public static string Generate(Gs1ReportData report, Gs1PrintProfile printProfile) =>
        Gs1CanonicalReport.GenerateTwoDimensionalHtml(report, printProfile);
}