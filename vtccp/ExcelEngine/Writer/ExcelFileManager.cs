namespace ExcelEngine.Writer;

using ExcelEngine.Models;

/// <summary>
/// Generates output file paths and names following the Webscan TruCheck naming convention,
/// and provides helper utilities for file management.
///
/// Webscan convention (replicated):
///   {JobName}_{YYYY-MM-DD}.xlsx
///   e.g. "CalCardProd_2026-03-17.xlsx"
///
/// VTCCP extensions when no job name is set:
///   VTCCP_{OperatorId}_{YYYY-MM-DD}.xlsx
///   e.g. "VTCCP_GW4_2026-03-17.xlsx"
/// </summary>
public static class ExcelFileManager
{
    /// <summary>
    /// Generate the output filename (no directory path) from session state and format.
    /// </summary>
    public static string GenerateFileName(SessionState session, OutputFormat format)
    {
        var ext = format == OutputFormat.Xls ? ".xls" : ".xlsx";
        var date = session.SessionStarted.ToString("yyyy-MM-dd");

        string baseName;
        if (!string.IsNullOrWhiteSpace(session.JobName))
        {
            baseName = Sanitize(session.JobName) + "_" + date;
        }
        else if (!string.IsNullOrWhiteSpace(session.OperatorId))
        {
            baseName = "VTCCP_" + Sanitize(session.OperatorId) + "_" + date;
        }
        else
        {
            baseName = "VTCCP_" + date;
        }

        return baseName + ext;
    }

    /// <summary>
    /// Resolve the full output file path. Uses session.OutputDirectory if set, otherwise
    /// defaults to the user's Documents\VTCCP folder.
    /// </summary>
    public static string ResolveOutputPath(SessionState session, OutputFormat format)
    {
        var dir = string.IsNullOrWhiteSpace(session.OutputDirectory)
            ? DefaultOutputDirectory()
            : session.OutputDirectory;

        Directory.CreateDirectory(dir);
        return Path.Combine(dir, GenerateFileName(session, format));
    }

    /// <summary>
    /// Check whether a file is locked by another process (e.g. open in Excel).
    /// Returns null if the file is accessible, otherwise a descriptive error message.
    /// </summary>
    public static string? CheckFileLocked(string filePath)
    {
        if (!File.Exists(filePath)) return null;
        try
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return null;
        }
        catch (IOException)
        {
            return $"The file '{Path.GetFileName(filePath)}' is open in another application (e.g. Excel). " +
                   "Please close it and try again.";
        }
    }

    /// <summary>
    /// Default output directory: %USERPROFILE%\Documents\VTCCP on Windows,
    /// or the user's home directory on other platforms.
    /// </summary>
    public static string DefaultOutputDirectory()
    {
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (string.IsNullOrEmpty(docs))
            docs = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return Path.Combine(docs, "VTCCP");
    }

    /// <summary>
    /// Sanitize a string for use in a filename — strip illegal characters, collapse spaces to underscores.
    /// </summary>
    private static string Sanitize(string input)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var result = string.Concat(input.Select(c => invalid.Contains(c) ? '_' : c));
        return result.Replace(' ', '_').Trim('_');
    }
}
