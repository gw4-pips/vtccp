namespace ExcelEngine.Writer;

using ExcelEngine.Models;

/// <summary>
/// Generates output file paths and names for VTCCP log files,
/// and provides helper utilities for file management.
///
/// Default naming convention:
///   {JobName}_{YYYY-MM-DD}.xlsx
///   e.g. "CalCardProd_2026-03-17.xlsx"
///
/// VTCCP extensions when no job name is set:
///   VTCCP_{OperatorId}_{YYYY-MM-DD}.xlsx    (operator known)
///   VTCCP_{YYYY-MM-DD}.xlsx                 (neither set)
///
/// Custom pattern:
///   Set <see cref="SessionState.FileNamePattern"/> to a format string accepted by
///   <see cref="ApplyPattern"/>.  Supported tokens:
///     {Job}        — SanitizeFileName(JobName)  or "VTCCP"
///     {Op}         — SanitizeFileName(OperatorId) or ""
///     {Roll}       — RollNumber
///     {Date}       — SessionStarted "yyyy-MM-dd"
///     {DateTime}   — SessionStarted "yyyy-MM-dd_HH-mm"
///   Example: "{Job}_{Op}_Roll{Roll}_{Date}"
/// </summary>
public static class ExcelFileManager
{
    // Explicit set of characters that are illegal in Windows file names and
    // commonly problematic across platforms. Using a fixed set (rather than
    // Path.GetInvalidFileNameChars) ensures deterministic, platform-independent
    // behaviour on both Windows and Linux (CI/mono) targets.
    private static readonly char[] _illegalChars =
    [
        '/', '\\', ':', '*', '?', '"', '<', '>', '|',  // Windows-illegal
        '[', ']',                                        // Excel range syntax / NTFS edge-cases
        '\0', '\x01', '\x02', '\x03', '\x04', '\x05',  // NUL and other control chars
        '\x06', '\x07', '\x08', '\x09', '\x0A', '\x0B',
        '\x0C', '\x0D', '\x0E', '\x0F',
    ];

    /// <summary>
    /// Generate the output filename (no directory path) from session state and format.
    /// If <see cref="SessionState.FileNamePattern"/> is set it is used; otherwise the
    /// VTCCP default pattern is applied.
    /// </summary>
    public static string GenerateFileName(SessionState session, OutputFormat format)
    {
        var ext  = format == OutputFormat.Xls ? ".xls" : ".xlsx";
        var date = session.SessionStarted.ToString("yyyy-MM-dd");

        string baseName;

        if (!string.IsNullOrWhiteSpace(session.FileNamePattern))
        {
            baseName = ApplyPattern(session.FileNamePattern, session);
        }
        else if (!string.IsNullOrWhiteSpace(session.JobName))
        {
            baseName = SanitizeFileName(session.JobName) + "_" + date;
        }
        else if (!string.IsNullOrWhiteSpace(session.OperatorId))
        {
            baseName = "VTCCP_" + SanitizeFileName(session.OperatorId) + "_" + date;
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
    /// Sanitize a string for use as a filename component. Uses a fixed, explicit character
    /// set (not <c>Path.GetInvalidFileNameChars()</c>) so behaviour is identical on Windows,
    /// Linux, and macOS. Illegal characters are replaced with '_'; spaces are also replaced
    /// with '_'; leading and trailing underscores are trimmed.
    ///
    /// Characters that are illegal in filenames: / \ : * ? " &lt; &gt; | [ ] and NUL/control chars.
    ///
    /// PRODUCT DECISION (confirmed with user, 2026-03-26):
    ///   '&amp;' IS preserved in filenames because it is legal in Windows NTFS paths.
    ///   '&amp;' appears in legitimate operator-supplied values such as Company Name, and
    ///   VTCCP preserves it at the filesystem/Excel layer.
    ///   The DataMan DMCC device protocol does NOT support '&amp;' in command strings —
    ///   that restriction is handled separately in Phase 2 by a dedicated SanitizeForDmcc()
    ///   function and is NOT applied here.
    /// </summary>
    public static string SanitizeFileName(string input)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;
        var chars = input.ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            if (Array.IndexOf(_illegalChars, chars[i]) >= 0 || char.IsControl(chars[i]))
                chars[i] = '_';
            else if (chars[i] == ' ')
                chars[i] = '_';
        }
        return new string(chars).Trim('_');
    }

    /// <summary>
    /// Apply a custom file-name pattern. Replaces known tokens; any remaining text is
    /// left verbatim (callers should only use characters legal in file names).
    /// </summary>
    public static string ApplyPattern(string pattern, SessionState session)
    {
        var job      = string.IsNullOrWhiteSpace(session.JobName)    ? "VTCCP"
                       : SanitizeFileName(session.JobName);
        var op       = string.IsNullOrWhiteSpace(session.OperatorId) ? string.Empty
                       : SanitizeFileName(session.OperatorId);
        var roll     = session.RollLabel;
        var date     = session.SessionStarted.ToString("yyyy-MM-dd");
        var dateTime = session.SessionStarted.ToString("yyyy-MM-dd_HH-mm");

        return pattern
            .Replace("{Job}",      job,      StringComparison.OrdinalIgnoreCase)
            .Replace("{Op}",       op,       StringComparison.OrdinalIgnoreCase)
            .Replace("{Roll}",     roll,     StringComparison.OrdinalIgnoreCase)
            .Replace("{DateTime}", dateTime, StringComparison.OrdinalIgnoreCase)
            .Replace("{Date}",     date,     StringComparison.OrdinalIgnoreCase);
    }
}
