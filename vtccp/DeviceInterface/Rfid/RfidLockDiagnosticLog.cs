// Copyright © 2026 VCCS. All rights reserved.

using System.Text;

namespace DeviceInterface.Rfid;

/// <summary>
/// Persistent release-build diagnostics for ASR-P35U lock checks.
/// </summary>
/// <remarks>
/// The log is intentionally independent of Debug.WriteLine so an on-device
/// production run records the immediate CheckTagStatus return and every SDK
/// success/error callback. The active file is capped and rotated once.
/// </remarks>
public static class RfidLockDiagnosticLog
{
    private const long MaxBytes = 1_000_000;
    private static readonly object Gate = new();

    /// <summary>
    /// Full path of the diagnostic log.
    /// Windows: %APPDATA%\VTCCP\rfid-lock-diagnostic.log.
    /// </summary>
    public static string FilePath
    {
        get
        {
            string appData =
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (string.IsNullOrWhiteSpace(appData))
                appData = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(appData, "VTCCP", "rfid-lock-diagnostic.log");
        }
    }

    internal static void Record(long requestId, string eventName, string details)
    {
        string line =
            $"{DateTimeOffset.Now:O} request={requestId} event={OneLine(eventName)} {OneLine(details)}";

        System.Diagnostics.Trace.WriteLine($"[RFID-LOCK] {line}");

        try
        {
            lock (Gate)
            {
                string path = FilePath;
                string directory = Path.GetDirectoryName(path)!;
                Directory.CreateDirectory(directory);

                if (File.Exists(path) && new FileInfo(path).Length >= MaxBytes)
                {
                    string previous = Path.Combine(
                        directory, "rfid-lock-diagnostic.previous.log");
                    File.Move(path, previous, overwrite: true);
                }

                File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch (Exception ex)
        {
            // Diagnostics must never break or delay the scan result path.
            System.Diagnostics.Trace.WriteLine(
                $"[RFID-LOCK] log-write-failed {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static string OneLine(string value) =>
        value.Replace('\r', ' ').Replace('\n', ' ');
}