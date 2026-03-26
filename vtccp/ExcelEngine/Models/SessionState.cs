namespace ExcelEngine.Models;

/// <summary>
/// Captures the operator-entered session context that applies to all records within a job.
/// These fields come entirely from the VTCCP UI (not from the device) because DMV devices
/// cannot pass all header fields through DMCC programming — and certain characters (&, /, etc.)
/// conflict with DMCC command syntax.
/// </summary>
public sealed class SessionState
{
    public string? JobName { get; set; }
    public string? OperatorId { get; set; }
    public int RollNumber { get; set; } = 1;
    public string? BatchNumber { get; set; }
    public string? CompanyName { get; set; }
    public string? ProductName { get; set; }
    public string? CustomNote { get; set; }

    // Device-supplied fields (populated from device connection or PDF import in Phase 2)
    public string? DeviceSerial { get; set; }
    public string? DeviceName { get; set; }
    public string? FirmwareVersion { get; set; }
    public DateTime? CalibrationDate { get; set; }

    // Output preferences
    public OutputFormat OutputFormat { get; set; } = OutputFormat.Xlsx;
    public string? OutputDirectory { get; set; }

    // Session tracking
    public DateTime SessionStarted { get; set; } = DateTime.Now;
    public int RecordCount { get; set; } = 0;

    /// <summary>User-defined fields (User 1, User 2 in Webscan schema)</summary>
    public string? User1 { get; set; }
    public string? User2 { get; set; }
}
