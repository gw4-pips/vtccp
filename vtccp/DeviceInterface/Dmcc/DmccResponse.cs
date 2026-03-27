namespace DeviceInterface.Dmcc;

/// <summary>
/// Parsed result of a single DMCC command exchange.
///
/// DMCC response wire format (CRLF = \r\n):
///   CRLF                    ← start marker (blank line)
///   {status_code}CRLF       ← 0 = OK, non-zero = error
///   CRLF                    ← present only when a body follows
///   {body_line_1}CRLF       ← zero or more body lines
///   {body_line_N}CRLF
///
/// Status codes:
///   0   = Success
///   6   = No read (device did not decode a symbol)
///   8   = Busy / command rejected
///   Other = device-specific error
/// </summary>
public sealed class DmccResponse
{
    /// <summary>DMCC status code. 0 = success.</summary>
    public int StatusCode { get; }

    /// <summary>True when StatusCode == 0.</summary>
    public bool IsSuccess => StatusCode == 0;

    /// <summary>Response body text, or empty string when status-only response.</summary>
    public string Body { get; }

    /// <summary>True when the body appears to contain an XML document.</summary>
    public bool IsXml => Body.TrimStart().StartsWith("<?xml", StringComparison.OrdinalIgnoreCase)
                      || Body.TrimStart().StartsWith("<", StringComparison.Ordinal);

    private DmccResponse(int statusCode, string body)
    {
        StatusCode = statusCode;
        Body       = body;
    }

    // ── Parsing ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Parses a raw DMCC response buffer (as received from the TCP stream).
    /// Tolerates minor formatting variations across firmware versions.
    /// </summary>
    public static DmccResponse Parse(string raw)
    {
        if (string.IsNullOrEmpty(raw))
            return new DmccResponse(DmccStatus.NoResponse, string.Empty);

        // Split on CRLF or just LF for robustness.
        var lines = raw.Split(["\r\n", "\n"], StringSplitOptions.None);

        // Walk past leading blank lines to find the status line.
        int idx = 0;
        while (idx < lines.Length && string.IsNullOrWhiteSpace(lines[idx]))
            idx++;

        // Status code line.
        if (idx >= lines.Length)
            return new DmccResponse(DmccStatus.ParseError, raw);

        if (!int.TryParse(lines[idx].Trim(), out int status))
            return new DmccResponse(DmccStatus.ParseError, raw);
        idx++;

        // Skip blank separator between status and body.
        while (idx < lines.Length && string.IsNullOrWhiteSpace(lines[idx]))
            idx++;

        // Remaining lines = body.
        string body = idx < lines.Length
            ? string.Join("\r\n", lines[idx..]).TrimEnd('\r', '\n')
            : string.Empty;

        return new DmccResponse(status, body);
    }

    public override string ToString() =>
        IsSuccess ? $"OK  | {(Body.Length > 60 ? Body[..60] + "…" : Body)}"
                  : $"ERR {StatusCode} | {Body}";
}

/// <summary>Synthetic DMCC status codes used internally when the wire protocol fails.</summary>
public static class DmccStatus
{
    public const int Ok         = 0;
    public const int NoRead     = 6;
    public const int Busy       = 8;
    public const int ParseError = -1;
    public const int NoResponse = -2;
    public const int Timeout    = -3;
}
