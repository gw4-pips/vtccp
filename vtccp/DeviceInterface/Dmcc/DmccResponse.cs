namespace DeviceInterface.Dmcc;

/// <summary>
/// Parsed result of a single DMCC command exchange.
///
/// Two wire formats are handled:
///
/// 1. Extended mode (COM.DMCC-RESPONSE=2) — actual device wire format:
///      ||[0]\r\n                    ← status 0 = OK
///      ||[101]\r\n                  ← invalid command
///      ||[102]\r\n                  ← invalid parameter
///      ||[104]\r\n                  ← rejected (reader state)
///    For commands with a data body (e.g. GET results):
///      ||[0]\r\n
///      ||[1]{base64-or-plain-data}\r\n
///
/// 2. Legacy / SDK-synthesised format (used internally in DataManSdkClient.SendAsync):
///      CRLF                    ← start marker
///      {status_code}CRLF       ← 0 = OK
///      CRLF
///      {body_line}CRLF
///
/// Status codes (device):
///   0   = Success
///   6   = No read
///   8   = Busy / rejected
///   101 = Invalid command
///   102 = Invalid parameter
///   104 = Parameter rejected due to reader state
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
    /// Parses a raw DMCC response buffer (as received from the TCP stream or
    /// synthesised internally by DataManSdkClient.SendAsync).
    /// Handles both the Extended wire format (||[N]\r\n) and the legacy
    /// SDK-synthesised format (\r\n{N}\r\n\r\n{body}).
    /// </summary>
    public static DmccResponse Parse(string raw)
    {
        if (string.IsNullOrEmpty(raw))
            return new DmccResponse(DmccStatus.NoResponse, string.Empty);

        string trimmed = raw.TrimStart();

        // ── Extended wire format: ||...[N]\r\n ────────────────────────────────
        // Sent by the device when COM.DMCC-RESPONSE = 2.
        // The prefix after "||" varies by firmware/session context; confirmed forms:
        //   ||[0]\r\n          — classic form (no session prefix)
        //   ||:::2[0]\r\n      — form observed on raw-TCP port-23 connections
        // Strategy: match any line starting with "||", find the rightmost [N] for status.
        if (trimmed.StartsWith("||", StringComparison.Ordinal))
        {
            var lines = raw.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries);

            // First line: find rightmost [...] group — that is the status code.
            int status = DmccStatus.ParseError;
            if (lines.Length > 0)
            {
                string first = lines[0].Trim();
                int rb = first.LastIndexOf(']');
                int lb = rb >= 0 ? first.LastIndexOf('[', rb) : -1;
                if (lb >= 0 && rb > lb &&
                    int.TryParse(first.AsSpan(lb + 1, rb - lb - 1), out int parsed))
                    status = parsed;
            }

            // Subsequent lines: strip any ||...] header prefix, collect body.
            var bodyParts = new System.Text.StringBuilder();
            for (int i = 1; i < lines.Length; i++)
            {
                string l = lines[i].Trim();
                if (l.StartsWith("||", StringComparison.Ordinal))
                {
                    int rb2 = l.IndexOf(']');
                    if (rb2 >= 0 && rb2 + 1 < l.Length)
                        l = l[(rb2 + 1)..];
                }
                if (bodyParts.Length > 0) bodyParts.Append("\r\n");
                bodyParts.Append(l);
            }

            return new DmccResponse(status, bodyParts.ToString());
        }

        // ── Legacy / SDK-synthesised format: \r\n{N}\r\n\r\n{body} ───────────
        var legLines = raw.Split(["\r\n", "\n"], StringSplitOptions.None);
        int idx = 0;
        while (idx < legLines.Length && string.IsNullOrWhiteSpace(legLines[idx]))
            idx++;

        if (idx >= legLines.Length)
            return new DmccResponse(DmccStatus.ParseError, raw);

        if (!int.TryParse(legLines[idx].Trim(), out int legStatus))
            return new DmccResponse(DmccStatus.ParseError, raw);
        idx++;

        while (idx < legLines.Length && string.IsNullOrWhiteSpace(legLines[idx]))
            idx++;

        string legBody = idx < legLines.Length
            ? string.Join("\r\n", legLines[idx..]).TrimEnd('\r', '\n')
            : string.Empty;

        return new DmccResponse(legStatus, legBody);
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
