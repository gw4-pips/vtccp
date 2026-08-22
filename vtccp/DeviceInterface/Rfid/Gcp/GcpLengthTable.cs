using System.Xml.Linq;

namespace DeviceInterface.Rfid.Gcp;

/// <summary>
/// In-memory lookup table built from the GS1 GCP Prefix Format List XML file.
///
/// XML format (bundled seed table dated 2026-06-03):
///   &lt;GCPPrefixFormatList date="2026-06-03T11:14:42.028Z"&gt;
///     &lt;entry prefix="000000" gcpLength="7" /&gt;
///     ...
///   &lt;/GCPPrefixFormatList&gt;
///
/// Lookup: given the leftmost digits of a GS1 Company Prefix, look up the
/// official GCP length. Used by <see cref="GcpValidator"/> to confirm a
/// company prefix is in the GS1 registry.
/// </summary>
public sealed class GcpLengthTable
{
    // Key = prefix string (variable length), Value = declared GCP length
    private readonly Dictionary<string, int> _prefixToLength;

    /// <summary>Date attribute from the GCPPrefixFormatList root element, if present.</summary>
    public DateTimeOffset? DataDate { get; }

    private GcpLengthTable(Dictionary<string, int> table, DateTimeOffset? date)
    {
        _prefixToLength = table;
        DataDate = date;
    }

    /// <summary>Number of entries in the table.</summary>
    public int EntryCount => _prefixToLength.Count;

    /// <summary>
    /// Load the GCP table from an XML file path.
    /// Throws <see cref="FileNotFoundException"/> when the file does not exist.
    /// Throws <see cref="InvalidDataException"/> when the XML is malformed.
    /// </summary>
    public static GcpLengthTable LoadFromFile(string xmlPath)
    {
        if (!File.Exists(xmlPath))
            throw new FileNotFoundException($"GCP prefix list not found: {xmlPath}", xmlPath);

        using var stream = File.OpenRead(xmlPath);
        return LoadFromStream(stream);
    }

    /// <summary>Load the GCP table from an already-open stream.</summary>
    public static GcpLengthTable LoadFromStream(Stream stream)
    {
        XDocument doc;
        try { doc = XDocument.Load(stream); }
        catch (Exception ex) { throw new InvalidDataException("GCP XML is malformed.", ex); }

        var root = doc.Root ?? throw new InvalidDataException("GCP XML has no root element.");

        DateTimeOffset? date = null;
        if (root.Attribute("date")?.Value is { Length: > 0 } dateStr
            && DateTimeOffset.TryParse(dateStr, out var parsedDate))
        {
            date = parsedDate;
        }

        var table = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var entry in root.Elements("entry"))
        {
            string? prefix = entry.Attribute("prefix")?.Value;
            string? lenStr = entry.Attribute("gcpLength")?.Value;

            if (prefix is { Length: > 0 } && int.TryParse(lenStr, out int len))
                table[prefix] = len;
        }

        if (table.Count == 0)
            throw new InvalidDataException("GCP XML contains no valid entries.");

        return new GcpLengthTable(table, date);
    }

    /// <summary>
    /// Try to find the GCP length for a given company prefix string.
    ///
    /// The table stores prefix keys of varying lengths. We progressively try the
    /// candidate prefix left-truncated from the full string (matching the GS1 spec
    /// where the prefix can be 6–12 digits representing a portion of the GCP).
    ///
    /// Returns true and sets <paramref name="gcpLength"/> when found.
    /// </summary>
    public bool TryLookup(string companyPrefix, out int gcpLength)
    {
        gcpLength = 0;
        if (string.IsNullOrEmpty(companyPrefix)) return false;

        // Try progressively shorter prefixes (longest first for best match)
        for (int len = Math.Min(companyPrefix.Length, 12); len >= 1; len--)
        {
            string candidate = companyPrefix[..len];
            if (_prefixToLength.TryGetValue(candidate, out gcpLength))
                return true;
        }
        return false;
    }

    /// <summary>
    /// Return true if the company prefix is present in the GS1 GCP registry
    /// and its declared GCP length matches <paramref name="claimedLength"/>.
    /// </summary>
    public bool IsValidGcp(string companyPrefix, int claimedLength)
    {
        return TryLookup(companyPrefix, out int registeredLength)
            && registeredLength == claimedLength;
    }
}
