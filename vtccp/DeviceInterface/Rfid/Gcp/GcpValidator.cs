using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.Gcp;

/// <summary>
/// Validates GS1 Company Prefixes using a loaded <see cref="GcpLengthTable"/>.
/// Thread-safe for concurrent reads after construction.
/// </summary>
public sealed class GcpValidator
{
    private readonly GcpLengthTable _table;

    public GcpValidator(GcpLengthTable table)
        => _table = table ?? throw new ArgumentNullException(nameof(table));

    /// <summary>Date of the GCP data used for validation.</summary>
    public DateTimeOffset? DataDate => _table.DataDate;

    /// <summary>Number of registered GCP prefixes in the loaded table.</summary>
    public int RegisteredPrefixCount => _table.EntryCount;

    /// <summary>
    /// Validate the Company Prefix from a parsed EPC against the GS1 GCP registry.
    /// Returns true if the prefix is found in the registry with the correct declared GCP length.
    /// Returns false if not found or if the partition-implied length disagrees with the registry.
    /// Returns null if the parsed EPC does not contain GCP information (unknown scheme).
    /// </summary>
    public bool? Validate(ParsedEpc epc)
    {
        if (epc.CompanyPrefix is null || epc.Partition is null)
            return null;

        int claimedLength = GetGcpLengthFromPartition(epc.Scheme, epc.Partition.Value);
        if (claimedLength < 0) return null;

        return _table.IsValidGcp(epc.CompanyPrefix, claimedLength);
    }

    /// <summary>
    /// Validate a raw company prefix string against the registry.
    /// The expected GCP length is looked up from the registry and must match
    /// the character count of <paramref name="companyPrefix"/>.
    /// </summary>
    public bool ValidateRaw(string companyPrefix)
    {
        if (string.IsNullOrEmpty(companyPrefix)) return false;
        return _table.TryLookup(companyPrefix, out int expectedLen)
            && expectedLen == companyPrefix.Length;
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private static int GetGcpLengthFromPartition(EpcScheme scheme, int partition)
    {
        // Both SGTIN-96 and SGTIN-198 use the same partition→L mapping
        if (scheme is not (EpcScheme.Sgtin96 or EpcScheme.Sgtin198)) return -1;
        return partition switch
        {
            0 => 12,
            1 => 11,
            2 => 10,
            3 =>  9,
            4 =>  8,
            5 =>  7,
            6 =>  6,
            _ => -1,
        };
    }
}
