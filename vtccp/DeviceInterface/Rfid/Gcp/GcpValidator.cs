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
    /// A missing registry entry is deliberately distinct from a found prefix whose
    /// registered length disagrees with the EPC partition.
    /// </summary>
    public GcpValidationStatus Validate(ParsedEpc epc)
    {
        if (epc.CompanyPrefix is null || epc.Partition is null)
            return GcpValidationStatus.NotChecked;

        int? claimedLength = GetEncodedGcpLength(epc);
        if (!claimedLength.HasValue)
            return GcpValidationStatus.NotChecked;

        if (!TryGetRegisteredLength(epc, out int registeredLength))
            return GcpValidationStatus.NotFound;

        return registeredLength == claimedLength.Value
            ? GcpValidationStatus.Valid
            : GcpValidationStatus.Invalid;
    }

    /// <summary>
    /// Returns the registered GCP length for an EPC's company prefix, when that
    /// prefix exists in the loaded table.
    /// </summary>
    public bool TryGetRegisteredLength(ParsedEpc? epc, out int registeredLength)
    {
        registeredLength = 0;
        return epc?.CompanyPrefix is not null
            && _table.TryLookup(epc.CompanyPrefix, out registeredLength);
    }

    /// <summary>
    /// Returns the GCP digit length encoded by an SGTIN partition, independent of
    /// whether a GCP lookup table is available.
    /// </summary>
    public static int? GetEncodedGcpLength(ParsedEpc? epc)
    {
        if (epc?.Partition is not int partition)
            return null;

        return GetGcpLengthFromPartition(epc.Scheme, partition) is int length and >= 0
            ? length
            : null;
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
