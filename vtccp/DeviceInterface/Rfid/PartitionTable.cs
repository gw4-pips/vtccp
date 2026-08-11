// Copyright © 2026 VCCS. All rights reserved.

namespace DeviceInterface.Rfid;

/// <summary>
/// Static GS1 EPC Tag Data Standard (TDS) v2.3 partition table for SGTIN encoding.
///
/// Partition value (3 bits, 0–6) determines how the 44-bit GCP+ItemRef field is split:
///   M = Company Prefix bit count
///   L = Company Prefix decimal digit count
///   N = Item Reference bit count
///   K = Item Reference decimal digit count
///
/// Invariants:
///   L + K = 13 always (together they encode the 13-digit GTIN-13 body)
///   M + N = 44 always (the fixed total field width in SGTIN-96/198)
///
/// The Item Reference field (K digits) includes the GTIN indicator digit as its
/// leading digit — it is NOT a separate field in the EPC encoding.
/// </summary>
public static class PartitionTable
{
    /// <summary>
    /// Represents one row in the GS1 TDS partition table.
    /// </summary>
    public readonly record struct Row(
        int Partition,
        int M,   // Company Prefix bits
        int L,   // Company Prefix decimal digits
        int N,   // Item Reference bits
        int K    // Item Reference decimal digits
    );

    /// <summary>
    /// All 7 rows of the GS1 TDS 2.3 SGTIN partition table (Table 14-1), indexed by
    /// partition value 0–6. Sorted ascending by partition.
    /// </summary>
    public static readonly IReadOnlyList<Row> Rows = new[]
    {
        new Row(0, M: 40, L: 12, N:  4, K: 1),
        new Row(1, M: 37, L: 11, N:  7, K: 2),
        new Row(2, M: 34, L: 10, N: 10, K: 3),
        new Row(3, M: 30, L:  9, N: 14, K: 4),
        new Row(4, M: 27, L:  8, N: 17, K: 5),
        new Row(5, M: 24, L:  7, N: 20, K: 6),
        new Row(6, M: 20, L:  6, N: 24, K: 7),
    };

    /// <summary>
    /// Look up a partition row by value (0–6).
    /// Throws <see cref="ArgumentOutOfRangeException"/> for values outside 0–6.
    /// </summary>
    public static Row Get(int partition)
    {
        if ((uint)partition > 6)
            throw new ArgumentOutOfRangeException(
                nameof(partition), partition, "SGTIN partition value must be 0–6.");
        return Rows[partition];
    }

    /// <summary>
    /// Try to look up a partition row. Returns false for values outside 0–6.
    /// </summary>
    public static bool TryGet(int partition, out Row row)
    {
        if ((uint)partition > 6)
        {
            row = default;
            return false;
        }
        row = Rows[partition];
        return true;
    }

    /// <summary>
    /// Return the GCP decimal digit count (L) for a given partition, or -1 if the
    /// partition value is out of range.
    /// </summary>
    public static int GcpDigits(int partition) =>
        (uint)partition <= 6 ? Rows[partition].L : -1;

    /// <summary>
    /// Return the Item Reference decimal digit count (K) for a given partition, or -1
    /// if the partition value is out of range.
    /// </summary>
    public static int ItemRefDigits(int partition) =>
        (uint)partition <= 6 ? Rows[partition].K : -1;
}
