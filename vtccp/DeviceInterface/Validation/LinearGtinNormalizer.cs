namespace DeviceInterface.Validation;

/// <summary>
/// Expands standalone retail linear symbols to the GTIN-14 representation
/// required by GS1 AI (01) and EPC comparisons.
/// </summary>
public static class LinearGtinNormalizer
{
    public static string? NormalizeToGtin14(
        string? symbology,
        string? decodedData)
    {
        string symbol = CompactSymbol(symbology);
        string data = decodedData?.Trim() ?? string.Empty;

        if (data.Length == 0 || data.Any(ch => !char.IsDigit(ch)))
            return null;

        return symbol switch
        {
            "UPCA" when data.Length == 12 => "00" + data,
            "EAN13" when data.Length == 13 => "0" + data,
            "EAN8" when data.Length == 8 => "000000" + data,
            "UPCE" when data.Length == 8 => ExpandUpceToGtin14(data),
            _ => null,
        };
    }

    public static string? BuildElementString(
        string? symbology,
        string? decodedData)
    {
        string? gtin14 = NormalizeToGtin14(symbology, decodedData);
        return gtin14 is null ? null : $"(01){gtin14}";
    }

    /// <summary>
    /// Expands an eight-digit UPC-E HRI in the form NabcdefC.
    /// The GS1 reference implementation documents these branches for
    /// number-system zero; non-zero number systems are intentionally rejected.
    /// </summary>
    public static string? ExpandUpceToUpca(string? decodedData)
    {
        string data = decodedData?.Trim() ?? string.Empty;
        if (data.Length != 8 || data.Any(ch => !char.IsDigit(ch)) || data[0] != '0')
            return null;

        char a = data[1], b = data[2], c = data[3];
        char d = data[4], e = data[5], f = data[6], check = data[7];

        string upca = f switch
        {
            '0' or '1' or '2' => $"0{a}{b}{f}0000{c}{d}{e}{check}",
            '3' => $"0{a}{b}{c}00000{d}{e}{check}",
            '4' => $"0{a}{b}{c}{d}00000{e}{check}",
            >= '5' and <= '9' => $"0{a}{b}{c}{d}{e}0000{f}{check}",
            _ => string.Empty,
        };

        return upca.Length == 12 && IsValidCheckDigit(upca) ? upca : null;
    }

    private static string? ExpandUpceToGtin14(string data)
    {
        string? upca = ExpandUpceToUpca(data);
        return upca is null ? null : "00" + upca;
    }

    private static string CompactSymbol(string? symbology)
        => new string((symbology ?? string.Empty)
            .Where(char.IsLetterOrDigit)
            .ToArray())
            .ToUpperInvariant();

    private static bool IsValidCheckDigit(string digits)
    {
        int sum = 0;
        for (int i = digits.Length - 2, positionFromRight = 0;
             i >= 0;
             i--, positionFromRight++)
        {
            int digit = digits[i] - '0';
            sum += digit * (positionFromRight % 2 == 0 ? 3 : 1);
        }

        int expected = (10 - sum % 10) % 10;
        return expected == digits[^1] - '0';
    }
}