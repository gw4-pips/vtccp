namespace DeviceInterface.Rfid;

/// <summary>
/// Accepts only successful, printable literal values returned by a reader SDK.
/// </summary>
internal static class ReaderIdentityValue
{
    private const int MaximumLength = 128;

    internal static string? FromSdkResponse(uint returnCode, string? value)
    {
        if (returnCode != 0 || string.IsNullOrWhiteSpace(value))
            return null;

        string literal = value.Trim();
        if (literal.Length > MaximumLength ||
            literal.Any(char.IsControl) ||
            literal.Contains('\uFFFD'))
        {
            return null;
        }

        return literal;
    }
}