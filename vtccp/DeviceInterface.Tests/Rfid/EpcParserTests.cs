namespace DeviceInterface.Tests.Rfid;

using DeviceInterface.Rfid;
using DeviceInterface.Rfid.Models;
using Xunit;

/// <summary>
/// Unit tests for <see cref="EpcParser.ParseHex"/> covering the 12 vectors
/// in vtccp/references/asr-p35u/test-vectors/epc-decode-vectors.json.
/// </summary>
public sealed class EpcParserTests
{
    // ── Live captures (SGTIN-96, partition 5, 7-digit GCP) ───────────────────

    [Fact]
    public void LiveA_Sgtin96_DecodesCorrectly()
    {
        var result = EpcParser.ParseHex("30342A7CC844C7D0F36A0676");

        Assert.Equal(EpcScheme.Sgtin96, result.Scheme);
        Assert.Equal("0696114", result.CompanyPrefix);
        Assert.Equal("00696114704318", result.Gtin14);
        Assert.Equal("72803288694", result.Serial);
        Assert.Null(result.ParseWarning);
    }

    [Fact]
    public void LiveB_Sgtin96_DecodesCorrectly()
    {
        var result = EpcParser.ParseHex("30342A7CC844C710F36A0650");

        Assert.Equal(EpcScheme.Sgtin96, result.Scheme);
        Assert.Equal("0696114", result.CompanyPrefix);
        Assert.Equal("00696114704288", result.Gtin14);
        // NOTE: the test-vector JSON lists serial "72803288400" but direct bit
        // extraction from the EPC bytes gives 72803288656. The vector was wrong.
        Assert.Equal("72803288656", result.Serial);
        Assert.Null(result.ParseWarning);
    }

    [Fact]
    public void LiveC_Sgtin96_DecodesCorrectly()
    {
        var result = EpcParser.ParseHex("30342A7CC844C750F36A066F");

        Assert.Equal(EpcScheme.Sgtin96, result.Scheme);
        Assert.Equal("0696114", result.CompanyPrefix);
        Assert.Equal("00696114704295", result.Gtin14);
        // NOTE: the test-vector JSON lists serial "72803288175" but direct bit
        // extraction from the EPC bytes gives 72803288687. The vector was wrong.
        Assert.Equal("72803288687", result.Serial);
        Assert.Null(result.ParseWarning);
    }

    // ── Defect repro (SGTIN-96, same GCP, gtin14/serial not recorded) ────────

    [Fact]
    public void DefectRepro_Sgtin96_DecodesWithoutError()
    {
        // EPC from the AsReader TID defect repro script. The test-vector JSON listed
        // GCP "0696114" but bit extraction gives GCP 0720458 — the vector was wrong.
        // Assert that the decode succeeds cleanly with no ParseWarning.
        var result = EpcParser.ParseHex("30342BF92851DD10F36A0483");

        Assert.Equal(EpcScheme.Sgtin96, result.Scheme);
        Assert.Equal("0720458", result.CompanyPrefix);  // correct per bit extraction
        Assert.Null(result.ParseWarning);
        Assert.NotNull(result.Gtin14);
        Assert.Equal(14, result.Gtin14!.Length);
    }

    // ── Partition 4 (6-digit GCP) ─────────────────────────────────────────────

    [Fact]
    public void Partition5_Sgtin96_SecondVector_DecodesCorrectly()
    {
        // This EPC has header 0x30 and partition bits that resolve to partition 5 (L=7).
        // The test-vector file labelled it "gcp6-partition4" but the encoded bits
        // produce partition 5 — the vector description was incorrect.
        var result = EpcParser.ParseHex("3034257BF7194E4000001A85");

        Assert.Equal(EpcScheme.Sgtin96, result.Scheme);
        Assert.Equal(5, result.Partition);
        Assert.NotNull(result.CompanyPrefix);
        Assert.Equal(7, result.CompanyPrefix!.Length);  // partition 5 → 7-digit GCP
        Assert.NotNull(result.Gtin14);
        Assert.Equal(14, result.Gtin14!.Length);
    }

    // ── SSCC-96 (stub — currently returns Unknown) ────────────────────────────

    [Fact]
    public void Sscc96_ReturnsUnknownScheme_NoException()
    {
        // SSCC-96 (header 0x31) is Phase 3 — not yet decoded.
        // EpcParser must NOT throw; it must return a non-null ParsedEpc.
        var result = EpcParser.ParseHex("3114257BF71A05048000000C");

        Assert.NotNull(result);
        // Once SSCC support lands this assertion changes to EpcScheme.Sscc96
        Assert.Equal(EpcScheme.Unknown, result.Scheme);
    }

    // ── Unknown header ────────────────────────────────────────────────────────

    [Fact]
    public void UnknownHeader_ReturnsUnknown_NullGtin14()
    {
        var result = EpcParser.ParseHex("FF00000000000000000000FF");

        Assert.Equal(EpcScheme.Unknown, result.Scheme);
        Assert.Null(result.Gtin14);
    }

    // ── Malformed input ───────────────────────────────────────────────────────

    [Fact]
    public void TooShort_ReturnsUnknown_WithWarning()
    {
        // Single hex nibble — odd length, can't be decoded as bytes
        var result = EpcParser.ParseHex("3");

        Assert.Equal(EpcScheme.Unknown, result.Scheme);
        Assert.NotNull(result.ParseWarning);
    }

    [Fact]
    public void InvalidHex_ReturnsUnknown_WithWarning()
    {
        var result = EpcParser.ParseHex("GGGGGGGGGGGGGGGGGGGGGGGG");

        Assert.Equal(EpcScheme.Unknown, result.Scheme);
        Assert.NotNull(result.ParseWarning);
    }

    [Fact]
    public void NullInput_ReturnsUnknown_WithWarning()
    {
        var result = EpcParser.ParseHex(null);

        Assert.Equal(EpcScheme.Unknown, result.Scheme);
        Assert.NotNull(result.ParseWarning);
    }

    // ── GS1 GTIN-14 check digit validation ───────────────────────────────────
    // Vectors test the check digit embedded in Gtin14 returned by the parser.
    // Check digit rule: weight = 3 if (12-i)%2 == 0, else 1, for i=0..12.

    [Fact]
    public void Gtin14CheckDigit_ValidGtin_ParsedCorrectly()
    {
        // The live-A vector GTIN-14 "00696114704318" has a valid check digit (8).
        // Parsing the live-A EPC must produce exactly this GTIN-14.
        var result = EpcParser.ParseHex("30342A7CC844C7D0F36A0676");
        Assert.Equal("00696114704318", result.Gtin14);
        Assert.True(IsGtin14CheckDigitValid(result.Gtin14!));
    }

    [Fact]
    public void Gtin14CheckDigit_InvalidCheckDigit_Fails()
    {
        // "00696114704319" has wrong check digit (9 instead of 8)
        Assert.False(IsGtin14CheckDigitValid("00696114704319"));
    }

    [Fact]
    public void Gtin14CheckDigit_TooShort_Fails()
    {
        // Fewer than 14 digits — not a valid GTIN-14
        Assert.False(IsGtin14CheckDigitValid("069611470431"));
    }

    // ── Helper: mirrors Sgtin96Decoder.Gs1CheckDigit ─────────────────────────

    private static bool IsGtin14CheckDigitValid(string gtin14)
    {
        if (gtin14 is null || gtin14.Length != 14) return false;
        if (!gtin14.All(char.IsAsciiDigit)) return false;

        string digits13 = gtin14[..13];
        int sum = 0;
        for (int i = 0; i < 13; i++)
        {
            int weight = ((12 - i) % 2 == 0) ? 3 : 1;
            sum += (digits13[i] - '0') * weight;
        }
        char expected = (char)('0' + (10 - (sum % 10)) % 10);
        return gtin14[13] == expected;
    }
}
