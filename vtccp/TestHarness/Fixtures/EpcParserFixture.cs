// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid;
using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;

namespace TestHarness.Fixtures;

/// <summary>
/// TestHarness Phase 6 fixture — EPC parser, scheme dispatch, GCP validation,
/// partition table, and RFID cross-validator (Phases 0.2–0.4).
///
/// Sub-tests:
///   6-A  SGTIN-96 decode — partition 5 (7-digit GCP)
///   6-B  SGTIN-96 decode — partition 0 (12-digit GCP, maximum)
///   6-C  SGTIN-96 decode — partition 6 (6-digit GCP, minimum)
///   6-D  PartitionTable — all 7 rows: L+K=13 and M+N=44 invariants
///   6-E  EpcParser.ParseHex — hex string convenience overload
///   6-F  EpcParser.ParseHex — unknown header returns EpcScheme.Unknown
///   6-G  EpcParser.ParseHex — null/empty string handled gracefully
///   6-H  GcpLengthTable — load from bundled XML file
///   6-I  GcpValidator — GCP match detected
///   6-J  GcpValidator — GCP length mismatch detected
///   6-K  RfidValidator — GTIN match + serial match (Pass result)
///   6-L  RfidValidator — GTIN mismatch (Fail result)
///   6-M  RfidValidator — no tag detected (NoTag result)
/// </summary>
public static class EpcParserFixture
{
    // ─── Test vectors ──────────────────────────────────────────────────────────
    //
    // All vectors verified by hand against GS1 EPC TDS 2.3 Table 14-1 and
    // the GTIN-14 reconstruction formula (confirmed in .agents/memory/sgtin96-gtin14-formula.md).
    //
    // Vector A — partition 5 (M=24 L=7  N=20 K=6)
    //   Filter=1 GCP="0614141" ItemRef="012345" Serial="100"
    //   payload13="0614141012345"  GTIN-14="06141410123454"
    //   Bit stream → 30 34 25 7B F4 0C 0E 40 00 00 00 64
    private const string VecA_Hex     = "3034257BF40C0E4000000064";
    private const string VecA_Gcp     = "0614141";
    private const string VecA_ItemRef = "012345";
    private const string VecA_Serial  = "100";
    private const string VecA_Gtin14  = "06141410123454";
    private const int    VecA_Filter  = 1;
    private const int    VecA_Partition = 5;

    // Vector B — partition 0 (M=40 L=12  N=4 K=1)
    //   Filter=0 GCP="000000000001" ItemRef="5" Serial="0"
    //   payload13="0000000000015"  check = (10 - (16%10))%10 = 4
    //   GTIN-14="00000000000154"
    //   Bit stream → 30 00 00 00 00 00 05 40 00 00 00 00
    private const string VecB_Hex     = "300000000000054000000000";
    private const string VecB_Gcp     = "000000000001";
    private const string VecB_ItemRef = "5";
    private const string VecB_Serial  = "0";
    private const string VecB_Gtin14  = "00000000000154";
    private const int    VecB_Filter    = 0;
    private const int    VecB_Partition = 0;

    // Vector C — partition 6 (M=20 L=6  N=24 K=7)
    //   Filter=0 GCP="123456" ItemRef="1234567" Serial="1"
    //   payload13="1234561234567"  check = (10 - (99%10))%10 = 1
    //   GTIN-14="12345612345671"
    //   Bit stream → 30 18 78 90 04 B5 A1 C0 00 00 00 01
    private const string VecC_Hex     = "3018789004B5A1C000000001";
    private const string VecC_Gcp     = "123456";
    private const string VecC_ItemRef = "1234567";
    private const string VecC_Serial  = "1";
    private const string VecC_Gtin14  = "12345612345671";
    private const int    VecC_Filter    = 0;
    private const int    VecC_Partition = 6;

    // ─────────────────────────────────────────────────────────────────────────

    public static Task<bool> RunAsync()
    {
        Console.WriteLine();
        Console.WriteLine("════════════════════════════════════════════════════════════");
        Console.WriteLine("  Phase 6 — EPC Parser / Partition Table / GCP Validation");
        Console.WriteLine("════════════════════════════════════════════════════════════");

        bool p6aPass = false, p6bPass = false, p6cPass = false,
             p6dPass = false, p6ePass = false, p6fPass = false,
             p6gPass = false, p6hPass = false, p6iPass = false,
             p6jPass = false, p6kPass = false, p6lPass = false,
             p6mPass = false;

        // ── 6-A: SGTIN-96 partition 5 ─────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-A: SGTIN-96 decode — partition 5 (7-digit GCP)");
        try
        {
            var bytes = Convert.FromHexString(VecA_Hex);
            var parsed = EpcParser.Parse(bytes);

            bool a1 = parsed.Scheme        == EpcScheme.Sgtin96;
            bool a2 = parsed.Filter        == VecA_Filter;
            bool a3 = parsed.Partition     == VecA_Partition;
            bool a4 = parsed.CompanyPrefix == VecA_Gcp;
            bool a5 = parsed.ItemReference == VecA_ItemRef;
            bool a6 = parsed.Serial        == VecA_Serial;
            bool a7 = parsed.Gtin14        == VecA_Gtin14;
            bool a8 = parsed.ParseWarning  is null;

            Console.WriteLine($"  Scheme=Sgtin96:         {(a1 ? "PASS" : $"FAIL ({parsed.Scheme})")}");
            Console.WriteLine($"  Filter=1:               {(a2 ? "PASS" : $"FAIL ({parsed.Filter})")}");
            Console.WriteLine($"  Partition=5:            {(a3 ? "PASS" : $"FAIL ({parsed.Partition})")}");
            Console.WriteLine($"  GCP=\"{VecA_Gcp}\":    {(a4 ? "PASS" : $"FAIL (\"{parsed.CompanyPrefix}\")")}");
            Console.WriteLine($"  ItemRef=\"{VecA_ItemRef}\":   {(a5 ? "PASS" : $"FAIL (\"{parsed.ItemReference}\")")}");
            Console.WriteLine($"  Serial=\"{VecA_Serial}\":        {(a6 ? "PASS" : $"FAIL (\"{parsed.Serial}\")")}");
            Console.WriteLine($"  GTIN14=\"{VecA_Gtin14}\": {(a7 ? "PASS" : $"FAIL (\"{parsed.Gtin14}\")")}");
            Console.WriteLine($"  NoWarning:              {(a8 ? "PASS" : $"FAIL (warning: {parsed.ParseWarning})")}");

            p6aPass = a1 && a2 && a3 && a4 && a5 && a6 && a7 && a8;
            Console.WriteLine($"  6-A: {(p6aPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-A EXCEPTION: {ex.Message}"); }

        // ── 6-B: SGTIN-96 partition 0 ─────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-B: SGTIN-96 decode — partition 0 (12-digit GCP, max)");
        try
        {
            var bytes = Convert.FromHexString(VecB_Hex);
            var parsed = EpcParser.Parse(bytes);

            bool b1 = parsed.Scheme        == EpcScheme.Sgtin96;
            bool b2 = parsed.Partition     == VecB_Partition;
            bool b3 = parsed.CompanyPrefix == VecB_Gcp;
            bool b4 = parsed.ItemReference == VecB_ItemRef;
            bool b5 = parsed.Serial        == VecB_Serial;
            bool b6 = parsed.Gtin14        == VecB_Gtin14;

            Console.WriteLine($"  Scheme=Sgtin96:             {(b1 ? "PASS" : $"FAIL ({parsed.Scheme})")}");
            Console.WriteLine($"  Partition=0:                {(b2 ? "PASS" : $"FAIL ({parsed.Partition})")}");
            Console.WriteLine($"  GCP=\"{VecB_Gcp}\": {(b3 ? "PASS" : $"FAIL (\"{parsed.CompanyPrefix}\")")}");
            Console.WriteLine($"  ItemRef=\"{VecB_ItemRef}\":              {(b4 ? "PASS" : $"FAIL (\"{parsed.ItemReference}\")")}");
            Console.WriteLine($"  Serial=\"{VecB_Serial}\":               {(b5 ? "PASS" : $"FAIL (\"{parsed.Serial}\")")}");
            Console.WriteLine($"  GTIN14=\"{VecB_Gtin14}\":     {(b6 ? "PASS" : $"FAIL (\"{parsed.Gtin14}\")")}");

            p6bPass = b1 && b2 && b3 && b4 && b5 && b6;
            Console.WriteLine($"  6-B: {(p6bPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-B EXCEPTION: {ex.Message}"); }

        // ── 6-C: SGTIN-96 partition 6 ─────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-C: SGTIN-96 decode — partition 6 (6-digit GCP, min)");
        try
        {
            var bytes = Convert.FromHexString(VecC_Hex);
            var parsed = EpcParser.Parse(bytes);

            bool c1 = parsed.Scheme        == EpcScheme.Sgtin96;
            bool c2 = parsed.Partition     == VecC_Partition;
            bool c3 = parsed.CompanyPrefix == VecC_Gcp;
            bool c4 = parsed.ItemReference == VecC_ItemRef;
            bool c5 = parsed.Serial        == VecC_Serial;
            bool c6 = parsed.Gtin14        == VecC_Gtin14;

            Console.WriteLine($"  Scheme=Sgtin96:       {(c1 ? "PASS" : $"FAIL ({parsed.Scheme})")}");
            Console.WriteLine($"  Partition=6:          {(c2 ? "PASS" : $"FAIL ({parsed.Partition})")}");
            Console.WriteLine($"  GCP=\"{VecC_Gcp}\":     {(c3 ? "PASS" : $"FAIL (\"{parsed.CompanyPrefix}\")")}");
            Console.WriteLine($"  ItemRef=\"{VecC_ItemRef}\": {(c4 ? "PASS" : $"FAIL (\"{parsed.ItemReference}\")")}");
            Console.WriteLine($"  Serial=\"{VecC_Serial}\":         {(c5 ? "PASS" : $"FAIL (\"{parsed.Serial}\")")}");
            Console.WriteLine($"  GTIN14=\"{VecC_Gtin14}\": {(c6 ? "PASS" : $"FAIL (\"{parsed.Gtin14}\")")}");

            p6cPass = c1 && c2 && c3 && c4 && c5 && c6;
            Console.WriteLine($"  6-C: {(p6cPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-C EXCEPTION: {ex.Message}"); }

        // ── 6-D: PartitionTable — invariants ─────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-D: PartitionTable — L+K=13 and M+N=44 for all 7 rows");
        try
        {
            bool allInvariantsOk = true;
            for (int p = 0; p <= 6; p++)
            {
                var row = PartitionTable.Get(p);
                bool lkOk = row.L + row.K == 13;
                bool mnOk = row.M + row.N == 44;

                Console.WriteLine(
                    $"  P={p}: M={row.M,2} L={row.L,2} N={row.N,2} K={row.K} " +
                    $"L+K={row.L + row.K} {(lkOk ? "✓" : "✗")}  " +
                    $"M+N={row.M + row.N} {(mnOk ? "✓" : "✗")}  " +
                    $"{(lkOk && mnOk ? "PASS" : "FAIL")}");

                allInvariantsOk &= lkOk && mnOk;
            }

            // Also verify PartitionTable.GcpDigits / ItemRefDigits helpers
            bool helperGcp     = PartitionTable.GcpDigits(5)     == 7;
            bool helperItemRef = PartitionTable.ItemRefDigits(5)  == 6;
            bool helperOob     = PartitionTable.GcpDigits(7)      == -1;
            bool helperTryGet  = PartitionTable.TryGet(3, out var row3) && row3.M == 30;
            bool helperTryOob  = !PartitionTable.TryGet(7, out _);

            Console.WriteLine($"  GcpDigits(5)==7:     {(helperGcp     ? "PASS" : $"FAIL ({PartitionTable.GcpDigits(5)})")}");
            Console.WriteLine($"  ItemRefDigits(5)==6: {(helperItemRef  ? "PASS" : $"FAIL ({PartitionTable.ItemRefDigits(5)})")}");
            Console.WriteLine($"  GcpDigits(7)==-1:    {(helperOob      ? "PASS" : $"FAIL ({PartitionTable.GcpDigits(7)})")}");
            Console.WriteLine($"  TryGet(3).M==30:     {(helperTryGet   ? "PASS" : $"FAIL (M={row3.M})")}");
            Console.WriteLine($"  TryGet(7)==false:    {(helperTryOob   ? "PASS" : "FAIL")}");

            p6dPass = allInvariantsOk && helperGcp && helperItemRef
                   && helperOob && helperTryGet && helperTryOob;
            Console.WriteLine($"  6-D: {(p6dPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-D EXCEPTION: {ex.Message}"); }

        // ── 6-E: EpcParser.ParseHex — hex string overload ─────────────────────
        Console.WriteLine();
        Console.WriteLine("6-E: EpcParser.ParseHex — hex string convenience overload");
        try
        {
            // Lower-case and whitespace should both be tolerated
            var parsed1 = EpcParser.ParseHex(VecA_Hex.ToLowerInvariant());
            var parsed2 = EpcParser.ParseHex("3034 257B F40C 0E40 0000 0064");

            bool e1 = parsed1.Scheme  == EpcScheme.Sgtin96 && parsed1.Gtin14 == VecA_Gtin14;
            bool e2 = parsed2.Scheme  == EpcScheme.Sgtin96 && parsed2.Gtin14 == VecA_Gtin14;
            bool e3 = EpcParser.ParseHex(null).Scheme == EpcScheme.Unknown;
            bool e4 = EpcParser.ParseHex("ZZZZ").Scheme == EpcScheme.Unknown;

            Console.WriteLine($"  Lower-case hex → Sgtin96+GTIN14: {(e1 ? "PASS" : $"FAIL (scheme={parsed1.Scheme} gtin14={parsed1.Gtin14})")}");
            Console.WriteLine($"  Space-padded hex → Sgtin96+GTIN14:{(e2 ? "PASS" : $"FAIL (scheme={parsed2.Scheme} gtin14={parsed2.Gtin14})")}");
            Console.WriteLine($"  null → Unknown:                   {(e3 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  \"ZZZZ\" → Unknown:                 {(e4 ? "PASS" : "FAIL")}");

            p6ePass = e1 && e2 && e3 && e4;
            Console.WriteLine($"  6-E: {(p6ePass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-E EXCEPTION: {ex.Message}"); }

        // ── 6-F: Unknown header returns EpcScheme.Unknown ──────────────────────
        Console.WriteLine();
        Console.WriteLine("6-F: EpcParser.ParseHex — unknown header 0xFF");
        try
        {
            // 12 bytes with header 0xFF — not a valid TDS scheme
            var parsed = EpcParser.ParseHex("FF0000000000000000000000");

            bool f1 = parsed.Scheme == EpcScheme.Unknown;
            bool f2 = parsed.ParseWarning is not null;
            bool f3 = parsed.ParseWarning!.Contains("FF", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"  Scheme=Unknown:          {(f1 ? "PASS" : $"FAIL ({parsed.Scheme})")}");
            Console.WriteLine($"  ParseWarning set:        {(f2 ? "PASS" : "FAIL (null)")}");
            Console.WriteLine($"  Warning contains \"FF\":   {(f3 ? "PASS" : $"FAIL (\"{parsed.ParseWarning}\")")}");

            p6fPass = f1 && f2 && f3;
            Console.WriteLine($"  6-F: {(p6fPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-F EXCEPTION: {ex.Message}"); }

        // ── 6-G: Empty / null input handled ───────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-G: EpcParser — null/empty byte array and empty hex");
        try
        {
            var g1Parsed = EpcParser.Parse([]);
            var g2Parsed = EpcParser.Parse(null!);   // null is defensive
            var g3Parsed = EpcParser.ParseHex("");

            bool g1 = g1Parsed.Scheme == EpcScheme.Unknown && g1Parsed.ParseWarning is not null;
            bool g2 = g2Parsed.Scheme == EpcScheme.Unknown && g2Parsed.ParseWarning is not null;
            bool g3 = g3Parsed.Scheme == EpcScheme.Unknown && g3Parsed.ParseWarning is not null;

            Console.WriteLine($"  Parse([]) → Unknown+Warning:     {(g1 ? "PASS" : $"FAIL (scheme={g1Parsed.Scheme} warn={g1Parsed.ParseWarning})")}");
            Console.WriteLine($"  Parse(null) → Unknown+Warning:   {(g2 ? "PASS" : $"FAIL (scheme={g2Parsed.Scheme} warn={g2Parsed.ParseWarning})")}");
            Console.WriteLine($"  ParseHex(\"\") → Unknown+Warning:  {(g3 ? "PASS" : $"FAIL (scheme={g3Parsed.Scheme} warn={g3Parsed.ParseWarning})")}");

            p6gPass = g1 && g2 && g3;
            Console.WriteLine($"  6-G: {(p6gPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-G EXCEPTION: {ex.Message}"); }

        // ── 6-H: GcpLengthTable — load from bundled XML ───────────────────────
        Console.WriteLine();
        Console.WriteLine("6-H: GcpLengthTable — load bundled gcp-prefix-format-list.xml");

        GcpLengthTable? gcpTable = null;
        try
        {
            // Locate the bundled data file relative to the build output directory.
            // Output dir is: vtccp/TestHarness/bin/{Config}/net8.0/ → up 4 = vtccp/
            string xmlPath = Path.GetFullPath(
                Path.Combine(AppContext.BaseDirectory,
                    "..", "..", "..", "..", "data", "gcp-prefix-format-list.xml"));

            gcpTable = GcpLengthTable.LoadFromFile(xmlPath);

            bool h1 = gcpTable.EntryCount > 100_000;        // 200,108 entries in the 2026-05-03 file
            bool h2 = gcpTable.DataDate.HasValue;
            bool h3 = gcpTable.DataDate?.Year is 2026;      // date="2026-05-03T..."

            // Spot-check a known prefix: "006" → gcpLength=7
            bool h4 = gcpTable.TryLookup("006", out int lookedUp) && lookedUp == 7;

            // gcpLength=0 entries (restricted prefixes 020-029) should be present
            bool h5 = gcpTable.TryLookup("020", out int restricted) && restricted == 0;

            Console.WriteLine($"  EntryCount>100000 ({gcpTable.EntryCount}): {(h1 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  DataDate set:                    {(h2 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  DataDate year=2026:              {(h3 ? "PASS" : $"FAIL ({gcpTable.DataDate})")}");
            Console.WriteLine($"  TryLookup(\"006\")==7:            {(h4 ? "PASS" : $"FAIL ({lookedUp})")}");
            Console.WriteLine($"  TryLookup(\"020\")==0 (restricted):{(h5 ? "PASS" : $"FAIL ({restricted})")}");

            p6hPass = h1 && h2 && h3 && h4 && h5;
            Console.WriteLine($"  6-H: {(p6hPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-H EXCEPTION: {ex.Message}"); }

        // ── 6-I: GcpValidator — match ─────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-I: GcpValidator — GCP match (prefix \"006\" → gcpLength=7)");
        try
        {
            if (gcpTable is null)
            {
                Console.WriteLine("  SKIP — GcpLengthTable not loaded (6-H failed).");
            }
            else
            {
                var validator = new GcpValidator(gcpTable);

                // GCP "0061234" (7 digits) — TryLookup will find prefix "006" → gcpLength=7.
                // Partition 5 implies L=7 → claimed length matches registered length → Match.
                var matchEpc = new ParsedEpc
                {
                    EpcBytes      = [],
                    Scheme        = EpcScheme.Sgtin96,
                    CompanyPrefix = "0061234",
                    Partition     = 5,    // L=7
                };

                bool? result  = validator.Validate(matchEpc);
                bool i1       = result == true;

                // ValidateRaw: the 7-char GCP should validate correctly
                bool i2 = validator.ValidateRaw("0061234");

                // ValidateRaw with too-long string (no match expected for 10-char entry)
                bool i3 = !validator.ValidateRaw("0061234999");

                Console.WriteLine($"  Validate(matchEpc)==true:    {(i1 ? "PASS" : $"FAIL ({result})")}");
                Console.WriteLine($"  ValidateRaw(\"0061234\")==true: {(i2 ? "PASS" : "FAIL")}");
                Console.WriteLine($"  ValidateRaw(10-char)==false: {(i3 ? "PASS" : "FAIL")}");

                p6iPass = i1 && i2 && i3;
                Console.WriteLine($"  6-I: {(p6iPass ? "PASS" : "FAIL")}");
            }
        }
        catch (Exception ex) { Console.WriteLine($"  6-I EXCEPTION: {ex.Message}"); }

        // ── 6-J: GcpValidator — mismatch ─────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-J: GcpValidator — GCP length mismatch detected");
        try
        {
            if (gcpTable is null)
            {
                Console.WriteLine("  SKIP — GcpLengthTable not loaded (6-H failed).");
            }
            else
            {
                var validator = new GcpValidator(gcpTable);

                // GCP "030001" (6 chars): TryLookup finds prefix "03000" → gcpLength=6.
                // But Partition=5 implies L=7 → claimed length 7 ≠ registered length 6 → Mismatch.
                var mismatchEpc = new ParsedEpc
                {
                    EpcBytes      = [],
                    Scheme        = EpcScheme.Sgtin96,
                    CompanyPrefix = "030001",
                    Partition     = 5,    // L=7, but this prefix is registered as length 6
                };

                bool? result = validator.Validate(mismatchEpc);
                bool j1      = result == false;

                // ValidateRaw confirms: "030001" (6-char) should be valid as a 6-char GCP
                bool j2 = validator.ValidateRaw("030001");

                Console.WriteLine($"  Validate(mismatchEpc)==false: {(j1 ? "PASS" : $"FAIL ({result})")}");
                Console.WriteLine($"  ValidateRaw(\"030001\")==true:   {(j2 ? "PASS" : "FAIL")} (6-char is correct)");

                p6jPass = j1 && j2;
                Console.WriteLine($"  6-J: {(p6jPass ? "PASS" : "FAIL")}");
            }
        }
        catch (Exception ex) { Console.WriteLine($"  6-J EXCEPTION: {ex.Message}"); }

        // ── 6-K: RfidValidator — GTIN match + serial match ────────────────────
        Console.WriteLine();
        Console.WriteLine("6-K: RfidValidator — GTIN match + serial match (Pass)");
        try
        {
            var validator = new RfidValidator(gcpValidator: null);

            var reads = new[]
            {
                new EpcReadResult
                {
                    EpcBytes = Convert.FromHexString(VecA_Hex),
                    ReadTime = DateTimeOffset.UtcNow,
                },
            };

            // GS1 DataMatrix payload: AI(01) = VecA_Gtin14, AI(21) = VecA_Serial.
            // FNC1 separator between AI groups represented as 0x1D.
            string barcodeData = $"]d201{VecA_Gtin14}\x1D21{VecA_Serial}";

            var barcodeRecord = new VerificationRecord
            {
                Symbology   = "GS1 DataMatrix",
                DecodedData = barcodeData,
            };

            var result = validator.Validate(reads, barcodeRecord, scanWindowMs: 1000);

            bool k1 = result.Status      == RfidValidationStatus.Pass;
            bool k2 = result.Gtin14Matches;
            bool k3 = result.SerialMatches;
            bool k4 = result.RfidGtin14  == VecA_Gtin14;
            bool k5 = result.BarcodeGtin14 == VecA_Gtin14;
            bool k6 = result.RfidSerial  == VecA_Serial;
            bool k7 = result.BarcodeSerial == VecA_Serial;
            bool k8 = result.MismatchDetail is null;

            Console.WriteLine($"  Status=Pass:             {(k1 ? "PASS" : $"FAIL ({result.Status})")}");
            Console.WriteLine($"  Gtin14Matches:           {(k2 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  SerialMatches:           {(k3 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  RfidGtin14={VecA_Gtin14}:{(k4 ? "PASS" : $"FAIL ({result.RfidGtin14})")}");
            Console.WriteLine($"  BarcodeGtin14 same:      {(k5 ? "PASS" : $"FAIL ({result.BarcodeGtin14})")}");
            Console.WriteLine($"  RfidSerial={VecA_Serial}:         {(k6 ? "PASS" : $"FAIL ({result.RfidSerial})")}");
            Console.WriteLine($"  BarcodeSerial same:      {(k7 ? "PASS" : $"FAIL ({result.BarcodeSerial})")}");
            Console.WriteLine($"  MismatchDetail=null:     {(k8 ? "PASS" : $"FAIL ({result.MismatchDetail})")}");

            p6kPass = k1 && k2 && k3 && k4 && k5 && k6 && k7 && k8;
            Console.WriteLine($"  6-K: {(p6kPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-K EXCEPTION: {ex.Message}"); }

        // ── 6-L: RfidValidator — GTIN mismatch ────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-L: RfidValidator — GTIN mismatch (Fail)");
        try
        {
            var validator = new RfidValidator(gcpValidator: null);

            var reads = new[]
            {
                new EpcReadResult
                {
                    EpcBytes = Convert.FromHexString(VecA_Hex),
                    ReadTime = DateTimeOffset.UtcNow,
                },
            };

            // Barcode carries a different GTIN14 (last digit off by 1)
            string wrongGtin14  = VecA_Gtin14[..13] + "9";   // "06141410123459" — bad check digit but tests mismatch
            string barcodeData  = $"]d201{wrongGtin14}\x1D21{VecA_Serial}";

            var barcodeRecord = new VerificationRecord
            {
                Symbology   = "GS1 DataMatrix",
                DecodedData = barcodeData,
            };

            var result = validator.Validate(reads, barcodeRecord, scanWindowMs: 1000);

            bool l1 = result.Status == RfidValidationStatus.Fail;
            bool l2 = !result.Gtin14Matches;
            bool l3 = result.MismatchDetail is not null;
            bool l4 = result.MismatchDetail?.Contains("GTIN14") == true;

            Console.WriteLine($"  Status=Fail:             {(l1 ? "PASS" : $"FAIL ({result.Status})")}");
            Console.WriteLine($"  Gtin14Matches=false:     {(l2 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  MismatchDetail set:      {(l3 ? "PASS" : "FAIL (null)")}");
            Console.WriteLine($"  Detail contains GTIN14:  {(l4 ? "PASS" : $"FAIL (\"{result.MismatchDetail}\")")}");

            p6lPass = l1 && l2 && l3 && l4;
            Console.WriteLine($"  6-L: {(p6lPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-L EXCEPTION: {ex.Message}"); }

        // ── 6-M: RfidValidator — no tag ───────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("6-M: RfidValidator — no tag detected (NoTag)");
        try
        {
            var validator = new RfidValidator(gcpValidator: null);

            var barcodeRecord = new VerificationRecord
            {
                Symbology   = "GS1 DataMatrix",
                DecodedData = $"]d201{VecA_Gtin14}\x1D21{VecA_Serial}",
            };

            var result = validator.Validate(
                reads:          Array.Empty<EpcReadResult>(),
                barcodeRecord:  barcodeRecord,
                scanWindowMs:   500);

            bool m1 = result.Status == RfidValidationStatus.NoTag;
            bool m2 = result.RawReads.Count == 0;
            bool m3 = result.SelectedRead is null;

            Console.WriteLine($"  Status=NoTag:    {(m1 ? "PASS" : $"FAIL ({result.Status})")}");
            Console.WriteLine($"  RawReads empty:  {(m2 ? "PASS" : $"FAIL ({result.RawReads.Count})")}");
            Console.WriteLine($"  SelectedRead=null:{(m3 ? "PASS" : "FAIL")}");

            p6mPass = m1 && m2 && m3;
            Console.WriteLine($"  6-M: {(p6mPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex) { Console.WriteLine($"  6-M EXCEPTION: {ex.Message}"); }

        // ── Phase 6 summary ───────────────────────────────────────────────────
        bool p6Pass = p6aPass && p6bPass && p6cPass && p6dPass
                   && p6ePass && p6fPass && p6gPass && p6hPass
                   && p6iPass && p6jPass && p6kPass && p6lPass && p6mPass;

        Console.WriteLine();
        Console.WriteLine($"Phase 6 verification: {(p6Pass ? "PASS" : "FAIL")}");
        if (!p6Pass)
        {
            if (!p6aPass) Console.WriteLine("  FAIL: 6-A SGTIN-96 partition 5");
            if (!p6bPass) Console.WriteLine("  FAIL: 6-B SGTIN-96 partition 0");
            if (!p6cPass) Console.WriteLine("  FAIL: 6-C SGTIN-96 partition 6");
            if (!p6dPass) Console.WriteLine("  FAIL: 6-D PartitionTable invariants");
            if (!p6ePass) Console.WriteLine("  FAIL: 6-E ParseHex convenience overload");
            if (!p6fPass) Console.WriteLine("  FAIL: 6-F Unknown header → Unknown");
            if (!p6gPass) Console.WriteLine("  FAIL: 6-G Empty/null input handled");
            if (!p6hPass) Console.WriteLine("  FAIL: 6-H GcpLengthTable load");
            if (!p6iPass) Console.WriteLine("  FAIL: 6-I GcpValidator match");
            if (!p6jPass) Console.WriteLine("  FAIL: 6-J GcpValidator mismatch");
            if (!p6kPass) Console.WriteLine("  FAIL: 6-K RfidValidator GTIN+serial match");
            if (!p6lPass) Console.WriteLine("  FAIL: 6-L RfidValidator GTIN mismatch");
            if (!p6mPass) Console.WriteLine("  FAIL: 6-M RfidValidator no tag");
        }
        Console.WriteLine();
        Console.WriteLine("Phase 6 complete.");

        return Task.FromResult(p6Pass);
    }
}
