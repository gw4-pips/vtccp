// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid;
using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;

namespace TestHarness.Fixtures;

/// <summary>
/// TestHarness Phase 7 — ASR-P35U hardware integration test.
///
/// Exercises the full TID read path that has been confirmed to work in isolation
/// but has never been run as part of the complete scan-flow sequence:
///   inventory (TriggerInventoryAsync) → TID read (ReadTidAsync) → validate (RfidScanCoordinator)
///
/// Run with: TestHarness --rfid-hw-test COM4
///
/// Prerequisites:
///   - ASR-P35U plugged into the specified COM port (VID=0x339C / PID=0x271B)
///   - AsReaderP3xU.dll present in vtccp\lib\asreader-p3xu-sdk-1.3.0\
///   - A GS1-tagged RFID item placed within ~5 cm of the antenna
///
/// Sub-tests:
///   7-A  Connect to the reader — ConnectAsync completes without exception
///   7-B  Inventory — at least one tag is detected within 5 s
///   7-C  TID read — ReadTidAsync returns a non-null, non-empty string within 2 s
///   7-D  TID format — result is valid uppercase hex, 8–32 characters (4–16 bytes)
///   7-E  TID MDID prefix — first 8 hex chars match a known chip family (advisory)
///   7-F  Full coordinator flow — RfidScanCoordinator delivers non-NoTag result
///        with Tid populated on the selected read
///   7-G  Disconnect — DisconnectAsync completes without exception
///
/// Each sub-test prints PASS or FAIL plus the observed value.  Exit code 0 = all
/// mandatory tests pass; non-zero if any mandatory test fails.
/// </summary>
public static class RfidHardwareFixture
{
    // ── Known MDID prefixes (first 8 hex chars = 4 bytes = TMN + MDID) ────────
    // Source: GS1 TID Memory Reference, Annex A; RAIN RFID Alliance tag list.
    // Add rows as new chip families are encountered in production.
    private static readonly IReadOnlyList<(string Prefix, string Description)> KnownMdidPrefixes =
    [
        ("E2801160", "Impinj Monza R6 / R6-P / R6-A"),
        ("E2801161", "Impinj Monza R6-B"),
        ("E2801170", "Impinj M730 / M750"),
        ("E2801171", "Impinj M775"),
        ("E2806890", "Alien Higgs-9"),
        ("E2806891", "Alien Higgs-EC"),
        ("E28068B0", "Alien Higgs-4"),
        ("E2800610", "NXP UCODE 8"),
        ("E2800611", "NXP UCODE 8m"),
        ("E2800690", "NXP UCODE 9"),
        ("E2003412", "Impinj M120"),
        ("E2800501", "EM Micro EM4325"),
        ("E2806800", "Alien ALN-9640"),
    ];

    // Inventory timeout for the hand-scan step.
    private static readonly TimeSpan InventoryTimeout = TimeSpan.FromSeconds(5);

    // TID read timeout — FW 1.8.0 callback fires within a few hundred ms when healthy.
    private static readonly TimeSpan TidTimeout = TimeSpan.FromMilliseconds(2000);

    // Coordinator scan-window (same window used in production).
    private const int CoordinatorWindowMs = 3000;

    /// <summary>
    /// Run the hardware integration test against the ASR-P35U on <paramref name="comPort"/>.
    /// Returns <c>true</c> when all mandatory tests pass.
    /// </summary>
    public static async Task<bool> RunAsync(string comPort)
    {
        Console.WriteLine();
        Console.WriteLine("════════════════════════════════════════════════════════════");
        Console.WriteLine($"  Phase 7 — ASR-P35U Hardware Integration Test  ({comPort})");
        Console.WriteLine("════════════════════════════════════════════════════════════");
        Console.WriteLine();
        Console.WriteLine("  PREREQUISITE: place a GS1-tagged item within ~5 cm of the antenna.");
        Console.WriteLine("  Press ENTER when ready …");
        Console.ReadLine();

        bool p7aPass = false, p7bPass = false, p7cPass = false,
             p7dPass = false,  p7fPass = false, p7gPass = false;
        bool p7eAdvisory = false;

        var reader = new AsReaderP35UEpcReader(txPowerDbm: 20);

        // ── 7-A: Connect ──────────────────────────────────────────────────────
        Console.WriteLine("7-A: Connect to ASR-P35U …");
        try
        {
            await reader.ConnectAsync(comPort);
            p7aPass = reader.IsConnected;
            Console.WriteLine($"  IsConnected={reader.IsConnected}: {(p7aPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  7-A EXCEPTION: {ex.Message}");
            Console.WriteLine("  Cannot continue without a connected reader.  Aborting Phase 7.");
            await reader.DisposeAsync();
            return false;
        }

        // ── 7-B: Inventory — tag detected ────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine($"7-B: TriggerInventoryAsync (timeout={InventoryTimeout.TotalSeconds:F0} s) …");
        IReadOnlyList<EpcReadResult> reads = [];
        try
        {
            reads = await reader.TriggerInventoryAsync(InventoryTimeout);
            p7bPass = reads.Count > 0;

            if (reads.Count > 0)
            {
                Console.WriteLine($"  Tags detected: {reads.Count}");
                Console.WriteLine($"  EPC[0]: {reads[0].EpcHex}");
                Console.WriteLine($"  RSSI[0]: {reads[0].Rssi?.ToString() ?? "(null)"} dBm");
                Console.WriteLine($"  TID[0] (from inventory): {reads[0].Tid ?? "(null — expected)"}");
            }
            else
            {
                Console.WriteLine("  No tags detected — is the item within range?");
            }
            Console.WriteLine($"  7-B: {(p7bPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  7-B EXCEPTION: {ex.Message}");
        }

        // ── 7-C: ReadTidAsync — TID returned ─────────────────────────────────
        Console.WriteLine();
        Console.WriteLine($"7-C: ReadTidAsync (timeout={TidTimeout.TotalMilliseconds:F0} ms) …");
        string? tid = null;
        if (!p7bPass)
        {
            Console.WriteLine("  SKIP — no tag from 7-B.");
        }
        else
        {
            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                tid = await reader.ReadTidAsync(reads[0].EpcBytes, TidTimeout);
                sw.Stop();

                p7cPass = !string.IsNullOrWhiteSpace(tid);
                Console.WriteLine($"  TID returned: {tid ?? "(null)"}");
                Console.WriteLine($"  Elapsed: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine($"  7-C: {(p7cPass ? "PASS" : "FAIL — null or empty TID (check FW 1.8.0 workaround / DLL)")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  7-C EXCEPTION: {ex.Message}");
            }
        }

        // ── 7-D: TID format — valid hex, 8–32 chars ───────────────────────────
        Console.WriteLine();
        Console.WriteLine("7-D: TID format — uppercase hex, 8–32 characters");
        if (!p7cPass || tid is null)
        {
            Console.WriteLine("  SKIP — no TID from 7-C.");
        }
        else
        {
            bool d1 = tid.Length is >= 8 and <= 32;
            bool d2 = tid.Length % 2 == 0;
            bool d3 = tid.All(c => c is >= '0' and <= '9' or >= 'A' and <= 'F');
            bool d4 = tid == tid.ToUpperInvariant();

            Console.WriteLine($"  Length={tid.Length} (8–32):    {(d1 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Even byte count:              {(d2 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Pure hex characters:          {(d3 ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Uppercase:                    {(d4 ? "PASS" : "FAIL")}");

            p7dPass = d1 && d2 && d3 && d4;
            Console.WriteLine($"  7-D: {(p7dPass ? "PASS" : "FAIL")}");
        }

        // ── 7-E: MDID prefix check (advisory) ────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("7-E: TID MDID prefix — known chip family check (advisory)");
        if (!p7cPass || tid is null)
        {
            Console.WriteLine("  SKIP — no TID from 7-C.");
        }
        else if (tid.Length < 8)
        {
            Console.WriteLine($"  SKIP — TID too short to extract MDID ({tid.Length} chars).");
        }
        else
        {
            string prefix8 = tid[..8];
            string prefix6 = tid[..6];

            var matched = KnownMdidPrefixes
                .FirstOrDefault(p =>
                    prefix8.StartsWith(p.Prefix, StringComparison.OrdinalIgnoreCase));

            if (matched.Description is not null)
            {
                p7eAdvisory = true;
                Console.WriteLine($"  Prefix {prefix8} matched: {matched.Description}");
                Console.WriteLine($"  7-E: ADVISORY PASS — chip family identified");
            }
            else
            {
                // Unknown prefix is not a failure — it may be a newer chip or
                // a private-range TID.  Log the prefix for the MDID table.
                Console.WriteLine($"  Prefix {prefix8} not in known MDID list.");
                Console.WriteLine($"  This is advisory only — check GS1 TID Memory Reference Annex A.");
                Console.WriteLine($"  To add: first 8 hex chars = {prefix8}");
                Console.WriteLine($"  7-E: ADVISORY UNKNOWN — update KnownMdidPrefixes if chip is confirmed");
            }
        }

        // ── 7-F: Full RfidScanCoordinator flow ───────────────────────────────
        Console.WriteLine();
        Console.WriteLine("7-F: Full RfidScanCoordinator flow — inventory → TID → validate …");
        Console.WriteLine("  Place the item near the antenna again. Press ENTER when ready …");
        Console.ReadLine();

        if (!p7aPass)
        {
            Console.WriteLine("  SKIP — reader not connected.");
        }
        else
        {
            try
            {
                // Build a synthetic barcode record — the coordinator validates RFID
                // against it.  Use EpcScheme.Unknown / empty GTIN to exercise the
                // NoTag / tag-found branch without needing a real matching barcode.
                var barcodeRecord = new VerificationRecord
                {
                    Symbology   = "GS1 DataMatrix",
                    DecodedData = "",   // no barcode data — coordinator will report NoGtin or Fail
                };

                var settings = new RfidScanCoordinatorSettings
                {
                    Enabled      = true,
                    ScanWindowMs = CoordinatorWindowMs,
                    FlagMismatchInReport = false,
                };

                var validator  = new RfidValidator(gcpValidator: null);

                // ownsReader=false so the coordinator does NOT dispose our reader at the end.
                var coordinator = new RfidScanCoordinator(reader, validator, settings, ownsReader: false);
                await using (coordinator)
                {
                    RfidValidationResult? result =
                        await coordinator.OnBarcodeScannedAsync(barcodeRecord);

                    if (result is null)
                    {
                        Console.WriteLine("  Coordinator returned null — disabled or concurrent scan.");
                        Console.WriteLine("  7-F: FAIL");
                    }
                    else
                    {
                        bool f1 = result.Status != RfidValidationStatus.NoTag;
                        bool f2 = result.SelectedRead is not null;
                        bool f3 = result.SelectedRead?.Tid is not null;

                        Console.WriteLine($"  Status: {result.Status}");
                        Console.WriteLine($"  SelectedRead != null: {result.SelectedRead is not null}");

                        if (result.SelectedRead is { } sel)
                        {
                            Console.WriteLine($"  EPC:  {sel.EpcHex}");
                            Console.WriteLine($"  TID:  {sel.Tid ?? "(null)"}");
                            Console.WriteLine($"  RSSI: {sel.Rssi?.ToString() ?? "(null)"} dBm");
                        }

                        Console.WriteLine($"  Status != NoTag:      {(f1 ? "PASS" : "FAIL — no tag detected")}");
                        Console.WriteLine($"  SelectedRead set:     {(f2 ? "PASS" : "FAIL")}");
                        Console.WriteLine($"  TID populated on read:{(f3 ? "PASS" : "FAIL — TID is null after coordinator")}");

                        p7fPass = f1 && f2 && f3;
                        Console.WriteLine($"  7-F: {(p7fPass ? "PASS" : "FAIL")}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  7-F EXCEPTION: {ex.Message}");
            }
        }

        // ── 7-G: Disconnect ───────────────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("7-G: DisconnectAsync …");
        try
        {
            await reader.DisconnectAsync();
            p7gPass = !reader.IsConnected;
            Console.WriteLine($"  IsConnected={reader.IsConnected} (expected false): {(p7gPass ? "PASS" : "FAIL")}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  7-G EXCEPTION: {ex.Message}");
        }
        finally
        {
            await reader.DisposeAsync();
        }

        // ── Phase 7 summary ───────────────────────────────────────────────────
        // 7-A, 7-B, 7-C, 7-D, 7-F, 7-G are mandatory.
        // 7-E is advisory (chip family lookup) and does not affect overall pass/fail.
        bool p7Pass = p7aPass && p7bPass && p7cPass && p7dPass && p7fPass && p7gPass;

        Console.WriteLine();
        Console.WriteLine("════════════════════════════════════════════════════════════");
        Console.WriteLine($"Phase 7 result: {(p7Pass ? "PASS" : "FAIL")}");
        if (!p7Pass)
        {
            if (!p7aPass) Console.WriteLine("  FAIL 7-A: Reader did not connect");
            if (!p7bPass) Console.WriteLine("  FAIL 7-B: No tag detected during inventory");
            if (!p7cPass) Console.WriteLine("  FAIL 7-C: TID not returned by ReadTidAsync");
            if (!p7dPass) Console.WriteLine("  FAIL 7-D: TID string failed format validation");
            if (!p7fPass) Console.WriteLine("  FAIL 7-F: Coordinator flow did not produce a TID-populated read");
            if (!p7gPass) Console.WriteLine("  FAIL 7-G: Reader did not disconnect cleanly");
        }
        Console.WriteLine($"7-E advisory: {(p7eAdvisory ? "chip family identified" : "prefix not in known list — see output above")}");
        Console.WriteLine("════════════════════════════════════════════════════════════");
        Console.WriteLine();
        Console.WriteLine("Recording instructions:");
        Console.WriteLine("  After a PASS run, note the following in the pre-customer checklist:");
        Console.WriteLine($"  - TID observed:  {tid ?? "(not captured)"}");
        Console.WriteLine($"  - MDID advisory: {(p7eAdvisory ? "chip family matched" : "unknown prefix")}");
        Console.WriteLine("  - All mandatory sub-tests: " + (p7Pass ? "PASS" : "FAIL"));

        return p7Pass;
    }
}
