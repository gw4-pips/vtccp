using InlineIo;
using InlineIo.Models;

namespace TestHarness.Fixtures;

/// <summary>
/// TestHarness fixture for the InlineIo relay assembly.
/// Exercises MockRelayBoard, IndicatorPoleController, and ConveyorInterruptController
/// without requiring any physical hardware.
/// </summary>
public static class InlineIoFixture
{
    public static async Task RunAsync()
    {
        Console.WriteLine();
        Console.WriteLine("════════════════════════════════════════════════════════════");
        Console.WriteLine("  InlineIo Fixture — MockRelayBoard / Indicator / Conveyor");
        Console.WriteLine("════════════════════════════════════════════════════════════");

        await using var board = new MockRelayBoard(relayCount: 8);
        await board.ConnectAsync();

        var map = RelayChannelMap.Default with
        {
            Red            = 1,
            Amber          = 2,
            Green          = 3,
            Blue           = -1,   // not wired
            ConveyorStop   = 4,
            ConveyorRestart = -1,  // single-channel mode (de-energise to resume)
            Buzzer         = -1,   // not wired
        };

        await using var indicator = new IndicatorPoleController(board, map)
        {
            FlashPeriodMs = 200, // faster for test visibility
        };
        var conveyor = new ConveyorInterruptController(board, map);

        // ── 1. Grade classification table ────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("1. Grade classification (no relay calls):");
        Console.WriteLine($"   {"Grade",-10} {"Colour",-10} {"Mode"}");
        Console.WriteLine($"   {new string('-', 30)}");
        foreach (decimal g in new[] { 0.0m, 1.5m, 1.8m, 2.0m, 2.3m, 2.4m, 2.8m, 2.9m, 3.4m, 3.5m, 4.0m })
        {
            var (c, m) = IndicatorPoleController.ClassifyGrade(g);
            Console.WriteLine($"   {g,-10:F1} {c,-10} {m}");
        }

        // ── 2. Steady GREEN ──────────────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("2. SetForGradeAsync(3.2) → GREEN STEADY …");
        var colour = await indicator.SetForGradeAsync(3.2m);
        Console.WriteLine($"   Active: {indicator.ActiveColour} / {indicator.ActiveMode}  (expected: Green/Steady) — {(colour == IndicatorColour.Green ? "PASS" : "FAIL")}");
        await Task.Delay(150);

        // ── 3. AMBER FLASH ───────────────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("3. SetForGradeAsync(2.5) → AMBER FLASH (3 toggles) …");
        await indicator.SetForGradeAsync(2.5m);
        Console.WriteLine($"   Active: {indicator.ActiveColour} / {indicator.ActiveMode}  (expected: Amber/Flash)  — {(indicator.ActiveColour == IndicatorColour.Amber && indicator.ActiveMode == IndicatorMode.Flash ? "PASS" : "FAIL")}");
        await Task.Delay(700); // let it flash a couple of times

        // ── 4. No-decode → RED FLASH + conveyor stop ─────────────────────────
        Console.WriteLine();
        Console.WriteLine("4. No-decode → RED FLASH + conveyor stop …");
        await indicator.SetForNoDecodeAsync();
        await conveyor.StopAsync();
        Console.WriteLine($"   Indicator: {indicator.ActiveColour}/{indicator.ActiveMode}  — {(indicator.ActiveColour == IndicatorColour.Red ? "PASS" : "FAIL")}");
        Console.WriteLine($"   Conveyor stopped: {conveyor.IsConveyorStopped}  — {(conveyor.IsConveyorStopped ? "PASS" : "FAIL")}");
        await Task.Delay(500);

        // ── 5. Clear + conveyor resume ───────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("5. ClearAsync + ResumeAsync …");
        await indicator.ClearAsync();
        await conveyor.ResumeAsync();
        Console.WriteLine($"   Indicator: {indicator.ActiveColour}  — {(indicator.ActiveColour == IndicatorColour.Off ? "PASS" : "FAIL")}");
        Console.WriteLine($"   Conveyor stopped: {conveyor.IsConveyorStopped}  — {(!conveyor.IsConveyorStopped ? "PASS" : "FAIL")}");

        // ── 6. RequiresConveyorStop classification ───────────────────────────
        Console.WriteLine();
        Console.WriteLine("6. RequiresConveyorStop:");
        foreach (var (grade, expected) in new (decimal?, bool)[]
            { (null, true), (1.5m, true), (1.8m, false), (3.0m, false) })
        {
            bool result = IndicatorPoleController.RequiresConveyorStop(grade);
            Console.WriteLine($"   grade={grade?.ToString() ?? "null",-6}  expected={expected}  got={result}  — {(result == expected ? "PASS" : "FAIL")}");
        }

        // ── Teardown ─────────────────────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("Teardown …");
        await conveyor.DisposeAsync();
        await board.DisconnectAsync();
        Console.WriteLine("InlineIo fixture complete.");
        Console.WriteLine();
    }
}
