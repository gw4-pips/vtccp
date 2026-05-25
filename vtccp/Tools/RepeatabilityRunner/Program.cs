// VTCCP RepeatabilityRunner — ad-hoc console tool
// Sends IMAGE.REPLAY N times on the currently-loaded device image and
// prints a per-parameter stability report.
//
// Usage (from vtccp/ directory):
//   dotnet run --project Tools/RepeatabilityRunner -- [host] [port] [reps]
//   dotnet run --project Tools/RepeatabilityRunner -- 10.10.10.7 44444 50
//
// Defaults: host=10.10.10.7  port=44444  reps=50
//
// Prerequisites:
//   1. Device must have an image in its buffer (prior IMAGE.LOAD or DMST load).
//   2. VTCCP.sln must be built at least once (dotnet build VTCCP.sln).
//   3. Run from the vtccp/ directory so project references resolve correctly.

using System.Diagnostics;
using DeviceInterface;
using DeviceInterface.Dmcc;
using ExcelEngine.Models;

const string DefaultHost = "10.10.10.7";
const int    DefaultPort = 44444;
const int    DefaultReps = 50;

string host = args.Length > 0 ? args[0] : DefaultHost;
int    port = args.Length > 1 && int.TryParse(args[1], out int p) ? p : DefaultPort;
int    reps = args.Length > 2 && int.TryParse(args[2], out int r) ? r : DefaultReps;

// ── Banner ────────────────────────────────────────────────────────────────────
Console.WriteLine("╔══════════════════════════════════════════════════════════════════╗");
Console.WriteLine("║  VTCCP Repeatability Runner                                      ║");
Console.WriteLine($"║  Target : {host}:{port}                                          ║");
Console.WriteLine($"║  Reps   : {reps}                                                    ║");
Console.WriteLine("╚══════════════════════════════════════════════════════════════════╝");
Console.WriteLine();

// ── Connect ───────────────────────────────────────────────────────────────────
var cfg = new DeviceConfig { Host = host, Port = port };
await using var session = new DeviceSession(cfg);

Console.Write("Connecting to device...");
try
{
    await session.ConnectAsync();
}
catch (Exception ex)
{
    Console.WriteLine($"  FAILED");
    Console.WriteLine($"  {ex.Message}");
    Console.WriteLine();
    Console.WriteLine("Is the device powered on and reachable? Check host/port and try again.");
    Environment.Exit(1);
    return;
}
Console.WriteLine("  OK");

var info = session.DeviceInfo;
Console.WriteLine($"  Model    : {info.Type ?? "unknown"}");
Console.WriteLine($"  Name     : {info.Name ?? "unknown"}");
Console.WriteLine($"  Serial   : {info.Serial ?? "unknown"}");
Console.WriteLine($"  Firmware : {info.FirmwareVersion ?? "unknown"}");
Console.WriteLine();
Console.WriteLine("NOTE: Device must already have an image loaded (prior IMAGE.LOAD or DMST scan).");
Console.WriteLine("      The DM rect image currently loaded will be replayed.");
Console.WriteLine();

// ── Progress header ───────────────────────────────────────────────────────────
string hdr = $"{"#",-5} {"Timestamp",-15} {"ms",-7} {"Formal Grade",-20} {"UEC%",-8} {"SC%",-8} {"ANU%",-9} {"GNU%",-9} {"FPD",-7} {"MOD",-5} {"RM",-5} {"DECODE",-7}";
Console.WriteLine(hdr);
Console.WriteLine(new string('─', hdr.Length + 5));

// ── Data collection ───────────────────────────────────────────────────────────
record RunData(int Rep, DateTime Timestamp, long ElapsedMs, VerificationRecord? Record);

var runs = new List<RunData>(reps);
var appCts = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) => { e.Cancel = true; appCts.Cancel(); };

for (int i = 1; i <= reps; i++)
{
    if (appCts.Token.IsCancellationRequested)
    {
        Console.WriteLine();
        Console.WriteLine($"  Cancelled at rep {i}/{reps}.");
        break;
    }

    var ts = DateTime.Now;
    var sw = Stopwatch.StartNew();
    VerificationRecord? rec = null;

    try
    {
        rec = await session.ReplayAndGetResultAsync(timeoutMs: 20_000, ct: appCts.Token);
    }
    catch (OperationCanceledException)
    {
        Console.WriteLine($"{i,-5} {"CANCELLED",-15}");
        break;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"{i,-5} {ts:HH:mm:ss.fff}   ERROR: {ex.Message}");
        runs.Add(new RunData(i, ts, sw.ElapsedMilliseconds, null));
        continue;
    }

    sw.Stop();
    runs.Add(new RunData(i, ts, sw.ElapsedMilliseconds, rec));

    if (rec is null)
    {
        Console.WriteLine($"{i,-5} {ts:HH:mm:ss.fff}   {sw.ElapsedMilliseconds,-7} {"TIMEOUT / NO READ",-20}");
        continue;
    }

    string formal  = rec.FormalGrade    ?? "—";
    string uecPct  = Fmt1(rec.UEC_Percent);
    string scPct   = Fmt1(rec.SC_Percent);
    string anuPct  = Fmt2(rec.ANU_Percent);
    string gnuPct  = Fmt2(rec.GNU_Percent);
    string fpd     = rec.FPD_Value.HasValue ? rec.FPD_Value.Value.ToString("F2") : "—";
    string mod     = FmtGrade(rec.MOD_Grade);
    string rm      = FmtGrade(rec.RM_Grade);
    string decode  = FmtGrade(rec.DECODE_Grade);

    Console.WriteLine($"{i,-5} {ts:HH:mm:ss.fff}   {sw.ElapsedMilliseconds,-7} {formal,-20} {uecPct,-8} {scPct,-8} {anuPct,-9} {gnuPct,-9} {fpd,-7} {mod,-5} {rm,-5} {decode,-7}");
}

Console.WriteLine(new string('─', hdr.Length + 5));
Console.WriteLine();

// ── Disconnect ────────────────────────────────────────────────────────────────
await session.DisconnectAsync();

// ── Timing summary ────────────────────────────────────────────────────────────
var good = runs.Where(r => r.Record is not null).ToList();
int  successCount = good.Count;
long avgMs        = successCount > 0 ? (long)good.Average(r => r.ElapsedMs) : 0;
long minMs        = successCount > 0 ? good.Min(r => r.ElapsedMs) : 0;
long maxMs        = successCount > 0 ? good.Max(r => r.ElapsedMs) : 0;
long totalMs      = successCount > 0 ? good.Sum(r => r.ElapsedMs) : 0;

Console.WriteLine("══ TIMING ══════════════════════════════════════════════════════════");
Console.WriteLine($"  Completed : {successCount}/{reps} reps");
Console.WriteLine($"  Total     : {totalMs:N0} ms ({totalMs / 1000.0:F1} s)");
Console.WriteLine($"  Per rep   : avg={avgMs} ms   min={minMs} ms   max={maxMs} ms");
Console.WriteLine();

// ── Deviation report ──────────────────────────────────────────────────────────
if (successCount == 0)
{
    Console.WriteLine("No successful reps — cannot generate deviation report.");
    Environment.Exit(2);
    return;
}

Console.WriteLine("══ DEVIATION REPORT ════════════════════════════════════════════════");
Console.WriteLine($"  Each parameter is checked across all {successCount} successful reps.");
Console.WriteLine($"  ★ = value changed at least once.  Values shown as: value×count");
Console.WriteLine();
Console.WriteLine($"  {"Parameter",-24} {"Status",-10} Values");
Console.WriteLine("  " + new string('─', 80));

var checks = new (string Label, Func<VerificationRecord, string?> Get)[]
{
    // Overall
    ("FormalGrade",          r => r.FormalGrade),
    ("OverallGrade",         r => FmtGrade(r.OverallGrade)),
    // Primary ISO 15415 / 29158 parameters
    ("UEC_Grade",            r => FmtGrade(r.UEC_Grade)),
    ("UEC_Percent",          r => r.UEC_Percent?.ToString("F4")),
    ("SC_Grade",             r => FmtGrade(r.SC_Grade)),
    ("SC_Percent",           r => r.SC_Percent?.ToString("F4")),
    ("MOD_Grade",            r => FmtGrade(r.MOD_Grade)),
    ("RM_Grade",             r => FmtGrade(r.RM_Grade)),
    ("ANU_Grade",            r => FmtGrade(r.ANU_Grade)),
    ("ANU_Percent",          r => r.ANU_Percent?.ToString("F6")),
    ("GNU_Grade",            r => FmtGrade(r.GNU_Grade)),
    ("GNU_Percent",          r => r.GNU_Percent?.ToString("F6")),
    ("FPD_Grade",            r => FmtGrade(r.FPD_Grade)),
    ("FPD_Value",            r => r.FPD_Value?.ToString("F6")),
    ("DECODE_Grade",         r => FmtGrade(r.DECODE_Grade)),
    // DM finder / quiet zone
    ("LLS_Grade",            r => FmtGrade(r.LLS_Grade)),
    ("BLS_Grade",            r => FmtGrade(r.BLS_Grade)),
    ("LQZ_Grade",            r => FmtGrade(r.LQZ_Grade)),
    ("BQZ_Grade",            r => FmtGrade(r.BQZ_Grade)),
    ("TQZ_Grade",            r => FmtGrade(r.TQZ_Grade)),
    ("RQZ_Grade",            r => FmtGrade(r.RQZ_Grade)),
    // Transition ratios & clock tracks
    ("TTR_Grade",            r => FmtGrade(r.TTR_Grade)),
    ("TTR_Percent",          r => r.TTR_Percent?.ToString("F6")),
    ("RTR_Grade",            r => FmtGrade(r.RTR_Grade)),
    ("RTR_Percent",          r => r.RTR_Percent?.ToString("F6")),
    ("TCT_Grade",            r => FmtGrade(r.TCT_Grade)),
    ("RCT_Grade",            r => FmtGrade(r.RCT_Grade)),
    // Codeword metadata
    ("ErrorsCorrected",      r => r.ErrorsCorrected?.ToString()),
    ("ErrorCapacityUsed",    r => r.ErrorCapacityUsed?.ToString()),
    ("DataCodewords",        r => r.DataCodewords?.ToString()),
    ("ErrorCorrectionBudget",r => r.ErrorCorrectionBudget?.ToString()),
};

bool anyDeviation = false;

foreach (var (label, get) in checks)
{
    var values  = good.Select(r => get(r.Record!) ?? "—").ToList();
    var distinct = values.GroupBy(v => v).OrderByDescending(g => g.Count()).ToList();
    bool stable  = distinct.Count <= 1;
    if (!stable) anyDeviation = true;

    string status = stable ? "STABLE" : "★ VARIES";
    string detail = stable
        ? (distinct.FirstOrDefault()?.Key ?? "—")
        : string.Join("  |  ", distinct.Select(g => $"{g.Key}×{g.Count()}"));

    string mark = stable ? "  " : "★ ";
    Console.WriteLine($"  {mark}{label,-24} {status,-10} {detail}");
}

Console.WriteLine();
if (!anyDeviation)
{
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"  ✓ ALL PARAMETERS STABLE across all {successCount} reps.");
    Console.WriteLine("    IMAGE.REPLAY grading is fully deterministic on this firmware.");
    Console.ResetColor();
}
else
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("  ★ DEVIATION DETECTED — see flagged parameters above.");
    Console.WriteLine("    Variation across IMAGE.REPLAY on identical pixels indicates");
    Console.WriteLine("    non-determinism in the firmware grading algorithm.");
    Console.ResetColor();
}

Console.WriteLine();
Console.WriteLine($"  Run started : {runs.First().Timestamp:yyyy-MM-dd HH:mm:ss.fff}");
Console.WriteLine($"  Run ended   : {runs.Last().Timestamp:yyyy-MM-dd HH:mm:ss.fff}");
Console.WriteLine();

// ── Helpers ───────────────────────────────────────────────────────────────────
static string FmtGrade(GradingResult? g) =>
    g is null ? "—" : $"{g.Letter}/{g.NumericGrade:F1}";

static string Fmt1(decimal? v) => v.HasValue ? $"{v.Value:F1}%" : "—";
static string Fmt2(decimal? v) => v.HasValue ? $"{v.Value:F2}%" : "—";
