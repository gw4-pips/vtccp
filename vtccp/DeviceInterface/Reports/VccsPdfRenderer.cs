// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  VCCS PDF renderer — WebView2 (Edge) primary, wkhtmltopdf silent fallback ║
// ║                                                                          ║
// ║  Renders the v23 HTML produced by VccsHtmlReportGenerator to a PDF file. ║
// ║  Path selection is fully transparent to the caller:                      ║
// ║    1. WebView2 runtime present  → CoreWebView2.PrintToPdfAsync           ║
// ║       (ships with Edge — present on Win11 and most updated Win10)        ║
// ║    2. Otherwise                 → bundled wkhtmltopdf.exe                ║
// ║       (<ExeDir>/resources/wkhtmltopdf.exe, Process.Start, no window)     ║
// ║  Detection is silent; failures are logged to Debug output only.          ║
// ╚══════════════════════════════════════════════════════════════════════════╝

using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelEngine.Models;
using Microsoft.Web.WebView2.Core;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DeviceInterface.Reports;

/// <summary>
/// Renders self-contained HTML to PDF.  WebView2-primary with silent
/// wkhtmltopdf fallback — the caller never knows which path ran.
/// </summary>
public static class VccsPdfRenderer
{
    // ── Public API ─────────────────────────────────────────────────────────

    /// <summary>
    /// Generates the v23 VCCS RFID report PDF for <paramref name="record"/> into
    /// <paramref name="outputDir"/>.  Fire-and-forget safe: never throws — all
    /// failures are caught and logged via Debug output.
    /// Output: &lt;timestamp&gt;_vccs_rfid_&lt;sessionId&gt;.pdf
    /// </summary>
    public static async Task GenerateReportAsync(
        VerificationRecord record,
        string             outputDir,
        string             sessionId,
        CancellationToken  ct = default)
    {
        if (string.IsNullOrWhiteSpace(outputDir)) return;
        if (!VccsHtmlReportGenerator.HasCorrelatedFilesystemHtml(record))
        {
            Debug.WriteLine(
                "[VCCS-PDF] Capture failure: PDF not generated because no correlated " +
                "DMST filesystem HTML report was attached to this scan.");
            return;
        }
        try
        {
            Directory.CreateDirectory(outputDir);
            string ts      = VccsHtmlReportGenerator.GetOutputTimestamp(record);
            string suffix  = string.IsNullOrWhiteSpace(sessionId) ? "" : $"_{sessionId}";
            string pdfPath = Path.Combine(outputDir, $"{ts}_vccs_rfid{suffix}.pdf");

            string html = VccsHtmlReportGenerator.Generate(record);
            await RenderAsync(html, pdfPath, ct).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[VCCS-PDF] GenerateReportAsync failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    /// <summary>
    /// Renders <paramref name="html"/> to <paramref name="pdfPath"/>.
    /// Tries WebView2 first; on any failure silently falls back to wkhtmltopdf.
    /// Throws only when BOTH paths fail.
    /// </summary>
    public static async Task RenderAsync(string html, string pdfPath, CancellationToken ct = default)
    {
        // The HTML is written to a temp file so both engines load it via file://
        // (avoids WebView2 NavigateToString's 2 MB limit and wkhtmltopdf stdin quirks).
        string tmpHtml = Path.Combine(Path.GetTempPath(),
            $"vccs_report_{Guid.NewGuid():N}.html");
        await File.WriteAllTextAsync(tmpHtml, html, ct).ConfigureAwait(false);

        try
        {
            if (IsWebView2Available())
            {
                try
                {
                    await Task.Run(() => RenderWithWebView2(tmpHtml, pdfPath), ct)
                              .ConfigureAwait(false);
                    AddPageNumbers(pdfPath);
                    Debug.WriteLine($"[VCCS-PDF] Rendered via WebView2: {pdfPath}");
                    return;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(
                        $"[VCCS-PDF] WebView2 render failed ({ex.GetType().Name}: {ex.Message}) — falling back to wkhtmltopdf");
                }
            }
            else
            {
                Debug.WriteLine("[VCCS-PDF] WebView2 runtime not detected — using wkhtmltopdf");
            }

            await RenderWithWkhtmltopdfAsync(tmpHtml, pdfPath, ct).ConfigureAwait(false);
            AddPageNumbers(pdfPath);
            Debug.WriteLine($"[VCCS-PDF] Rendered via wkhtmltopdf: {pdfPath}");
        }
        finally
        {
            try { File.Delete(tmpHtml); } catch { /* temp cleanup is best-effort */ }
        }
    }

    private static void AddPageNumbers(string pdfPath)
    {
        using PdfDocument document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
        int pageCount = document.PageCount;
        XFont font = new("Arial", 7);
        XBrush brush = XBrushes.Gray;

        for (int index = 0; index < pageCount; index++)
        {
            PdfPage page = document.Pages[index];
            using XGraphics graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
            string label = $"Page {index + 1} of {pageCount}";
            graphics.DrawString(
                label,
                font,
                brush,
                new XRect(0, page.Height.Point - 22, page.Width.Point - 29, 12),
                XStringFormats.BottomRight);
        }

        document.Save(pdfPath);
    }

    // ── WebView2 path ──────────────────────────────────────────────────────

    private static bool IsWebView2Available()
    {
        try
        {
            string? v = CoreWebView2Environment.GetAvailableBrowserVersionString();
            return !string.IsNullOrEmpty(v);
        }
        catch
        {
            // Loader missing or no runtime installed — silent.
            return false;
        }
    }

    /// <summary>
    /// Runs a headless WebView2 print on a dedicated STA thread with a hidden
    /// host window and a manual message pump (no WPF/WinForms dependency —
    /// DeviceInterface targets plain net8.0).
    /// </summary>
    private static void RenderWithWebView2(string htmlPath, string pdfPath)
    {
        Exception? failure = null;

        var thread = new Thread(() =>
        {
            IntPtr hwnd = IntPtr.Zero;
            CoreWebView2Controller? controller = null;
            try
            {
                hwnd = CreateWindowExW(0, "STATIC", "VccsPdfHost", 0,
                    0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                if (hwnd == IntPtr.Zero)
                    throw new InvalidOperationException("CreateWindowExW failed");

                string userDataFolder = Path.Combine(Path.GetTempPath(), "vtccp-webview2");

                var envTask = CoreWebView2Environment.CreateAsync(null, userDataFolder);
                PumpUntil(envTask, TimeSpan.FromSeconds(30));
                var env = envTask.Result;

                var ctrlTask = env.CreateCoreWebView2ControllerAsync(hwnd);
                PumpUntil(ctrlTask, TimeSpan.FromSeconds(30));
                controller = ctrlTask.Result;
                var webView = controller.CoreWebView2;

                bool navDone = false;
                webView.NavigationCompleted += (_, e) =>
                {
                    if (!e.IsSuccess)
                        failure ??= new InvalidOperationException(
                            $"WebView2 navigation failed: {e.WebErrorStatus}");
                    navDone = true;
                };
                webView.Navigate(new Uri(htmlPath).AbsoluteUri);
                PumpWhile(() => !navDone, TimeSpan.FromSeconds(30));
                if (failure is not null) throw failure;

                // Letter page, zero margins — the v23 .page div is exactly
                // 8.5×11 in with its own internal padding; print CSS handles
                // backgrounds and hides preview chrome.
                var settings = env.CreatePrintSettings();
                settings.MarginTop    = 0;
                settings.MarginBottom = 0;
                settings.MarginLeft   = 0;
                settings.MarginRight  = 0;
                settings.PageWidth    = 8.5;
                settings.PageHeight   = 11.0;
                settings.ShouldPrintBackgrounds     = true;
                settings.ShouldPrintHeaderAndFooter = false;

                var printTask = webView.PrintToPdfAsync(pdfPath, settings);
                PumpUntil(printTask, TimeSpan.FromSeconds(60));
                if (!printTask.Result)
                    throw new InvalidOperationException("PrintToPdfAsync returned false");
            }
            catch (Exception ex)
            {
                failure = ex is AggregateException { InnerException: { } inner } ? inner : ex;
            }
            finally
            {
                try { controller?.Close(); } catch { }
                if (hwnd != IntPtr.Zero) DestroyWindow(hwnd);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        if (!thread.Join(TimeSpan.FromSeconds(120)))
            throw new TimeoutException("WebView2 render thread did not finish within 120 s");
        if (failure is not null) throw failure;
    }

    private static void PumpUntil(Task task, TimeSpan timeout) =>
        PumpWhile(() => !task.IsCompleted, timeout);

    /// <summary>
    /// Minimal Win32 message pump — dispatches pending messages while the
    /// condition holds.  WebView2 async completions are delivered via posted
    /// messages, so the pump is what drives them on this thread.
    /// </summary>
    private static void PumpWhile(Func<bool> condition, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (condition())
        {
            if (DateTime.UtcNow > deadline)
                throw new TimeoutException("WebView2 operation timed out");
            while (PeekMessageW(out MSG msg, IntPtr.Zero, 0, 0, PM_REMOVE))
            {
                TranslateMessage(ref msg);
                DispatchMessageW(ref msg);
            }
            Thread.Sleep(10);
        }
    }

    // ── wkhtmltopdf fallback ───────────────────────────────────────────────

    private static async Task RenderWithWkhtmltopdfAsync(
        string htmlPath, string pdfPath, CancellationToken ct)
    {
        string? exeDir = Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);
        string exePath = exeDir is not null
            ? Path.Combine(exeDir, "resources", "wkhtmltopdf.exe")
            : "wkhtmltopdf.exe";
        if (!File.Exists(exePath))
            throw new FileNotFoundException(
                "wkhtmltopdf.exe not found (WebView2 also unavailable)", exePath);

        // --print-media-type applies the same @media print rules WebView2 uses;
        // zero margins + Letter match the WebView2 print settings exactly.
        var psi = new ProcessStartInfo
        {
            FileName  = exePath,
            Arguments =
                "-q --print-media-type --page-size Letter " +
                "-T 0 -B 0 -L 0 -R 0 --disable-smart-shrinking " +
                "--enable-local-file-access " +
                $"\"{htmlPath}\" \"{pdfPath}\"",
            UseShellExecute        = false,
            CreateNoWindow         = true,
            WindowStyle            = ProcessWindowStyle.Hidden,
            RedirectStandardError  = true,
            RedirectStandardOutput = true,
        };

        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start wkhtmltopdf");

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(90));
        try
        {
            await proc.WaitForExitAsync(timeoutCts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            try { proc.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("wkhtmltopdf did not finish within 90 s");
        }

        if (proc.ExitCode != 0)
        {
            string err = await proc.StandardError.ReadToEndAsync(ct).ConfigureAwait(false);
            throw new InvalidOperationException(
                $"wkhtmltopdf exit code {proc.ExitCode}: {err.Trim()}");
        }
    }

    // ── Win32 interop ──────────────────────────────────────────────────────

    private const uint PM_REMOVE = 0x0001;

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint   message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint   time;
        public int    pt_x;
        public int    pt_y;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateWindowExW(
        uint dwExStyle, string lpClassName, string lpWindowName, uint dwStyle,
        int x, int y, int nWidth, int nHeight,
        IntPtr hWndParent, IntPtr hMenu, IntPtr hInstance, IntPtr lpParam);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool PeekMessageW(out MSG lpMsg, IntPtr hWnd,
        uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref MSG lpMsg);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessageW(ref MSG lpMsg);
}
