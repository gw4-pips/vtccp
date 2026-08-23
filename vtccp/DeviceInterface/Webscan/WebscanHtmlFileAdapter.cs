namespace DeviceInterface.Webscan;

using ExcelEngine.Models;

/// <summary>
/// File-export adapter for Webscan TruCheck.
///
/// Webscan results arrive as local HTML files because the verifier is USB
/// connected. This adapter never opens a TCP/HTTP connection and never deletes
/// or rewrites the source HTML or its sibling image.
/// </summary>
public sealed class WebscanHtmlFileAdapter : IDisposable
{
    public const string ConfiguredReportDirectory =
        @"C:\dev\vtccp\TC-829 VeriWedge Dev Reports";

    private readonly string _watchDirectory;
    private readonly object _lock = new();
    private readonly HashSet<string> _processing =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, FileFingerprint> _imported =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<Task> _inFlight = [];
    private FileSystemWatcher? _watcher;
    private CancellationTokenSource? _lifetimeCts;
    private bool _accepting;
    private bool _disposed;

    public WebscanHtmlFileAdapter(string? watchDirectory = null)
    {
        _watchDirectory = string.IsNullOrWhiteSpace(watchDirectory)
            ? ConfiguredReportDirectory
            : Path.GetFullPath(watchDirectory);
    }

    public string WatchDirectory => _watchDirectory;
    public event EventHandler<VerificationRecord>? RecordParsed;
    public event EventHandler<string>? ParseFailed;

    public void Start()
    {
        ThrowIfDisposed();
        Directory.CreateDirectory(_watchDirectory);
        lock (_lock)
        {
            if (_watcher is not null) return;
            _lifetimeCts = new CancellationTokenSource();
            _accepting = true;
        }

        _watcher = new FileSystemWatcher(_watchDirectory, "*.html")
        {
            NotifyFilter = NotifyFilters.FileName |
                           NotifyFilters.LastWrite |
                           NotifyFilters.Size,
            IncludeSubdirectories = false,
            EnableRaisingEvents = true,
        };
        _watcher.Created += OnFileChanged;
        _watcher.Changed += OnFileChanged;
    }

    public async Task StopAsync()
    {
        FileSystemWatcher? watcher;
        CancellationTokenSource? lifetimeCts;
        Task[] inFlight;
        lock (_lock)
        {
            _accepting = false;
            watcher = _watcher;
            _watcher = null;
            lifetimeCts = _lifetimeCts;
            _lifetimeCts = null;
            inFlight = _inFlight.ToArray();
        }

        if (watcher is not null)
        {
            watcher.EnableRaisingEvents = false;
            watcher.Created -= OnFileChanged;
            watcher.Changed -= OnFileChanged;
            watcher.Dispose();
        }
        if (lifetimeCts is not null)
            await lifetimeCts.CancelAsync();

        try
        {
            await Task.WhenAll(inFlight);
        }
        catch (OperationCanceledException)
        {
            // A report still being written at session stop is intentionally
            // abandoned. The original source file remains untouched.
        }
        finally
        {
            lifetimeCts?.Dispose();
        }
    }

    /// <summary>
    /// Imports one report without changing it on disk. This is the deterministic
    /// entry point used by the controlled-scan test and by the watcher callback.
    /// </summary>
    public async Task<VerificationRecord> ImportFileAsync(
        string sourcePath,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        if (!sourcePath.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Webscan adapter accepts .html exports only.", nameof(sourcePath));

        string fullPath = Path.GetFullPath(sourcePath);
        string rawHtml = await ReadStableTextAsync(fullPath, cancellationToken);
        VerificationRecord record;
        if (rawHtml.Contains("Symbol 2 Verification Report",
                StringComparison.OrdinalIgnoreCase))
        {
            WebscanHtmlCompositeReport composite =
                WebscanHtmlParser.ParseComposite(rawHtml, fullPath);
            if (!composite.ParseSucceeded)
                throw new InvalidDataException(
                    composite.ParseError ?? "Webscan composite HTML parse failed.");
            record = composite.ToVerificationRecord();
        }
        else
        {
            WebscanHtmlReport report = WebscanHtmlParser.Parse(rawHtml, fullPath);
            if (!report.ParseSucceeded)
                throw new InvalidDataException(
                    report.ParseError ?? "Webscan HTML parse failed.");
            record = report.ToVerificationRecord();
        }
        cancellationToken.ThrowIfCancellationRequested();

        // The source HTML and image remain untouched. Keep the raw input inside
        // the report during parsing so callers can archive it if required.
        cancellationToken.ThrowIfCancellationRequested();
        RaiseRecordParsed(record);
        return record;
    }

    private void OnFileChanged(object sender, FileSystemEventArgs args)
    {
        string fullPath = Path.GetFullPath(args.FullPath);
        FileFingerprint? fingerprint = TryGetFingerprint(fullPath);
        Task? task;
        lock (_lock)
        {
            if (!_accepting || _lifetimeCts is null)
                return;
            if (fingerprint is { } current &&
                _imported.TryGetValue(fullPath, out FileFingerprint previous) &&
                previous == current)
                return;
            if (!_processing.Add(fullPath))
                return;
            task = ProcessFileAsync(fullPath, _lifetimeCts.Token);
            _inFlight.Add(task);
            _ = task.ContinueWith(completed =>
            {
                lock (_lock)
                {
                    _processing.Remove(fullPath);
                    _inFlight.Remove(completed);
                }
            }, TaskScheduler.Default);
        }
    }

    private async Task ProcessFileAsync(string sourcePath, CancellationToken cancellationToken)
    {
        try
        {
            await ImportFileAsync(sourcePath, cancellationToken);
            FileFingerprint? fingerprint = TryGetFingerprint(sourcePath);
            if (fingerprint is { } current)
            {
                lock (_lock) _imported[sourcePath] = current;
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // Session stop cancelled this import before it could emit a record.
        }
        catch (Exception ex)
        {
            ParseFailed?.Invoke(this, $"{Path.GetFileName(sourcePath)}: {ex.Message}");
        }
    }

    private static async Task<string> ReadStableTextAsync(
        string sourcePath,
        CancellationToken cancellationToken)
    {
        const int attempts = 20;
        for (int attempt = 0; attempt < attempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                await using var stream = new FileStream(
                    sourcePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
                using var reader = new StreamReader(stream);
                string text = await reader.ReadToEndAsync(cancellationToken);
                if (text.Contains("</html>", StringComparison.OrdinalIgnoreCase))
                    return text;
            }
            catch (IOException) when (attempt < attempts - 1)
            {
                // Webscan may still be finishing the export.
            }
            await Task.Delay(100, cancellationToken);
        }

        throw new IOException($"Webscan report did not become readable: {sourcePath}");
    }

    private static FileFingerprint? TryGetFingerprint(string sourcePath)
    {
        try
        {
            var info = new FileInfo(sourcePath);
            return info.Exists
                ? new FileFingerprint(info.Length, info.LastWriteTimeUtc.Ticks)
                : null;
        }
        catch (IOException)
        {
            return null;
        }
    }

    private void RaiseRecordParsed(VerificationRecord record)
        => RecordParsed?.Invoke(this, record);

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WebscanHtmlFileAdapter));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        StopAsync().GetAwaiter().GetResult();
    }

    private readonly record struct FileFingerprint(long Length, long LastWriteUtcTicks);
}