namespace DeviceInterface.FileAdapters;

using ExcelEngine.Models;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// Imports CSV files written by the VTCCP Axicon automatic-export templates.
/// Source exports are never deleted or modified.
/// </summary>
public sealed class AxiconFileAdapter : IDisposable
{
    public const string Brand = "AXICON";
    private readonly string _watchFolder;
    private readonly string? _artifactCopyDirectory;
    private readonly object _lock = new();
    private readonly HashSet<string> _processing = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _importedHashes = new(StringComparer.Ordinal);
    private readonly HashSet<string> _acceptingHashes = new(StringComparer.Ordinal);
    private readonly HashSet<Task> _inFlight = [];
    private FileSystemWatcher? _watcher;
    private CancellationTokenSource? _lifetimeCts;
    private bool _accepting;
    private bool _disposed;

    public AxiconFileAdapter(string watchFolder, string? artifactCopyDirectory = null)
    {
        if (string.IsNullOrWhiteSpace(watchFolder))
            throw new ArgumentException("Watch folder path must not be empty.", nameof(watchFolder));
        _watchFolder = Path.GetFullPath(watchFolder);
        _artifactCopyDirectory = string.IsNullOrWhiteSpace(artifactCopyDirectory)
            ? null : Path.GetFullPath(artifactCopyDirectory);
    }

    public event EventHandler<VerificationRecord>? RecordParsed;
    /// <summary>Raised for watcher imports that cannot be read or parsed.</summary>
    public event EventHandler<string>? ParseFailed;
    public string WatchFolder => _watchFolder;

    public void Start()
    {
        ThrowIfDisposed();
        Directory.CreateDirectory(_watchFolder);
        lock (_lock)
        {
            if (_watcher is not null) return;
            _lifetimeCts = new CancellationTokenSource();
            _accepting = true;
        }
        _watcher = new FileSystemWatcher(_watchFolder, "*.csv")
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size,
            EnableRaisingEvents = true,
        };
        _watcher.Created += OnFileChanged;
        _watcher.Changed += OnFileChanged;
    }

    /// <summary>Stops new imports and drains/cancels imports already in progress.</summary>
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
        if (lifetimeCts is not null) await lifetimeCts.CancelAsync().ConfigureAwait(false);
        try { await Task.WhenAll(inFlight).ConfigureAwait(false); }
        catch (OperationCanceledException) { }
        finally { lifetimeCts?.Dispose(); }
    }

    /// <summary>
    /// Deterministically imports one completed Axicon CSV. The source file is
    /// read only; if an archive directory was configured, a byte-for-byte copy
    /// is made after parsing succeeds.
    /// </summary>
    public async Task<VerificationRecord> ImportFileAsync(
        string sourcePath, CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        if (!sourcePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Axicon adapter accepts .csv exports only.", nameof(sourcePath));

        string fullPath = Path.GetFullPath(sourcePath);
        ImportedFile imported = await PrepareImportAsync(fullPath, cancellationToken).ConfigureAwait(false);
        VerificationRecord record = imported.Record;
        if (_artifactCopyDirectory is not null)
        {
            string archivedPath = await CopyArtifactAsync(
                imported.Bytes, fullPath, _artifactCopyDirectory, imported.Hash,
                imported.LastWriteTime, cancellationToken).ConfigureAwait(false);
            record = record with { SourceArtifactPath = archivedPath };
        }
        cancellationToken.ThrowIfCancellationRequested();
        RecordParsed?.Invoke(this, record);
        return record;
    }

    internal static async Task<VerificationRecord> ParseFileAsync(string filePath)
        => await new AxiconFileAdapter(Path.GetDirectoryName(Path.GetFullPath(filePath)) ?? ".")
            .ImportFileAsync(filePath).ConfigureAwait(false);

    /// <summary>
    /// Compatibility helper retained for callers that only require the brand
    /// marker. File ingestion always uses <see cref="ImportFileAsync"/>, which
    /// reports malformed content explicitly rather than returning this fallback.
    /// </summary>
    public static VerificationRecord BuildRecord(string? rawContent = null, DateTime? timestamp = null)
    {
        AxiconAutomaticExportReport report = AxiconAutomaticExportParser.Parse(
            rawContent ?? string.Empty, string.Empty, timestamp ?? DateTime.Now);
        return report.ParseSucceeded
            ? report.ToVerificationRecord()
            : new VerificationRecord
            {
                VerificationDateTime = timestamp ?? DateTime.Now,
                Symbology = "Unknown",
                VerifierBrand = Brand,
            };
    }

    private void OnFileChanged(object sender, FileSystemEventArgs args)
    {
        string path = Path.GetFullPath(args.FullPath);
        Task? task;
        lock (_lock)
        {
            if (!_accepting || _lifetimeCts is null || !_processing.Add(path))
                return;
            task = ProcessFileAsync(path, _lifetimeCts.Token);
            _inFlight.Add(task);
            _ = task.ContinueWith(done =>
            {
                lock (_lock) { _processing.Remove(path); _inFlight.Remove(done); }
            }, TaskScheduler.Default);
        }
    }

    private async Task ProcessFileAsync(string path, CancellationToken cancellationToken)
    {
        try
        {
            ImportedFile imported = await PrepareImportAsync(path, cancellationToken).ConfigureAwait(false);
            lock (_lock)
            {
                if (_importedHashes.Contains(imported.Hash) || !_acceptingHashes.Add(imported.Hash))
                    return;
            }
            try
            {
                VerificationRecord record = imported.Record;
                if (_artifactCopyDirectory is not null)
                {
                    string archivedPath = await CopyArtifactAsync(
                        imported.Bytes, path, _artifactCopyDirectory, imported.Hash,
                        imported.LastWriteTime, cancellationToken).ConfigureAwait(false);
                    record = record with { SourceArtifactPath = archivedPath };
                }
                cancellationToken.ThrowIfCancellationRequested();
                RecordParsed?.Invoke(this, record);
                lock (_lock) _importedHashes.Add(imported.Hash);
            }
            finally
            {
                lock (_lock) _acceptingHashes.Remove(imported.Hash);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        catch (Exception ex) { ParseFailed?.Invoke(this, $"{Path.GetFileName(path)}: {ex.Message}"); }
    }

    private static async Task<ImportedFile> PrepareImportAsync(string path, CancellationToken cancellationToken)
    {
        StableFile stable = await ReadStableFileAsync(path, cancellationToken).ConfigureAwait(false);
        string text = Encoding.UTF8.GetString(stable.Bytes);
        AxiconAutomaticExportReport report = AxiconAutomaticExportParser.Parse(text, path, stable.LastWriteTime);
        if (!report.ParseSucceeded)
            throw new InvalidDataException(report.ParseError ?? "Axicon automatic-export parse failed.");
        return new ImportedFile(report.ToVerificationRecord(), stable.Bytes,
            Convert.ToHexString(SHA256.HashData(stable.Bytes)), stable.LastWriteTime);
    }

    private static async Task<StableFile> ReadStableFileAsync(string path, CancellationToken cancellationToken)
    {
        const int attempts = 12;
        for (int attempt = 0; attempt < attempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                FileFingerprint first = GetFingerprint(path);
                await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
                using var memory = new MemoryStream();
                await stream.CopyToAsync(memory, cancellationToken).ConfigureAwait(false);
                byte[] bytes = memory.ToArray();
                if (bytes.Length == 0) continue;
                await Task.Delay(250, cancellationToken).ConfigureAwait(false);
                if (first != GetFingerprint(path)) continue;
                await Task.Delay(250, cancellationToken).ConfigureAwait(false);
                if (first == GetFingerprint(path))
                    return new StableFile(bytes, new DateTime(first.LastWriteUtcTicks, DateTimeKind.Utc).ToLocalTime());
            }
            catch (IOException) when (attempt < attempts - 1) { }
            await Task.Delay(100, cancellationToken).ConfigureAwait(false);
        }
        throw new IOException($"Axicon export did not remain unchanged for 500ms: {path}");
    }

    private static async Task<string> CopyArtifactAsync(
        byte[] bytes,
        string sourcePath,
        string directory,
        string contentHash,
        DateTime sourceLastWriteTime,
        CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(directory);
        string stem = Path.GetFileNameWithoutExtension(sourcePath);
        string extension = Path.GetExtension(sourcePath);
        string stamp = sourceLastWriteTime.ToUniversalTime().ToString("yyyyMMddTHHmmssfffffffZ");
        string target = Path.Combine(directory, $"{stem}.{stamp}.{contentHash[..16]}{extension}");
        if (File.Exists(target))
        {
            if (File.ReadAllBytes(target).AsSpan().SequenceEqual(bytes))
                return target;
            throw new IOException($"Refusing to overwrite distinct Axicon artifact: {target}");
        }
        string temporary = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await File.WriteAllBytesAsync(temporary, bytes, cancellationToken).ConfigureAwait(false);
            try
            {
                File.Move(temporary, target);
            }
            catch (IOException) when (File.Exists(target))
            {
                if (File.ReadAllBytes(target).AsSpan().SequenceEqual(bytes))
                    return target;
                throw new IOException($"Refusing to overwrite distinct Axicon artifact: {target}");
            }
            return target;
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private static FileFingerprint GetFingerprint(string path)
    {
        var info = new FileInfo(path);
        if (!info.Exists) throw new FileNotFoundException("Axicon export was not found.", path);
        return new FileFingerprint(info.Length, info.LastWriteTimeUtc.Ticks);
    }
    private void ThrowIfDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(AxiconFileAdapter)); }
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        lock (_lock) { _accepting = false; _lifetimeCts?.Cancel(); }
        if (_watcher is not null) { _watcher.EnableRaisingEvents = false; _watcher.Created -= OnFileChanged; _watcher.Changed -= OnFileChanged; _watcher.Dispose(); }
        _lifetimeCts?.Dispose();
    }
    private readonly record struct FileFingerprint(long Length, long LastWriteUtcTicks);
    private readonly record struct StableFile(byte[] Bytes, DateTime LastWriteTime);
    private readonly record struct ImportedFile(
        VerificationRecord Record, byte[] Bytes, string Hash, DateTime LastWriteTime);
}