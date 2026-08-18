// Copyright © 2026 VCCS. All rights reserved.

using AsReaderP3xU;
using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// <see cref="IEpcReader"/> implementation for the AsReader ASR-P35U UHF RFID Reader.
///
/// Uses the manufacturer's C# SDK (AsReaderP3xU.dll, v1.3.0) directly — no raw serial
/// framing required.  The SDK manages the VCP internally.
///
/// Prerequisites:
///   1. AsReaderP3xU.dll placed in vtccp\lib\asreader-p3xu-sdk-1.3.0\
///      (not committed — obtain from the AsReader SDK zip; see PLACE-DLL-HERE.md).
///   2. USB cable connected; COM port assigned by Windows Device Manager.
///      VID=0x339C / PID=0x271B (enumerates as a standard VCP — no FTDI driver required).
///   3. Pass the assigned port name to <see cref="ConnectAsync"/> (e.g. "COM4").
///
/// Firmware notes (unit KE00048, FW 1.8.0 / SDK 1.3.0):
///   - <b>CallBackCommandData never fires for ReadMemory</b> (confirmed DLL defect,
///     vendor notified 2026-08-08).  ReadMemory results arrive via
///     CallBackReadTagData instead (tagdata.data or tagdata.tid).
///   - TID read sequence: inventory → cbTag (EPC) → cbComplete → ReadMemory →
///     next cbTag carries TID.  See <see cref="ReadTidAsync"/>.
///   - RSSI: SDK delivers as float; values 128–255 = negative dBm via two's complement.
///
/// Thread safety: all public methods are serialised through <see cref="_lock"/>.
/// SDK callbacks fire on DLL threads; state shared with callbacks is guarded by
/// <see cref="_stateLock"/>.
/// </summary>
public sealed class AsReaderP35UEpcReader : IEpcReader
{
    // ── SDK object ────────────────────────────────────────────────────────────

    private AsReader? _device;

    // ── Connection state ──────────────────────────────────────────────────────

    private volatile bool _connected;
    private bool _disposed;

    /// <inheritdoc />
    public bool IsConnected => _connected && _device is not null;

    // ── Async serialisation ───────────────────────────────────────────────────

    private readonly SemaphoreSlim _lock = new(1, 1);

    // ── Active inventory state (guarded by _stateLock) ────────────────────────

    private readonly object _stateLock = new();
    private List<EpcReadResult>? _pendingResults;
    private TaskCompletionSource<bool>? _inventoryTcs;
    /// <summary>
    /// Set to true when StartInventory is called with maxTags > 0
    /// so <see cref="OnComplete"/> can distinguish an SDK-managed auto-stop
    /// from a normal continuous-mode round completion.
    /// </summary>
    private bool _hwStopExpected;

    // ── One-shot TID hook ─────────────────────────────────────────────────────

    /// <summary>
    /// Set immediately before <c>ReadMemory(MEM_TID)</c> is called.
    /// Consumed by the very next <see cref="OnTagRead"/> invocation (FW 1.8.0
    /// delivers ReadMemory results via cbTag, not cbCommand).
    /// Cleared and invoked atomically; never left set after use.
    /// </summary>
    private volatile Action<AsReader.InventoryResult>? _pendingTidCb;

    // ── Settings ──────────────────────────────────────────────────────────────

    private readonly int _txPowerDbm;
    private const int DefaultTxPowerDbm = 20;  // 20 dBm: safe short-range working default
    private const int MinTxPower        = 13;
    private const int MaxTxPower        = 27;

    // ── Constructor ───────────────────────────────────────────────────────────

    /// <param name="txPowerDbm">TX power in dBm, clamped to 13–27 (US REGION_US range).</param>
    public AsReaderP35UEpcReader(int txPowerDbm = DefaultTxPowerDbm)
    {
        _txPowerDbm = Math.Clamp(txPowerDbm, MinTxPower, MaxTxPower);
    }

    // ── IEpcReader — Connect / Disconnect ─────────────────────────────────────

    /// <inheritdoc />
    /// <remarks>
    /// Loads <c>AsReaderP3xU.dll</c>, creates an <see cref="AsReader"/> instance,
    /// registers all six SDK delegates, calls <c>ConnectWithVCP</c>, then sets
    /// <c>REGION_US</c> and TX power.  All six delegates must be registered in a
    /// single <c>SetDelegate</c> call before connecting — SDK requirement.
    /// </remarks>
    public async Task ConnectAsync(string portName, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (IsConnected)
                throw new InvalidOperationException(
                    "Reader already connected. Call DisconnectAsync first.");

            var dev = new AsReader();

            // All six delegates must be registered in a single SetDelegate call,
            // and the call must happen BEFORE ConnectWithVCP — SDK contract.
            dev.SetDelegate(
                new AsReader.CallBackReadTagData(OnTagRead),
                new AsReader.CallBackErrorCode(OnError),
                new AsReader.CallBackSuccessCode(OnSuccess),
                new AsReader.CallBackCommandData(OnCommandData),
                new AsReader.CallBackReadComplete(OnComplete),
                new AsReader.CallBackTriggerHandler(OnTrigger)
            );

            uint ret = dev.ConnectWithVCP(portName);
            if (ret != 0)
                throw new IOException(
                    $"AsReader ConnectWithVCP(\"{portName}\") failed — error code {ret}. " +
                    "Verify the COM port is correct and the reader is connected.");

            // Region MUST be set before StartInventory.
            dev.SetRegion(Types.RegionType.REGION_US);
            dev.SetTxPower((uint)_txPowerDbm);

            _device    = dev;
            _connected = true;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <inheritdoc />
    public async Task DisconnectAsync(CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (_device is { } dev)
            {
                try { dev.StopInventory(); }  catch { /* best effort */ }
                try { dev.DisConnect(); }     catch { /* best effort */ }
            }
            _device    = null;
            _connected = false;
            AbortActiveInventory();
        }
        finally
        {
            _lock.Release();
        }
    }

    // ── IEpcReader — Inventory ────────────────────────────────────────────────

    /// <inheritdoc />
    /// <remarks>
    /// Uses single-tag mode (<c>maxTags=1</c>): the SDK auto-stops after reading
    /// the first tag, which eliminates the race between a Python-style stop timer
    /// and the StopInventory call.  For multi-tag sweeps, call this method once
    /// per tag or extend with a continuous-mode variant.
    ///
    /// Returns immediately if no tag is detected before <paramref name="timeout"/>
    /// (<c>cbComplete(false)</c> fires, or the timeout elapses).
    ///
    /// Includes TID if the hardware delivers it automatically; explicit TID
    /// reading via <see cref="ReadTidAsync"/> requires a separate call after
    /// inventory (see FW 1.8.0 firmware note in class summary).
    /// </remarks>
    public async Task<IReadOnlyList<EpcReadResult>> TriggerInventoryAsync(
        TimeSpan      timeout,
        CancellationToken ct = default)
    {
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync first.");

        TaskCompletionSource<bool> tcs;
        lock (_stateLock)
        {
            _pendingResults = new List<EpcReadResult>();
            _hwStopExpected = true;   // maxTags=1 → SDK will auto-stop
            _pendingTidCb   = null;
            tcs             = new TaskCompletionSource<bool>(
                                  TaskCreationOptions.RunContinuationsAsynchronously);
            _inventoryTcs   = tcs;
        }

        // maxTags=1: stop after the first tag — hardware-managed, no timer race.
        // SDK signature: StartInventory(bool rssiEnabled, byte maxTags, byte maxSecs, ushort maxCycles, bool an1)
        _device!.StartInventory(true, 1, 0, 0, true);

        // Wait for: hardware auto-stop (cbComplete), timeout, or cancellation.
        using var linkedCts =
            CancellationTokenSource.CreateLinkedTokenSource(ct);
        linkedCts.CancelAfter(timeout);
        try
        {
            await tcs.Task.WaitAsync(linkedCts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // Timeout or external cancellation — stop and return whatever we have.
            try { _device?.StopInventory(); } catch { /* best effort */ }
        }

        List<EpcReadResult> results;
        lock (_stateLock)
        {
            results         = _pendingResults ?? new List<EpcReadResult>();
            _pendingResults = null;
            _inventoryTcs   = null;
            _hwStopExpected = false;
        }
        return results.AsReadOnly();
    }

    /// <inheritdoc />
    public Task CancelAsync(CancellationToken ct = default)
    {
        try { _device?.StopInventory(); } catch { /* best effort */ }
        AbortActiveInventory();
        return Task.CompletedTask;
    }

    // ── TID reading ───────────────────────────────────────────────────────────

    /// <summary>
    /// Read the TID memory bank of a specific tag after inventory.
    ///
    /// Call this after <see cref="TriggerInventoryAsync"/> returns an EPC, before
    /// starting the next inventory.
    ///
    /// FW 1.8.0 defect workaround: the result does NOT arrive via
    /// <c>CallBackCommandData</c> (confirmed DLL bug).  This method registers a
    /// one-shot hook on <see cref="OnTagRead"/> (<see cref="_pendingTidCb"/>) and
    /// waits for the next cbTag call to deliver the TID in
    /// <c>tagdata.data</c> or <c>tagdata.tid</c>.
    /// </summary>
    /// <param name="epcBytes">EPC bytes identifying the target tag.</param>
    /// <param name="timeout">Maximum wait for the TID result callback.</param>
    /// <returns>TID hex string (uppercase), or null on timeout / error.</returns>
    public async Task<string?> ReadTidAsync(
        byte[] epcBytes,
        TimeSpan timeout,
        CancellationToken ct = default)
    {
        if (!IsConnected || _device is null)
            return null;

        var tidTcs = new TaskCompletionSource<string?>(
            TaskCreationOptions.RunContinuationsAsynchronously);

        // Register one-shot hook consumed by the next OnTagRead (FW 1.8.0 path).
        _pendingTidCb = result =>
        {
            try
            {
                var td  = result.tagdata;
                // FW 1.8.0: ReadMemory result comes back in tagdata.data first,
                // tagdata.tid as fallback.
                string? raw = td.data ?? td.tid;
                string? tid = null;
                if (!string.IsNullOrWhiteSpace(raw))
                    tid = raw.Trim().ToUpperInvariant().Replace(" ", "");
                tidTcs.TrySetResult(tid);
            }
            catch (Exception ex)
            {
                Dbg($"ReadTidAsync hook error: {ex.Message}");
                tidTcs.TrySetResult(null);
            }
        };

        // SDK: ReadMemory(MemBankType, startAddr, length, password, epcBytes)
        // 4 words = 8 bytes = 64-bit TID, password=0 (no access protection).
        uint ret = _device.ReadMemory(
            Types.MemBankType.MEM_TID,
            0,
            4,
            0,
            epcBytes
        );

        if (ret != 0)
        {
            _pendingTidCb = null;   // clear the hook — ReadMemory was rejected
            Dbg($"ReadMemory returned {ret} — TID read rejected");
            return null;
        }

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        linkedCts.CancelAfter(timeout);
        try
        {
            return await tidTcs.Task.WaitAsync(linkedCts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            Interlocked.Exchange(ref _pendingTidCb, null);  // disarm stale hook
            Dbg("ReadTidAsync timed out");
            return null;
        }
    }

    // ── SDK callbacks (fire on DLL threads) ───────────────────────────────────

    private void OnTagRead(AsReader.InventoryResult result)
    {
        try
        {
            // Check for one-shot TID hook first (FW 1.8.0: ReadMemory delivers
            // result via cbTag, not cbCommand).
            var tidCb = Interlocked.Exchange(ref _pendingTidCb, null);
            if (tidCb is not null)
            {
                tidCb(result);
                return;
            }

            var td = result.tagdata;
            if (td.epc is not string epcHex || string.IsNullOrWhiteSpace(epcHex))
                return;

            epcHex = epcHex.Trim().ToUpperInvariant().Replace(" ", "");

            // PC word: SDK delivers as hex string e.g. "3000"
            ushort pcWord = 0;
            if (td.pc is string pcStr)
                ushort.TryParse(pcStr,
                    System.Globalization.NumberStyles.HexNumber,
                    null, out pcWord);

            // RSSI: SDK delivers as float; two's complement for values 128–255.
            int? rssi = null;
            try
            {
                double raw = Convert.ToDouble(result.rssi);
                rssi = raw is >= 128 and <= 255 ? (int)(raw - 256) : (int)raw;
            }
            catch { /* RSSI not critical */ }

            // TID may arrive directly in continuous PC_EPC_TID inventory mode
            // (if SetHIDInventoryMode were callable — see ASREADER_TID_DEFECT.md).
            // In standard mode it will be null here; ReadTidAsync handles it separately.
            string? tid = null;
            if (td.tid is string tidRaw && !string.IsNullOrWhiteSpace(tidRaw))
                tid = tidRaw.Trim().ToUpperInvariant().Replace(" ", "");

            byte[] epcBytes;
            try { epcBytes = Convert.FromHexString(epcHex); }
            catch { return; }   // malformed EPC hex — discard

            var readResult = new EpcReadResult
            {
                EpcBytes = epcBytes,
                PcWord   = pcWord,
                Rssi     = rssi,
                ReadTime = DateTimeOffset.UtcNow,
                Tid      = tid,
            };

            Dbg($"TAG epc={epcHex} tid={tid ?? "(none)"} rssi={rssi} pc=0x{pcWord:X4}");

            lock (_stateLock)
                _pendingResults?.Add(readResult);
        }
        catch (Exception ex)
        {
            // Never propagate exceptions from a DLL callback thread.
            Dbg($"OnTagRead error: {ex.Message}");
        }
    }

    private void OnError(uint errorCode)
    {
        Dbg($"DLL error callback: {errorCode}");
        // Only treat as disconnect when inventory is actually running.
        // Errors while idle are typically spurious QC-command responses.
        lock (_stateLock)
        {
            if (_inventoryTcs is not null)
            {
                _connected = false;
                AbortActiveInventory();
            }
        }
    }

    private void OnSuccess(uint successCode)
    {
        // successCode: 40=PermaLock, 41=Lock, 42=Unlock (from CheckTagStatus)
        Dbg($"DLL success callback: {successCode}");
    }

    private void OnCommandData(byte[]? data)
    {
        // NOTE: This callback never fires for ReadMemory on FW 1.8.0 (confirmed
        // DLL defect — see AsReader TID defect report).  ReadMemory results
        // arrive via OnTagRead instead.  Kept wired for future SDK corrections
        // and to satisfy the mandatory six-delegate SetDelegate requirement.
        Dbg($"DLL command callback: {(data?.Length ?? 0)} bytes " +
            "(expected empty — see ASREADER_TID_DEFECT.md)");
    }

    private void OnComplete(bool completeStatus)
    {
        Dbg($"DLL read complete: status={completeStatus} hwStopExpected={_hwStopExpected}");

        lock (_stateLock)
        {
            if (_hwStopExpected && completeStatus)
            {
                // Hardware auto-stopped after reaching the maxTags limit (clean path).
                _hwStopExpected = false;
                _inventoryTcs?.TrySetResult(true);
            }
            else if (!completeStatus)
            {
                // SDK convention: false = unexpected stop / hardware disconnect.
                _connected      = false;
                _hwStopExpected = false;
                _inventoryTcs?.TrySetResult(false);
            }
            // completeStatus=true without _hwStopExpected = normal continuous-mode
            // round completion; inventory still running — no action.
        }
    }

    private void OnTrigger(int triggerState)
    {
        // triggerState: 1 = hardware SCAN button pressed, 0 = released.
        // Can be used to drive trigger-mode inventory in the UI layer.
        Dbg($"DLL trigger: {triggerState}");
    }

    // ── Port discovery ─────────────────────────────────────────────────────────

    /// <summary>
    /// Returns all COM ports available on this machine.
    /// The ASR-P35U enumerates as VID=0x339C / PID=0x271B.
    /// Full VID/PID discrimination requires WMI; this returns all ports as a
    /// safe fallback for a COM port picker UI.
    /// </summary>
    public static IReadOnlyList<string> GetAvailablePorts() =>
        System.IO.Ports.SerialPort.GetPortNames()
              .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
              .ToList()
              .AsReadOnly();

    // ── Helpers ────────────────────────────────────────────────────────────────

    private void AbortActiveInventory()
    {
        lock (_stateLock)
        {
            _hwStopExpected = false;
            _pendingTidCb   = null;
            _inventoryTcs?.TrySetCanceled();
            _inventoryTcs   = null;
        }
    }

    [System.Diagnostics.Conditional("DEBUG")]
    private static void Dbg(string msg) =>
        System.Diagnostics.Debug.WriteLine($"[AsReaderP35U] {msg}");

    // ── IAsyncDisposable ───────────────────────────────────────────────────────

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        await DisconnectAsync().ConfigureAwait(false);
        _lock.Dispose();
    }
}
