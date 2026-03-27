namespace DeviceInterface.Testing;

using System.Net;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// In-process TCP server that speaks the DMCC wire protocol for offline testing.
///
/// Usage:
///   await using var mock = new MockDmccServer();
///   mock.SetResponse("GET DEVICE.TYPE", "DM260Q");
///   mock.SetResponse("GET SYMBOL.RESULT", MockDmccServer.SampleDm2DXml);
///   int port = mock.Port;
///   // point DmccClient at 127.0.0.1:port
///
/// The mock accepts a single connection, handles commands sequentially,
/// and stops when Dispose/DisposeAsync is called.
/// </summary>
public sealed class MockDmccServer : IAsyncDisposable
{
    private readonly TcpListener _server;
    private readonly CancellationTokenSource _cts = new();
    private readonly Dictionary<string, string> _responses = new(StringComparer.OrdinalIgnoreCase);
    private Task? _serveTask;

    /// <summary>Port the mock is listening on (OS-assigned).</summary>
    public int Port { get; }

    public MockDmccServer()
    {
        _server = new TcpListener(IPAddress.Loopback, 0);
        _server.Start();
        Port = ((IPEndPoint)_server.LocalEndpoint).Port;

        // Default well-known responses.
        _responses["GET DEVICE.TYPE"]      = "DM260Q";
        _responses["GET FIRMWARE.VER"]     = "5.7.4.0015";
        _responses["GET DEVICE.NAME"]      = "VTCCP-Test-Device";
        _responses["GET DEVICE.ID"]        = "DM-TEST-001";
        _responses["GET CALIBRATION.DATE"] = "2025-01-01";
        _responses["SET DMCC.RESULT-FORMAT FULL"] = "";
        _responses["TRIGGER"]              = "";
        _responses["GET SYMBOL.RESULT"]    = SampleDm2DXml;

        _serveTask = Task.Run(() => ServeLoopAsync(_cts.Token));
    }

    // ── Configuration ─────────────────────────────────────────────────────────

    /// <summary>Override the response body for a specific command.</summary>
    public void SetResponse(string command, string responseBody) =>
        _responses[command.Trim()] = responseBody;

    /// <summary>Remove a configured response (unknown commands return status 8 / Busy).</summary>
    public void RemoveResponse(string command) =>
        _responses.Remove(command.Trim());

    // ── Sample XML payloads ────────────────────────────────────────────────────

    /// <summary>Minimal GS1 DataMatrix verification result XML for 2D round-trip testing.</summary>
    public static readonly string SampleDm2DXml = """
        <?xml version="1.0"?>
        <DMCCResponse>
          <ResponseData>
            <DMSymVerResponse>
              <DateTime>2025-06-15T09:30:00.000</DateTime>
              <SymbologyName>GS1 DataMatrix</SymbologyName>
              <DecodedData><![CDATA[<F1>010123456789012817251231101234-LOT-A<F1>2199887766]]></DecodedData>
              <FormalGrade>4.0/16/660/45Q</FormalGrade>
              <OverallGrade>A</OverallGrade>
              <OverallGradeNumeric>4.0</OverallGradeNumeric>
              <ApertureRef>16</ApertureRef>
              <Wavelength>660</Wavelength>
              <Lighting>45Q</Lighting>
              <Standard>ISO 15415:2011</Standard>
              <UECPercent>100</UECPercent>
              <UECGrade>A</UECGrade>
              <SCPercent>84</SCPercent>
              <SCGrade>A</SCGrade>
              <MODGrade>A</MODGrade>
              <ANUPercent>0.2</ANUPercent>
              <ANUGrade>A</ANUGrade>
              <GNUPercent>2.3</GNUPercent>
              <GNUGrade>A</GNUGrade>
              <FPDGrade>A</FPDGrade>
              <DecodeGrade>A</DecodeGrade>
              <AGValue>4.0</AGValue>
              <AGGrade>A</AGGrade>
              <MatrixSize>22x22</MatrixSize>
              <EncodedCharacters>20</EncodedCharacters>
              <TotalCodewords>144</TotalCodewords>
              <DataCodewords>12</DataCodewords>
              <ErrorCorrectionBudget>62</ErrorCorrectionBudget>
              <ErrorsCorrected>0</ErrorsCorrected>
              <ErrorCapacityUsed>0</ErrorCapacityUsed>
              <ErrorCorrectionType>ECC 200</ErrorCorrectionType>
              <NominalXDim>0.010</NominalXDim>
              <PixelsPerModule>4.2</PixelsPerModule>
              <LLSGrade>A</LLSGrade>
              <BLSGrade>A</BLSGrade>
              <LQZGrade>A</LQZGrade>
              <BQZGrade>A</BQZGrade>
              <TQZGrade>A</TQZGrade>
              <RQZGrade>A</RQZGrade>
              <TTRPercent>95.5</TTRPercent>
              <TTRGrade>A</TTRGrade>
              <RTRPercent>94.2</RTRPercent>
              <RTRGrade>A</RTRGrade>
              <TCTGrade>A</TCTGrade>
              <RCTGrade>A</RCTGrade>
            </DMSymVerResponse>
          </ResponseData>
        </DMCCResponse>
        """;

    /// <summary>Minimal UPC-A verification result XML for 1D round-trip testing.</summary>
    public static readonly string SampleUpcAXml = """
        <?xml version="1.0"?>
        <DMCCResponse>
          <ResponseData>
            <DMSymVerResponse>
              <DateTime>2025-06-15T09:31:00.000</DateTime>
              <SymbologyName>UPC-A</SymbologyName>
              <DecodedData><![CDATA[012345678905]]></DecodedData>
              <FormalGrade>4.0/06/660</FormalGrade>
              <OverallGrade>A</OverallGrade>
              <OverallGradeNumeric>4.0</OverallGradeNumeric>
              <ApertureRef>6</ApertureRef>
              <Wavelength>660</Wavelength>
              <Lighting>Diffuse</Lighting>
              <Standard>ISO 15416:2016</Standard>
              <SymbolAnsiGrade>A</SymbolAnsiGrade>
              <AvgEdge>59</AvgEdge>
              <AvgSC>84</AvgSC>
              <AvgMinEC>4.0</AvgMinEC>
              <AvgMOD>4.0</AvgMOD>
              <AvgDefect>4.0</AvgDefect>
              <AvgDEC>4.0</AvgDEC>
              <AvgLQZ>10</AvgLQZ>
              <AvgRQZ>11</AvgRQZ>
              <AvgMinQZ>10</AvgMinQZ>
              <BWGPercent>8</BWGPercent>
              <Magnification>1.0</Magnification>
              <ScanResults>
                <Scan number="1">
                  <Edge>4</Edge><SC>4</SC><MinEC>4.0</MinEC><MOD>4.0</MOD>
                  <Defect>4.0</Defect><DEC>4.0</DEC><LQZ>4</LQZ><RQZ>4</RQZ><HQZ>4</HQZ>
                </Scan>
                <Scan number="2">
                  <Edge>4</Edge><SC>4</SC><MinEC>4.0</MinEC><MOD>4.0</MOD>
                  <Defect>4.0</Defect><DEC>4.0</DEC><LQZ>4</LQZ><RQZ>4</RQZ><HQZ>4</HQZ>
                </Scan>
              </ScanResults>
            </DMSymVerResponse>
          </ResponseData>
        </DMCCResponse>
        """;

    // ── Server loop ───────────────────────────────────────────────────────────

    private async Task ServeLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            TcpClient client;
            try   { client = await _server.AcceptTcpClientAsync(ct); }
            catch { break; }

            _ = Task.Run(() => HandleClientAsync(client, ct), CancellationToken.None);
        }
    }

    private async Task HandleClientAsync(TcpClient client, CancellationToken ct)
    {
        using (client)
        {
            await using var stream = client.GetStream();
            using var reader = new System.IO.StreamReader(stream, Encoding.ASCII, leaveOpen: true);
            using var writer = new System.IO.StreamWriter(stream, Encoding.ASCII, leaveOpen: true)
            {
                NewLine    = "\r\n",
                AutoFlush  = false,
            };

            // Send welcome banner.
            await writer.WriteLineAsync("Welcome to the DataMan DMCC (MockServer)");
            await writer.FlushAsync(ct);

            while (!ct.IsCancellationRequested)
            {
                string? line;
                try   { line = await reader.ReadLineAsync(ct); }
                catch { break; }
                if (line is null) break;

                string cmd = line.Trim();
                if (string.IsNullOrEmpty(cmd)) continue;

                string body = _responses.TryGetValue(cmd, out var r) ? r : "";
                int status  = _responses.ContainsKey(cmd) ? 0 : DmccStatus.Busy;

                // Write response: blank line, status, blank line, body (if any).
                await writer.WriteLineAsync("");       // start marker
                await writer.WriteLineAsync(status.ToString());
                if (!string.IsNullOrEmpty(body))
                {
                    await writer.WriteLineAsync("");   // body separator
                    await writer.WriteLineAsync(body);
                }
                await writer.FlushAsync(ct);
            }
        }
    }

    public async ValueTask DisposeAsync()
    {
        _cts.Cancel();
        _server.Stop();
        if (_serveTask is not null)
        {
            try { await _serveTask; } catch { /* expected cancellation */ }
        }
        _cts.Dispose();
    }

    // Re-use status constant from main namespace.
    private static class DmccStatus { public const int Busy = 8; }
}
