# Test-LiveImgSend.ps1
# Probes LIVEIMG.MODE=2 + LIVEIMG.SEND on the DM475V.
# Saves the response body as liveimg_response.bin (inspect for JPEG magic FF D8 FF).
# Always restores LIVEIMG.MODE=0 before exit, even on error.
# NEVER calls COM.DMCC-SAVE.
#
# Usage: .\Test-LiveImgSend.ps1 [-Host 10.10.10.7] [-Port 23]

param(
    [string]$DeviceHost = "10.10.10.7",
    [int]$Port = 23,
    [string]$OutFile = "liveimg_response.bin"
)

$ErrorActionPreference = "Stop"

function Send-Dmcc {
    param($Stream, $Command)
    $bytes = [System.Text.Encoding]::ASCII.GetBytes("||>$Command`r`n")
    $Stream.Write($bytes, 0, $bytes.Length)
    $Stream.Flush()
    Write-Host "  SENT: ||>$Command"
}

function Read-DmccResponse {
    param($Stream, [int]$TimeoutMs = 3000)
    $buf = New-Object byte[] 65536
    $response = [System.Text.StringBuilder]::new()
    $deadline = [System.Diagnostics.Stopwatch]::StartNew()
    while ($deadline.ElapsedMilliseconds -lt $TimeoutMs) {
        if ($Stream.DataAvailable) {
            $n = $Stream.Read($buf, 0, $buf.Length)
            if ($n -gt 0) {
                $response.Append([System.Text.Encoding]::ASCII.GetString($buf, 0, $n)) | Out-Null
                # Stop when we see a bare CRLF line (end of DMCC response)
                if ($response.ToString() -match "\r\n\r\n|\|\|\r\n") { break }
            }
        }
        Start-Sleep -Milliseconds 50
    }
    return $response.ToString()
}

function Read-BinaryResponse {
    param($Stream, [int]$TimeoutMs = 5000)
    # Read raw bytes until no more data for 500ms
    $all = [System.Collections.Generic.List[byte]]::new()
    $buf = New-Object byte[] 65536
    $deadline = [System.Diagnostics.Stopwatch]::StartNew()
    $lastData = [System.Diagnostics.Stopwatch]::StartNew()
    while ($deadline.ElapsedMilliseconds -lt $TimeoutMs) {
        if ($Stream.DataAvailable) {
            $n = $Stream.Read($buf, 0, $buf.Length)
            if ($n -gt 0) {
                $all.AddRange($buf[0..($n-1)])
                $lastData.Restart()
            }
        } elseif ($lastData.ElapsedMilliseconds -gt 500 -and $all.Count -gt 0) {
            break
        }
        Start-Sleep -Milliseconds 20
    }
    return $all.ToArray()
}

$tcp = $null
$stream = $null

try {
    Write-Host ""
    Write-Host "=== LIVEIMG.SEND probe ===" -ForegroundColor Cyan
    Write-Host "Device: $DeviceHost`:$Port"
    Write-Host ""

    $tcp = New-Object System.Net.Sockets.TcpClient
    $tcp.Connect($DeviceHost, $Port)
    $stream = $tcp.GetStream()
    $stream.ReadTimeout = 5000
    Write-Host "Connected." -ForegroundColor Green

    # Step 1 — enable live image mode
    Write-Host "`n[1] SET LIVEIMG.MODE 2"
    Send-Dmcc $stream "SET LIVEIMG.MODE 2"
    Start-Sleep -Milliseconds 300
    $r1 = Read-DmccResponse $stream
    Write-Host "  RESPONSE: $($r1.Trim())"

    if ($r1 -match "Error|error|ERR") {
        Write-Host "  *** Device rejected LIVEIMG.MODE=2 — may not be supported on this firmware." -ForegroundColor Yellow
    } else {
        Write-Host "  OK" -ForegroundColor Green
    }

    # Step 2 — request one frame: full size (0), JPEG (1), quality 85
    Write-Host "`n[2] LIVEIMG.SEND 0 1 85"
    Send-Dmcc $stream "LIVEIMG.SEND 0 1 85"
    Start-Sleep -Milliseconds 200

    # Read raw response (may be binary JPEG)
    $rawBytes = Read-BinaryResponse $stream 5000
    Write-Host "  Received $($rawBytes.Length) bytes"

    if ($rawBytes.Length -gt 3) {
        $magic = "{0:X2} {1:X2} {2:X2}" -f $rawBytes[0], $rawBytes[1], $rawBytes[2]
        Write-Host "  First 3 bytes: $magic"

        if ($rawBytes[0] -eq 0xFF -and $rawBytes[1] -eq 0xD8 -and $rawBytes[2] -eq 0xFF) {
            Write-Host "  *** JPEG MAGIC DETECTED — unencrypted image!" -ForegroundColor Green
            [System.IO.File]::WriteAllBytes($OutFile, $rawBytes)
            Write-Host "  Saved to: $OutFile" -ForegroundColor Green
        } elseif ($rawBytes.Length -lt 50) {
            $ascii = [System.Text.Encoding]::ASCII.GetString($rawBytes)
            Write-Host "  Short response (likely DMCC error): $ascii" -ForegroundColor Yellow
        } else {
            Write-Host "  Not a JPEG — encrypted or unknown format." -ForegroundColor Yellow
            [System.IO.File]::WriteAllBytes($OutFile, $rawBytes)
            Write-Host "  Raw bytes saved to: $OutFile (inspect manually)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  No data returned — command may not be supported at LIVEIMG.MODE=2." -ForegroundColor Yellow
        Write-Host "  (DMCC reference says 'mode 3' — will try that next if needed)" -ForegroundColor Gray
    }

} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
} finally {
    # Always restore LIVEIMG.MODE to 0
    if ($stream -ne $null -and $tcp.Connected) {
        Write-Host "`n[3] Restoring LIVEIMG.MODE=0 (safety restore)"
        try {
            Send-Dmcc $stream "SET LIVEIMG.MODE 0"
            Start-Sleep -Milliseconds 300
            $r3 = Read-DmccResponse $stream 2000
            Write-Host "  RESPONSE: $($r3.Trim())"
        } catch {
            Write-Host "  (restore failed — device may have already reset)" -ForegroundColor Yellow
        }
        $stream.Close()
        $tcp.Close()
        Write-Host "Disconnected."
    }
    Write-Host ""
    Write-Host "=== Done ===" -ForegroundColor Cyan
}
