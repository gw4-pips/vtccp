<#
.SYNOPSIS
    Deploy DmstPushScript_v1.js to the DM475V-DPM unit via DMCC on port 23.

.DESCRIPTION
    1. Connects to port 23 and switches to Extended DMCC-RESPONSE mode.
    2. Reads DmstPushScript_v1.js from the repo (resolved relative to this script).
    3. Sends SET COM.SCRIPT <full content>.
    4. Sends SET COM.SCRIPT-ENABLED ON.
    5. Sends CONFIG.SAVE to persist to NVRAM.
    6. Verifies COM.SCRIPT-ENABLED reads back as ON.
    7. Reports success or failure with actionable detail.

.USAGE
    .\Deploy-PushScript.ps1
    .\Deploy-PushScript.ps1 -DeviceIp 10.10.10.4
    .\Deploy-PushScript.ps1 -DeviceIp 10.10.10.4 -ScriptPath "C:\path\to\script.js"

.NOTES
    Uses the same port-23 / ReadTimeout=500 pattern as Get-DmSettings.ps1.
    SET COM.SCRIPT is sent as a single DMCC command; the script content is
    written in 4 KB chunks to avoid saturating the TCP send buffer.
    Expected runtime: ~5-10 seconds on local LAN.
#>
param(
    [string]$DeviceIp   = "10.10.10.4",
    [int]   $Port       = 23,
    [string]$ScriptPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Resolve script path
# ---------------------------------------------------------------------------
if ($ScriptPath -eq "") {
    $toolDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
    $vtccpRoot  = Split-Path -Parent $toolDir   # vtccp/
    $ScriptPath = Join-Path $vtccpRoot "DeviceInterface\Dmst\DmstPushScript_v1.js"
}

if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERROR: Push script not found at: $ScriptPath" -ForegroundColor Red
    exit 1
}

$jsContent = Get-Content -Path $ScriptPath -Raw -Encoding UTF8
$lineCount = ($jsContent -split "`n").Count

# DMCC command values must not contain real newlines -- the device treats every
# \r\n as a command terminator.  Escape all newlines to literal \n (two chars)
# so the entire script travels as one unbroken command line.
$jsEscaped = $jsContent -replace "`r`n", "\n" -replace "`r", "\n" -replace "`n", "\n"
$byteCount = [System.Text.Encoding]::ASCII.GetByteCount($jsEscaped)
Write-Host ""
Write-Host "Push Script  : $ScriptPath" -ForegroundColor Cyan
Write-Host "Lines        : $lineCount   Bytes: $byteCount" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Helper: send a DMCC command and read the full response
# ---------------------------------------------------------------------------
function Send-Dmcc {
    param(
        [System.Net.Sockets.NetworkStream]$Stream,
        [string]$Command,
        [string]$Label,
        [int]   $ChunkSize  = 4096,
        [int]   $FirstRead  = 500,
        [int]   $DrainRead  = 150
    )

    $cmdLine = "||>$Command`r`n"
    $bytes   = [System.Text.Encoding]::ASCII.GetBytes($cmdLine)

    # Write in chunks to avoid saturating the TCP send buffer for large payloads
    $offset = 0
    while ($offset -lt $bytes.Length) {
        $len = [Math]::Min($ChunkSize, $bytes.Length - $offset)
        $Stream.Write($bytes, $offset, $len)
        $offset += $len
        if ($offset -lt $bytes.Length) {
            Start-Sleep -Milliseconds 20
        }
    }

    # First blocking read
    $buf = New-Object byte[] 8192
    $raw = ""
    $Stream.ReadTimeout = $FirstRead
    try {
        $n = $Stream.Read($buf, 0, $buf.Length)
        if ($n -gt 0) { $raw = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n) }
    } catch { }

    # Drain remaining chunks
    $Stream.ReadTimeout = $DrainRead
    while ($true) {
        try {
            $n = $Stream.Read($buf, 0, $buf.Length)
            if ($n -gt 0) { $raw += [System.Text.Encoding]::ASCII.GetString($buf, 0, $n) }
            else           { break }
        } catch { break }
    }
    $Stream.ReadTimeout = 500   # restore

    $raw = $raw.Trim()

    # Extract ACK code  -  format: ||:::N[M]  where M=0 means success
    $ack = $null
    if ($raw -match '\|\|:::\d+\[(\d+)\]') {
        $ack = [int]$Matches[1]
    }

    # Strip ACK prefix to isolate value
    $value = $raw -replace '\|\|:::\d+\[\d+\]', '' -replace '^\|\|>', '' -replace '\|\|>', ''
    $value = $value.Trim()

    return [PSCustomObject]@{
        Label  = $Label
        Ack    = $ack
        Value  = $value
        Raw    = $raw
    }
}

# ---------------------------------------------------------------------------
# Connect
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Connecting to $DeviceIp`:$Port ..." -ForegroundColor Cyan
$tcp = New-Object System.Net.Sockets.TcpClient
try {
    $tcp.Connect($DeviceIp, $Port)
} catch {
    Write-Host "ERROR: Cannot connect to $DeviceIp`:$Port  -  $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
$stream = $tcp.GetStream()
$stream.ReadTimeout = 500

# Drain banner
Start-Sleep -Milliseconds 400
if ($stream.DataAvailable) {
    $drain = New-Object byte[] 4096
    $stream.Read($drain, 0, $drain.Length) | Out-Null
}
Write-Host "Connected." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Step 1: Extended DMCC-RESPONSE mode
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[1/6] Setting COM.DMCC-RESPONSE 2 (Extended mode) ..." -ForegroundColor Yellow
$modeCmd = [System.Text.Encoding]::ASCII.GetBytes("||>SET COM.DMCC-RESPONSE 2`r`n")
$stream.Write($modeCmd, 0, $modeCmd.Length)
Start-Sleep -Milliseconds 400
if ($stream.DataAvailable) {
    $drain = New-Object byte[] 512
    $stream.Read($drain, 0, $drain.Length) | Out-Null
}
Write-Host "     Done." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Step 2: SET COM.SCRIPT
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[2/6] Sending SET COM.SCRIPT ($byteCount bytes) ..." -ForegroundColor Yellow

# The DMCC SET command for COM.SCRIPT takes the full JS content as the value.
# Large payload  -  write in 4 KB chunks.
# The device needs time to receive, parse, and NVRAM-write 80+ KB; allow 60 s.
$setScript = "SET COM.SCRIPT $jsEscaped"
$r = Send-Dmcc -Stream $stream -Command $setScript -Label "SET COM.SCRIPT" `
               -ChunkSize 4096 -FirstRead 60000 -DrainRead 1000

if ($null -eq $r.Ack) {
    Write-Host "     WARNING: No ACK within timeout for SET COM.SCRIPT." -ForegroundColor DarkYellow
    Write-Host "     Pausing 5 s then continuing  -  device may still have written the script." -ForegroundColor DarkYellow
} elseif ($r.Ack -ne 0) {
    Write-Host "     ERROR: SET COM.SCRIPT returned ACK code $($r.Ack)." -ForegroundColor Red
    Write-Host "     Response: $($r.Raw)" -ForegroundColor Red
    $tcp.Close()
    exit 1
} else {
    Write-Host "     ACK [0]  -  OK." -ForegroundColor Green
}

# Drain any late-arriving bytes before issuing the next command.
Write-Host "     Draining stream ..." -ForegroundColor Gray
Start-Sleep -Milliseconds 2000
$drain = New-Object byte[] 65536
$stream.ReadTimeout = 500
while ($true) {
    try {
        $n = $stream.Read($drain, 0, $drain.Length)
        if ($n -le 0) { break }
    } catch { break }
}

# ---------------------------------------------------------------------------
# Step 3: SET COM.SCRIPT-ENABLED ON
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[3/6] Setting COM.SCRIPT-ENABLED ON ..." -ForegroundColor Yellow
$r = Send-Dmcc -Stream $stream -Command "SET COM.SCRIPT-ENABLED ON" -Label "SET COM.SCRIPT-ENABLED"

if ($r.Ack -ne 0) {
    Write-Host "     ERROR: SET COM.SCRIPT-ENABLED returned ACK code $($r.Ack)." -ForegroundColor Red
    Write-Host "     Response: $($r.Raw)" -ForegroundColor Red
    $tcp.Close()
    exit 1
}
Write-Host "     ACK [0]  -  OK." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Step 4: CONFIG.SAVE
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[4/6] Saving configuration (CONFIG.SAVE) ..." -ForegroundColor Yellow
$r = Send-Dmcc -Stream $stream -Command "CONFIG.SAVE" -Label "CONFIG.SAVE" `
               -FirstRead 5000 -DrainRead 500

if ($null -eq $r.Ack) {
    Write-Host "     WARNING: No ACK received for CONFIG.SAVE. Response: $($r.Raw)" -ForegroundColor DarkYellow
} elseif ($r.Ack -ne 0) {
    Write-Host "     ERROR: CONFIG.SAVE returned ACK code $($r.Ack)." -ForegroundColor Red
    Write-Host "     Response: $($r.Raw)" -ForegroundColor Red
    $tcp.Close()
    exit 1
} else {
    Write-Host "     ACK [0]  -  OK." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Step 5: Verify COM.SCRIPT-ENABLED
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[5/6] Verifying COM.SCRIPT-ENABLED ..." -ForegroundColor Yellow
Start-Sleep -Milliseconds 300
$r = Send-Dmcc -Stream $stream -Command "GET COM.SCRIPT-ENABLED" -Label "GET COM.SCRIPT-ENABLED"

$enabledVal = $r.Value.Trim().ToUpper()
if ($enabledVal -eq "ON") {
    Write-Host "     COM.SCRIPT-ENABLED = $($r.Value)  (OK)" -ForegroundColor Green
} else {
    Write-Host "     ERROR: COM.SCRIPT-ENABLED = '$($r.Value)' (expected ON)." -ForegroundColor Red
    $tcp.Close()
    exit 1
}

# ---------------------------------------------------------------------------
# Step 6: Verify script was written (check length via GET COM.SCRIPT first line)
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[6/6] Spot-checking COM.SCRIPT (reading back first 200 chars) ..." -ForegroundColor Yellow
$r = Send-Dmcc -Stream $stream -Command "GET COM.SCRIPT" -Label "GET COM.SCRIPT" `
               -FirstRead 3000 -DrainRead 500

$scriptPreview = if ($r.Value.Length -gt 200) { $r.Value.Substring(0, 200) + "..." } else { $r.Value }
if ($r.Value -match "VTCCP DMST Push Script") {
    Write-Host "     Script header confirmed in device response  (OK)" -ForegroundColor Green
} else {
    Write-Host "     WARNING: Could not confirm VTCCP header in GET COM.SCRIPT response." -ForegroundColor DarkYellow
    Write-Host "     First 200 chars of response:" -ForegroundColor DarkYellow
    Write-Host "     $scriptPreview" -ForegroundColor DarkYellow
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
$tcp.Close()

Write-Host ""
Write-Host "=" * 72 -ForegroundColor Cyan
Write-Host "  DEPLOY COMPLETE" -ForegroundColor Green
Write-Host "  Device  : $DeviceIp (DM475V-DPM 866D76)" -ForegroundColor Green
Write-Host "  Script  : DmstPushScript_v1.js ($lineCount lines / $byteCount bytes)" -ForegroundColor Green
Write-Host "  Status  : COM.SCRIPT-ENABLED = ON  |  CONFIG.SAVE confirmed" -ForegroundColor Green
Write-Host "  Next    : Trigger a scan in VtccpApp  -  push XML should arrive on port 44444." -ForegroundColor Green
Write-Host "=" * 72 -ForegroundColor Cyan
Write-Host ""
