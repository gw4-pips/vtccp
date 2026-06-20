param(
    [string]$DeviceIp = "10.10.10.7",
    [int]   $Port     = 23
)

# Fix-LiveimgMode.ps1
# Reads the 7 NVRAM parameters whose known-good values are documented in
# WORKING-NOTES.md, compares each to the expected value, and reports deviations.
#
# If LIVEIMG.MODE specifically is wrong it will be corrected in place.
# For any other deviation the script recommends CONFIG.DEFAULT, which is the
# only confirmed fix for the broader NVRAM corruption pattern.
#
# Safe to run with DMST open.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File Fix-LiveimgMode.ps1
#   powershell -ExecutionPolicy Bypass -File Fix-LiveimgMode.ps1 -DeviceIp 10.10.10.7

$ErrorActionPreference = "Stop"

# Known-good values (confirmed 2026-05-31, DM-KnownGood-Snapshot)
$KnownGood = [ordered]@{
    "LIVEIMG.MODE"            = "0"
    "DATA.IMAGE-TYPE"         = "0"
    "DATA.RESULT-TYPE"        = "513"
    "DATA.RESULT-ENCODING"    = "0"
    "DATA.RESULT-ALWAYSSEND"  = "1"
    "IMAGE.FORMAT"            = "1"
    "IMAGE.SIZE"              = "1"
}

function Connect-Device {
    $c = New-Object System.Net.Sockets.TcpClient
    $c.Connect($DeviceIp, $Port)
    $s = $c.GetStream()
    $s.ReadTimeout = 3000
    $w = New-Object System.IO.StreamWriter($s)
    $w.AutoFlush = $true
    return @{ C = $c; S = $s; W = $w }
}

function Send-Dmcc {
    param($conn, [string]$cmd, [int]$delayMs = 450)
    $conn.W.WriteLine("||>$cmd")
    Start-Sleep -Milliseconds $delayMs
    $buf = New-Object byte[] 4096
    $raw = ""
    try {
        if ($conn.S.DataAvailable) {
            $n   = $conn.S.Read($buf, 0, $buf.Length)
            $raw = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n).Trim()
        }
    } catch { }
    return $raw
}

function Extract-Value {
    param([string]$ack)
    if ($ack -match '\[0\](.+)') { return $Matches[1].Trim() }
    return $ack.Trim()
}

Write-Host ""
Write-Host "=== VTCCP device diagnostic — known-good parameter check ===" -ForegroundColor Cyan
Write-Host "Device  : $DeviceIp`:$Port"
Write-Host "Snapshot: DM-KnownGood 2026-05-31"
Write-Host ""

$conn = Connect-Device
Start-Sleep -Milliseconds 400

# drain banner
$buf0 = New-Object byte[] 256
if ($conn.S.DataAvailable) { $conn.S.Read($buf0, 0, $buf0.Length) | Out-Null }

Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 600 | Out-Null

$readings   = @{}
$deviations = @()

foreach ($key in $KnownGood.Keys) {
    $ack = Send-Dmcc $conn "GET $key" 500
    $val = Extract-Value $ack
    $readings[$key] = $val
    $expected = $KnownGood[$key]
    $ok = ($val -eq $expected)
    $status = if ($ok) { "OK  " } else { "FAIL" }
    $color  = if ($ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1,-30} = {2,-10}  (expected {3})" -f $status, $key, $val, $expected) -ForegroundColor $color
    if (-not $ok) { $deviations += $key }
}

Write-Host ""

if ($deviations.Count -eq 0) {
    Write-Host "All parameters match known-good values." -ForegroundColor Green
    Write-Host ""
    Write-Host "If the TC panel image is still blank, the root cause is NOT one of these" -ForegroundColor Yellow
    Write-Host "7 parameters.  Run the CONFIG.DEFAULT recovery:" -ForegroundColor Yellow
    Write-Host "  1. Telnet/nc 10.10.10.7 23" -ForegroundColor Yellow
    Write-Host "  2. CONFIG.DEFAULT   (then CONFIG.SAVE, then REBOOT)" -ForegroundColor Yellow
    Write-Host "  or use VTCCP-Reset.ps1 in this same folder." -ForegroundColor Yellow

} elseif ($deviations -contains "LIVEIMG.MODE" -and $deviations.Count -eq 1) {
    Write-Host "LIVEIMG.MODE is wrong — correcting in place..." -ForegroundColor Yellow
    $setAck = Send-Dmcc $conn "SET LIVEIMG.MODE 0" 600
    $verAck = Send-Dmcc $conn "GET LIVEIMG.MODE"   500
    $newVal = Extract-Value $verAck
    if ($newVal -eq "0") {
        $saveAck = Send-Dmcc $conn "COM.DMCC-SAVE" 2000
        Write-Host "  SET + verify OK.  COM.DMCC-SAVE: $saveAck" -ForegroundColor Green
        Write-Host "  Trigger a scan in DMST and confirm the image persists." -ForegroundColor Green
    } else {
        Write-Host "  ERROR: readback after SET = '$newVal'.  Run VTCCP-Reset.ps1." -ForegroundColor Red
    }

} else {
    Write-Host "DEVIATIONS FOUND in $($deviations.Count) parameter(s): $($deviations -join ', ')" -ForegroundColor Red
    Write-Host ""
    Write-Host "A targeted SET may not fix NVRAM corruption.  Recommended: CONFIG.DEFAULT reset." -ForegroundColor Yellow
    Write-Host "Run VTCCP-Reset.ps1 (in this folder) for the full 4-phase recovery." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Or to run manually:" -ForegroundColor Yellow
    Write-Host "  Telnet/nc $DeviceIp 23" -ForegroundColor Yellow
    Write-Host "  CONFIG.DEFAULT" -ForegroundColor Yellow
    Write-Host "  CONFIG.SAVE" -ForegroundColor Yellow
    Write-Host "  REBOOT" -ForegroundColor Yellow
}

try { $conn.C.Close() } catch { }
Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host ""
