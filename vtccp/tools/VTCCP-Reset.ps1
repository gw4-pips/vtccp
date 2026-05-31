# VTCCP-Reset.ps1
# Factory-reset DM475-63530E-PIPS-Verif-Lab and restore all settings.
#
# Run from PowerShell (Administrator) on the verifier PC:
#   powershell -ExecutionPolicy Bypass -File VTCCP-Reset.ps1
#
# BEFORE RUNNING: close DMST completely so it does not hold port 23.

$DeviceIp      = "10.10.10.7"
$DevicePort    = 23
$RebootWaitSec = 45

# ---------- helpers ----------------------------------------------------------

function Connect-Device {
    param([int]$TimeoutSec = 10)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $c = New-Object System.Net.Sockets.TcpClient
            $c.Connect($DeviceIp, $DevicePort)
            $s = $c.GetStream()
            $s.ReadTimeout = 3000
            $w = New-Object System.IO.StreamWriter($s)
            $w.AutoFlush = $true
            return @{ C = $c; S = $s; W = $w }
        } catch {
            Start-Sleep -Seconds 1
        }
    }
    throw "Could not connect to $DeviceIp`:$DevicePort after $TimeoutSec seconds."
}

function Send-Dmcc {
    param($conn, [string]$cmd, [int]$delayMs = 400)
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
    return $ack
}

function Close-Device {
    param($conn)
    try { $conn.C.Close() } catch { }
}

# =============================================================================
# PHASE 1 -- read current values, issue CONFIG.DEFAULT + REBOOT
# =============================================================================

Write-Host ""
Write-Host "=== PHASE 1: read current values then CONFIG.DEFAULT ===" -ForegroundColor Cyan

$conn = Connect-Device
Start-Sleep -Milliseconds 600
Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 600 | Out-Null

$curType   = Send-Dmcc $conn "GET DATA.RESULT-TYPE"       500
$curAlways = Send-Dmcc $conn "GET DATA.RESULT-ALWAYSSEND" 500
Write-Host "  Current DATA.RESULT-TYPE       = $curType"   -ForegroundColor Yellow
Write-Host "  Current DATA.RESULT-ALWAYSSEND = $curAlways" -ForegroundColor Yellow

Write-Host ""
Write-Host "  Issuing CONFIG.DEFAULT + CONFIG.SAVE + REBOOT..." -ForegroundColor Red
Send-Dmcc $conn "CONFIG.DEFAULT" 2000 | Out-Null
Send-Dmcc $conn "CONFIG.SAVE"    2000 | Out-Null
Send-Dmcc $conn "REBOOT"         1000 | Out-Null
Close-Device $conn

Write-Host "  Waiting $RebootWaitSec seconds for device to reboot..." -ForegroundColor Cyan
Start-Sleep -Seconds $RebootWaitSec
Write-Host "  Wait complete."

# =============================================================================
# PHASE 2 -- reconnect (with retry), capture factory defaults, restore block
# =============================================================================

Write-Host ""
Write-Host "=== PHASE 2: reconnect + factory defaults + symbology restore ===" -ForegroundColor Cyan
Write-Host "  Connecting (will retry for up to 30 seconds)..."

$conn = Connect-Device -TimeoutSec 30
Start-Sleep -Milliseconds 800
Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 800 | Out-Null

$factTypeAck   = Send-Dmcc $conn "GET DATA.RESULT-TYPE"       600
$factAlwaysAck = Send-Dmcc $conn "GET DATA.RESULT-ALWAYSSEND" 600
$factTypeVal   = Extract-Value $factTypeAck
$factAlwaysVal = Extract-Value $factAlwaysAck

Write-Host "  Factory DATA.RESULT-TYPE       = $factTypeVal"   -ForegroundColor Green
Write-Host "  Factory DATA.RESULT-ALWAYSSEND = $factAlwaysVal" -ForegroundColor Green
Write-Host "  (Will be re-stamped after the .dmb load in Phase 4)" -ForegroundColor DarkGray

Write-Host ""
Write-Host "  Applying DMCC restore block..."

$restoreBlock = @(
    "SET TRAIN.AUTO-DISABLE ON",
    "SET SYMBOL.DATAMATRIX ON",
    "SET SYMBOL.QR ON",
    "SET SYMBOL.C128 ON",
    "SET SYMBOL.C93 ON",
    "SET SYMBOL.C39 ON",
    "SET SYMBOL.CODABAR ON",
    "SET SYMBOL.I2O5 ON",
    "SET SYMBOL.UPC-EAN ON",
    "SET SYMBOL.PDF417 ON",
    "SET SYMBOL.DATABAR ON",
    "SET CAMERA.MIRROR-HORIZONTAL ON",
    "SET CAMERA.MIRROR-VERTICAL ON",
    "SET DEVICE.TIMEZONE America/New_York",
    "SET NTP.ENABLE ON",
    "SET NTP.SERVER1 time.nist.gov",
    "SET TRUCHECK.COMPANY-NAME Product Identification and Processing Systems, Inc."
)

foreach ($cmd in $restoreBlock) {
    $ack = Send-Dmcc $conn $cmd 500
    Write-Host "    $cmd  -->  $ack"
}

# Explicitly stamp the factory DATA.RESULT values now, before the .dmb can
# overwrite them. Phase 4 will do this again after the .dmb load -- this is
# the safety net in case something goes wrong between here and Phase 4.
$x1 = Send-Dmcc $conn "SET DATA.RESULT-TYPE $factTypeVal"        600
$x2 = Send-Dmcc $conn "SET DATA.RESULT-ALWAYSSEND $factAlwaysVal" 600
Write-Host "  SET DATA.RESULT-TYPE $factTypeVal       --> $x1"
Write-Host "  SET DATA.RESULT-ALWAYSSEND $factAlwaysVal --> $x2"

Send-Dmcc $conn "CONFIG.SAVE" 2000 | Out-Null
Write-Host "  Phase 2 complete and saved." -ForegroundColor Green
Write-Host "  NOTE: the .dmb load in Phase 3 will overwrite DATA.RESULT-TYPE" -ForegroundColor DarkGray
Write-Host "        and DATA.RESULT-ALWAYSSEND back to 513. Phase 4 fixes that." -ForegroundColor DarkGray
Close-Device $conn

# =============================================================================
# PHASE 3 -- manual: load .dmb in DMST
# =============================================================================

Write-Host ""
Write-Host "=== PHASE 3: MANUAL STEP -- load the .dmb backup in DMST ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Open DMST and let it connect to the device"
Write-Host "  2. File -> Open Settings..."
Write-Host "     File: C:\Users\Administrator\Documents\DM475-63530E-PIPS-Verif-Lab\Settings Backups\6353OE Settings, 2026-05-31-1822.dmb"
Write-Host "  3. Click  Write Settings to Verifier  (blue arrow toolbar button)"
Write-Host "  4. Click  the floppy disk icon  (CONFIG.SAVE)"
Write-Host "  5. Wait for DMST to confirm the write is complete"
Write-Host "  6. CLOSE DMST before pressing Enter"
Write-Host ""
Read-Host "Press Enter when DMST has finished writing and you have closed DMST"

# =============================================================================
# PHASE 4 -- re-stamp factory DATA.RESULT values over .dmb residue
# =============================================================================

Write-Host ""
Write-Host "=== PHASE 4: re-applying factory DATA.RESULT values ===" -ForegroundColor Cyan
Write-Host "  Connecting..."

$conn = Connect-Device -TimeoutSec 15
Start-Sleep -Milliseconds 600
Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 600 | Out-Null

$a1 = Send-Dmcc $conn "SET DATA.RESULT-TYPE $factTypeVal"        600
$a2 = Send-Dmcc $conn "SET DATA.RESULT-ALWAYSSEND $factAlwaysVal" 600
$a3 = Send-Dmcc $conn "CONFIG.SAVE" 2000

Write-Host "  SET DATA.RESULT-TYPE $factTypeVal       --> $a1"
Write-Host "  SET DATA.RESULT-ALWAYSSEND $factAlwaysVal --> $a2"
Write-Host "  CONFIG.SAVE                             --> $a3"

Close-Device $conn

Write-Host ""
Write-Host "=== DONE ===" -ForegroundColor Green
Write-Host ""
Write-Host "  Open DMST and trigger a scan."
Write-Host "  Watch the TC panel -- the image should persist after the scan completes."
Write-Host ""
Write-Host "  KEY FINDINGS -- record these:"
Write-Host "    DATA.RESULT-TYPE       (factory) = $factTypeVal"
Write-Host "    DATA.RESULT-ALWAYSSEND (factory) = $factAlwaysVal"
Write-Host ""
Write-Host "  If the image persists: root cause confirmed -- SDK residue in those two keys."
Write-Host "  If the image still disappears: report the factory values above for further diagnosis."
Write-Host ""
