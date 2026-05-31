# VTCCP-Reset.ps1
# Factory-reset the DM475-63530E-PIPS-Verif-Lab and restore all settings,
# while capturing the post-reset factory values for DATA.RESULT-TYPE and
# DATA.RESULT-ALWAYSSEND BEFORE the .dmb backup overwrites them.
#
# Run from PowerShell (Administrator) on the verifier PC:
#   powershell -ExecutionPolicy Bypass -File VTCCP-Reset.ps1
#
# BEFORE RUNNING:
#   - Close DMST completely (it holds port 23)
#   - Have your .dmb backup ready (DMST File -> Save Settings)

$DeviceIp      = "10.10.10.7"
$DevicePort    = 23
$RebootWaitSec = 38     # increase to 45 if Phase 2 reconnect fails

# ── helpers ───────────────────────────────────────────────────────────────────

function Connect-Device {
    $c = New-Object System.Net.Sockets.TcpClient
    $c.Connect($DeviceIp, $DevicePort)
    $s = $c.GetStream()
    $s.ReadTimeout = 3000
    $w = New-Object System.IO.StreamWriter($s)
    $w.AutoFlush = $true
    return @{ C = $c; S = $s; W = $w }
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
    # ACK format:  ||:::1[0]<value>  — strip everything up to and including [0]
    if ($ack -match '\[0\](.+)') { return $Matches[1].Trim() }
    return $ack   # fallback: return raw if format unexpected
}

function Close-Device {
    param($conn)
    try { $conn.C.Close() } catch { }
}

# =============================================================================
# PHASE 1  —  read current (corrupted) values, then factory-reset + reboot
# =============================================================================
Write-Host "`n=== PHASE 1: reading current values then issuing CONFIG.DEFAULT ===" -ForegroundColor Cyan

$conn = Connect-Device
Start-Sleep -Milliseconds 600
Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 600 | Out-Null

$curType   = Send-Dmcc $conn "GET DATA.RESULT-TYPE"       500
$curAlways = Send-Dmcc $conn "GET DATA.RESULT-ALWAYSSEND" 500
Write-Host "  Current DATA.RESULT-TYPE       = $curType"   -ForegroundColor Yellow
Write-Host "  Current DATA.RESULT-ALWAYSSEND = $curAlways" -ForegroundColor Yellow

Write-Host "`n  Issuing CONFIG.DEFAULT + CONFIG.SAVE + REBOOT..." -ForegroundColor Red
Send-Dmcc $conn "CONFIG.DEFAULT" 1200 | Out-Null
Send-Dmcc $conn "CONFIG.SAVE"    1200 | Out-Null
Send-Dmcc $conn "REBOOT"          500 | Out-Null
Close-Device $conn

Write-Host "  Device rebooting — waiting $RebootWaitSec seconds..." -ForegroundColor Cyan
for ($i = $RebootWaitSec; $i -gt 0; $i--) {
    Write-Host -NoNewline "`r  $i s remaining...   "
    Start-Sleep -Seconds 1
}
Write-Host ""

# =============================================================================
# PHASE 2  —  reconnect, capture factory defaults, apply DMCC restore block
# =============================================================================
Write-Host "`n=== PHASE 2: reconnect + read factory defaults + restore symbologies ===" -ForegroundColor Cyan

$conn = Connect-Device
Start-Sleep -Milliseconds 600
Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 600 | Out-Null

$factTypeAck   = Send-Dmcc $conn "GET DATA.RESULT-TYPE"       500
$factAlwaysAck = Send-Dmcc $conn "GET DATA.RESULT-ALWAYSSEND" 500
$factTypeVal   = Extract-Value $factTypeAck
$factAlwaysVal = Extract-Value $factAlwaysAck

Write-Host "  Factory DATA.RESULT-TYPE       = $factTypeVal"   -ForegroundColor Green
Write-Host "  Factory DATA.RESULT-ALWAYSSEND = $factAlwaysVal" -ForegroundColor Green
Write-Host "  (These will be re-stamped after the .dmb load in Phase 4)" -ForegroundColor DarkGray

Write-Host "`n  Applying DMCC restore block (symbologies / mirror / NTP / timezone)..."
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
    $ack = Send-Dmcc $conn $cmd 400
    Write-Host "    $cmd  -->  $ack"
}

Send-Dmcc $conn "CONFIG.SAVE" 1200 | Out-Null
Write-Host "  DMCC restore block saved." -ForegroundColor Green
Close-Device $conn

# =============================================================================
# PHASE 3  —  human step: load .dmb in DMST
# =============================================================================
Write-Host @"

=== PHASE 3: MANUAL STEP — load your .dmb backup in DMST ================

  1. Open DMST and let it connect to the device
  2. File  ->  Open Settings...  ->  browse to your .dmb file:
       C:\Users\Administrator\Documents\DM475-63530E-PIPS-Verif-Lab\Settings Backups\6353OE Settings, 2026-05-31-1822.dmb
  3. Click  Write Settings to Verifier  (blue arrow toolbar button)
  4. Click  the floppy disk icon  (CONFIG.SAVE)
  5. Wait for DMST to confirm the write is complete
  6. Close DMST again before pressing Enter below

"@ -ForegroundColor Yellow

Read-Host "Press Enter when DMST has finished writing the .dmb and you have closed DMST"

# =============================================================================
# PHASE 4  —  reconnect and stamp factory DATA.RESULT values back over .dmb residue
# =============================================================================
Write-Host "`n=== PHASE 4: re-applying factory DATA.RESULT values ===" -ForegroundColor Cyan

$conn = Connect-Device
Start-Sleep -Milliseconds 600
Send-Dmcc $conn "SET COM.DMCC-RESPONSE 2" 600 | Out-Null

$a1 = Send-Dmcc $conn "SET DATA.RESULT-TYPE $factTypeVal"       500
$a2 = Send-Dmcc $conn "SET DATA.RESULT-ALWAYSSEND $factAlwaysVal" 500
$a3 = Send-Dmcc $conn "CONFIG.SAVE" 1200

Write-Host "  SET DATA.RESULT-TYPE $factTypeVal          -->  $a1"
Write-Host "  SET DATA.RESULT-ALWAYSSEND $factAlwaysVal  -->  $a2"
Write-Host "  CONFIG.SAVE                                -->  $a3"

Close-Device $conn

Write-Host @"

=== DONE ====================================================================

  Open DMST and trigger a scan. Watch the TC panel — the image should
  persist after the scan completes.

  KEY FINDINGS — record these:
    DATA.RESULT-TYPE       (factory) = $factTypeVal
    DATA.RESULT-ALWAYSSEND (factory) = $factAlwaysVal

  If the image persists: root cause confirmed — SDK residue in those two keys.
  If the image still disappears: report the factory values above and we diagnose further.

"@ -ForegroundColor Green
