<#
.SYNOPSIS
    Full DMCC parameter dump for a Cognex DataMan 475V unit.
    629 parameters extracted from the DataMan Communications Reference (dmcc-html/).
    Unsupported params are recorded as "(unsupported)" - not errors.

.USAGE
    .\Get-DmSettings.ps1
    .\Get-DmSettings.ps1 -DeviceIp 10.10.10.4
    .\Get-DmSettings.ps1 -DeviceIp 10.10.10.4 -OutputFile "C:\ref\factory-defaults.txt"

.NOTES
    Runtime: ~75 seconds for 629 params on local LAN.
    Diff two output files with: Compare-Object (gc file1.txt) (gc file2.txt)
#>
param(
    [string]$DeviceIp   = "10.10.10.4",
    [int]   $Port       = 23,
    [string]$OutputFile = ""
)

$params = @(
    # -- AZTEC -----------------------------------------------------------------
    "AZTEC.TRAINED-MODEL"
    # -- BEEP ------------------------------------------------------------------
    "BEEP.DATAVALID-FAIL"
    "BEEP.DATAVALID-FAIL-OL"
    "BEEP.GLOBAL-ENABLE"
    "BEEP.GOOD"
    "BEEP.GOOD-OL"
    "BEEP.NO-READ"
    "BEEP.NO-READ-OL"
    "BEEPER.ENABLED"
    # -- 1D SYMBOLOGIES --------------------------------------------------------
    "C11.CHKCHAR"
    "C11.CHKCHAR-OPTION"
    "C11.CODESIZE"
    "C11.QZ-SIZE"
    "C11.VERIFICATION"
    "C128.CODESIZE"
    "C128.QZ-SIZE"
    "C128.VERIFICATION"
    "C25.CHKCHAR"
    "C25.CODESIZE"
    "C25.QZ-SIZE"
    "C25.VERIFICATION"
    "C25.XMTCHK"
    "C39.ASCII"
    "C39.CHKCHAR"
    "C39.CODESIZE"
    "C39.QZ-SIZE"
    "C39.VERIFICATION"
    "C39.XMTCHK"
    "C93.CODESIZE"
    "C93.VERIFICATION"
    "CODABAR.CHKCHAR"
    "CODABAR.CODESIZE"
    "CODABAR.QZ-SIZE"
    "CODABAR.VERIFICATION"
    "CODABAR.XMTCHK"
    "UCC.CC-C"
    "UCC.XMTMODE"
    # -- CAMERA ----------------------------------------------------------------
    "CAMERA.AUTO-REGULATION"
    "CAMERA.BURST-LENGTH"
    "CAMERA.EXPOSURE"
    "CAMERA.EXPOSURE-LIMIT-US"
    "CAMERA.EXPOSURE-US"
    "CAMERA.FOV"
    "CAMERA.FOV-ENABLED"
    "CAMERA.FULL-BLACK-LEVEL-CORRECTION"
    "CAMERA.GAIN"
    "CAMERA.GAIN-LIMIT"
    "CAMERA.HDR-MODE"
    "CAMERA.INTERVAL-US"
    "CAMERA.MAX-EXPOSURE-US"
    "CAMERA.MIRROR-HORIZONTAL"
    "CAMERA.MIRROR-VERTICAL"
    "CAMERA.TARGET-BRIGHTNESS"
    "CAMERA.XPAND-ROI"
    "CAMERA.XPAND-ROI-TYPE"
    # -- CLIENT (TCP outbound) -------------------------------------------------
    "CLIENT.CLIENT-PORT"
    "CLIENT.CLOSE-CONNECTION"
    "CLIENT.CONNECTION-TIMEOUT"
    "CLIENT.ENABLED"
    "CLIENT.HOST-ADDRESS"
    "CLIENT.HOST-PORT"
    "CLIENT.IDLE-TIMEOUT"
    "CLIENT.LINGER-TIME"
    "CLIENT.OPEN-CONNECTION"
    "CLIENT.PROTOCOL"
    "CLIENT.RECONNECT-DELAY"
    # -- CMD -------------------------------------------------------------------
    "CMD.COMMAND-FOOTER"
    "CMD.COMMAND-HEADER"
    "CMD.ECHO"
    "CMD.TRIGGER"
    "CMD.TRIGGER-OFF"
    "CMD.TRIGGER-ON"
    # -- COM -------------------------------------------------------------------
    "COM.BAUD-RATE"
    "COM.BUFFER-AUTO-FLUSH"
    "COM.BUFFER-DATA"
    "COM.BUFFER-MODE"
    "COM.COMMUNICATIONS-MODULE"
    "COM.DATA-BITS"
    "COM.DMCC-CHECKSUM"
    "COM.DMCC-DEFAULT-TARGET"
    "COM.DMCC-HEADER"
    "COM.DMCC-RESPONSE"
    "COM.DMCC-TARGET"
    "COM.MULTI-PORT"
    "COM.PARITY"
    "COM.PROTOCOL"
    "COM.READER-IN-CRADLE"
    "COM.SCRIPT"
    "COM.SCRIPT-ENABLED"
    "COM.SCRIPT-ERROR"
    "COM.STOP-BITS"
    "COM.XLATEPRNT"
    # -- CODE QUALITY ----------------------------------------------------------
    "CODE.INFO"
    "CQ.CALIB-DATE"
    "CQ.ILLUMINATION"
    "CQ.METRICS"
    "CQ.MINPASS"
    "CQ.PROCESSM"
    "CQ.RESET-CALIBRATION"
    "CQ.UNITS"
    # -- DATA ------------------------------------------------------------------
    "DATA.IMAGE-TYPE"
    "DATA.RESULT-ALWAYSSEND"
    "DATA.RESULT-ENCODING"
    "DATA.RESULT-TYPE"
    # -- DATABAR ---------------------------------------------------------------
    "DATABAR.EXPANDED"
    "DATABAR.GROUP"
    "DATABAR.LIMITED"
    # -- DATAMATRIX ------------------------------------------------------------
    "DATAMATRIX.ALGORITHM-VERSION"
    "DATAMATRIX.DAMAGE"
    "DATAMATRIX.EXTREME-PRINT-GROWTH"
    "DATAMATRIX.LEARNING-NUMROWS-NUMCOLS"
    "DATAMATRIX.LEARNING-PARTIAL-MODEL"
    "DATAMATRIX.LEARNING-PARTIAL-MODEL-GRIDS"
    "DATAMATRIX.LOW-RES-2D"
    "DATAMATRIX.PERSPECTIVE-DISTORTION"
    "DATAMATRIX.PRINT-DEFORMATION"
    "DATAMATRIX.QUALITY-METRICS"
    "DATAMATRIX.SYMBOL-DIFFICULTY"
    "DATAMATRIX.TRAINED-FLEX-GRID"
    "DATAMATRIX.TRAINED-IGNORE-MIRROR"
    "DATAMATRIX.TRAINED-IGNORE-POLARITY"
    "DATAMATRIX.TRAINED-MODEL"
    "DATAMATRIX.USAGE"
    # -- DECODER ---------------------------------------------------------------
    "DECODER.AI-ENHANCER"
    "DECODER.ALLOW-ALL-MST-RESULTS"
    "DECODER.CENTERING-WINDOW"
    "DECODER.DISPLAY-TARGET"
    "DECODER.MAC-SCAN-TIMEOUT"
    "DECODER.NO-READ-FEEDBACK"
    "DECODER.REREAD-BASED-ON-SYMBOLOGY"
    "DECODER.REREAD-MODE"
    "DECODER.REREAD-NEVER2X"
    "DECODER.REREAD-NOT-LAST-N"
    "DECODER.REREAD-TIME"
    "DECODER.ROI"
    "DECODER.TARGET-DECODING"
    "DECODER.TIMEOUT"
    "DECODER.USE-CENTERING"
    # -- DETECTION -------------------------------------------------------------
    "DETECT.MAX-SPEED"
    "DETECT.MIN-SPEED"
    "DETECT.PHOTO-EYE-DISTANCE"
    "DETECTION.ACTIVE"
    "DETECTION.ENABLE"
    "DETECTION.SENSITIVITY"
    # -- DEVICE ----------------------------------------------------------------
    "DEVICE.BACKUP"
    "DEVICE.CHARGE-PROFILES-ENABLE"
    "DEVICE.DEFAULT"
    "DEVICE.DESCRIPTION"
    "DEVICE.FEATURE-KEYS"
    "DEVICE.FIRMWARE-VER"
    "DEVICE.LENS-SERIAL"
    "DEVICE.LOG"
    "DEVICE.MAC-ADDRESS"
    "DEVICE.NAME"
    "DEVICE.REBOOT"
    "DEVICE.RESTORE"
    "DEVICE.SERIAL-NUMBER"
    "DEVICE.TIMEZONE"
    "DEVICE.TYPE"
    "DEVICE.UPTIME"
    # -- DOTCODE ---------------------------------------------------------------
    "DOTCODE.TRAINED-MODEL"
    "DOTCODE.USAGE"
    # -- ETHERNET / NETWORK ----------------------------------------------------
    "ETH.DHCP"
    "ETH.GATEWAY"
    "ETH.IP-ADDRESS"
    "ETH.MAC-ADDRESS"
    "ETH.MTU"
    "ETH.SUBNET-MASK"
    "ETH.VLAN-ENABLE"
    "ETH.VLAN-ID"
    # -- EVENT -----------------------------------------------------------------
    "EVENT.ENABLED"
    "EVENT.PORT"
    "EVENT.RESULT"
    "EVENT.RESULT-ENCODING"
    # -- FILTER ----------------------------------------------------------------
    "FILTER.ENABLE"
    "FILTER.FILTER1"
    "FILTER.FILTER2"
    "FILTER.FILTER3"
    "FILTER.FILTER4"
    "FILTER.FILTER5"
    "FILTER.FILTER6"
    "FILTER.FILTER7"
    "FILTER.FILTER8"
    "FILTER.MODE"
    # -- FORMAT / FORMATTING ---------------------------------------------------
    "FORMAT.ENABLE"
    "FORMAT.FORMAT"
    # -- GENERAL ---------------------------------------------------------------
    "GENERAL.IP-ADDR"
    "GENERAL.IP-PORT"
    # -- HDR -------------------------------------------------------------------
    "HDR.MODE"
    # -- I2O5 ------------------------------------------------------------------
    "I2O5.CHKCHAR"
    "I2O5.CODESIZE"
    "I2O5.QZ-SIZE"
    "I2O5.VERIFICATION"
    "I2O5.XMTCHK"
    # -- IMAGE -----------------------------------------------------------------
    "IMAGE.FORMAT"
    "IMAGE.QUALITY"
    "IMAGE.SEND"
    "IMAGE.SIZE"
    "IMAGE.TRANSFER-FORMAT"
    # -- IO --------------------------------------------------------------------
    "IO.DEBOUNCE-TIME"
    "IO.INPUT-LOGIC"
    "IO.LINE-STATE"
    "IO.OUTPUT-LOGIC"
    "IO.OUTPUT-PULSE-DURATION"
    "IO.OUTPUT-STATE"
    "IO.TREND-ANALYSIS-FAIL"
    "IO.TREND-ANALYSIS-PASS"
    "IO.TREND-ANALYSIS-WARN"
    # -- LIVEIMG ---------------------------------------------------------------
    "LIVEIMG.MODE"
    "LIVEIMG.SEND"
    # -- MATCH -----------------------------------------------------------------
    "MATCH.MATCH-STRING"
    "MATCH.MASTER-DATABASE"
    "MATCH.MODE"
    "MATCH.STRING-COUNT"
    # -- MAXCODE ---------------------------------------------------------------
    "MAXCODE.TRAINED-MODEL"
    # -- MICRO-QR --------------------------------------------------------------
    "MICRO-QR.TRAINED-MODEL"
    # -- MST -------------------------------------------------------------------
    "MST.DMST-LISTEN-PORT"
    "MST.DMST-SEND-HOST"
    "MST.ENABLE"
    "MST.RESULT-ALWAYSSEND"
    "MST.RESULT-DETAIL"
    # -- NTP -------------------------------------------------------------------
    "NTP.ENABLE"
    "NTP.SERVER1"
    "NTP.SERVER2"
    # -- OCR -------------------------------------------------------------------
    "OCR.CHARSET"
    "OCR.ENABLE"
    "OCR.FONT"
    "OCR.LINESCAN-MODE"
    "OCR.MODE"
    "OCR.NOREAD-THRESHOLD"
    "OCR.REGION"
    "OCR.RESULT-DISPLAY"
    "OCR.TRAINED-MODEL"
    # -- PDF417 ----------------------------------------------------------------
    "PDF417.CODESIZE"
    "PDF417.MACRO"
    "PDF417.MICRO"
    "PDF417.TRAINED-MODEL"
    "PDF417.VERIFICATION"
    # -- POWERUP ---------------------------------------------------------------
    "POWERUP.LINE-STATE"
    "POWERUP.TRIGGER"
    # -- QR --------------------------------------------------------------------
    "QR.CODESIZE"
    "QR.MODEL"
    "QR.TRAINED-MODEL"
    "QR.USAGE"
    # -- RESULT ----------------------------------------------------------------
    "RESULT.DETAIL"
    "RESULT.DISPLAY"
    "RESULT.EXTERNAL"
    "RESULT.MULTICODE-MODE"
    "RESULT.MULTICODE-TIMEOUT"
    "RESULT.NOREAD-ENABLED"
    "RESULT.PARTIALLY-DECODED"
    "RESULT.SEND"
    "RESULT.SEPARATOR"
    "RESULT.SORT"
    "RESULT.STORE"
    # -- STATISTICS ------------------------------------------------------------
    "STATISTICS.CLEAR"
    "STATISTICS.ENABLED"
    "STATISTICS.FAIL-COUNT"
    "STATISTICS.GOOD-COUNT"
    "STATISTICS.MATCH"
    "STATISTICS.NOREAD-COUNT"
    "STATISTICS.PASS-COUNT"
    "STATISTICS.TOTAL-COUNT"
    "STATISTICS.TREND-WINDOW"
    "STATISTICS.WARN-COUNT"
    # -- SYMBOL ----------------------------------------------------------------
    "SYMBOL.C128"
    "SYMBOL.C39"
    "SYMBOL.C93"
    "SYMBOL.CODABAR"
    "SYMBOL.DATABAR"
    "SYMBOL.DATAMATRIX"
    "SYMBOL.DOTCODE"
    "SYMBOL.I2O5"
    "SYMBOL.PDF417"
    "SYMBOL.PHARMACODE"
    "SYMBOL.QR"
    "SYMBOL.UPC-EAN"
    # -- TRAIN -----------------------------------------------------------------
    "TRAIN.AUTO-DISABLE"
    "TRAIN.MAX-CODES"
    "TRAIN.MODE"
    "TRAIN.MULTICODE-ORDERING"
    "TRAIN.PITCH-ANGLE"
    "TRAIN.POLARITY"
    "TRAIN.POSITION-TOLERANCE"
    "TRAIN.TRAINED"
    # -- TRIGGER ---------------------------------------------------------------
    "TRIGGER.BURST-LENGTH"
    "TRIGGER.DELAY"
    "TRIGGER.ENABLED"
    "TRIGGER.END-FRAME"
    "TRIGGER.GATE-DELAY"
    "TRIGGER.GATE-ENABLED"
    "TRIGGER.GATE-LENGTH"
    "TRIGGER.IN-QUEUE"
    "TRIGGER.MAX-ON-TIME"
    "TRIGGER.MIN-INTERVAL"
    "TRIGGER.OVERLAP"
    "TRIGGER.START-FRAME"
    "TRIGGER.TIMEOUT"
    "TRIGGER.TYPE"
    # -- TRUCHECK --------------------------------------------------------------
    "TRUCHECK.APERTURE"
    "TRUCHECK.APERTURE-SIZE"
    "TRUCHECK.APPLICATION-CUSTOM-MAXIMUM-X-DIM"
    "TRUCHECK.APPLICATION-CUSTOM-MINIMUM-X-DIM"
    "TRUCHECK.APPLICATION-CUSTOM-PASS-GRADE"
    "TRUCHECK.APPLICATION-GS1-TABLE"
    "TRUCHECK.APPLICATION-STANDARD"
    "TRUCHECK.AUTO-BATCH"
    "TRUCHECK.BATCH-NUMBER"
    "TRUCHECK.CALIBRATE-CUSTOM-ON"
    "TRUCHECK.CALIBRATE-OFF"
    "TRUCHECK.CALIBRATE-ON"
    "TRUCHECK.CALIBRATION-DATE"
    "TRUCHECK.COMPANY-NAME"
    "TRUCHECK.CUSTOM-NOTE"
    "TRUCHECK.DECODE-GRADE"
    "TRUCHECK.DOT-PEEN"
    "TRUCHECK.FPD"
    "TRUCHECK.FPD-GRADE"
    "TRUCHECK.GNU"
    "TRUCHECK.GNU-GRADE"
    "TRUCHECK.GRADING-STANDARD"
    "TRUCHECK.ISO15415-VERSION"
    "TRUCHECK.ISO15416-VERSION"
    "TRUCHECK.ISO29158-VERSION"
    "TRUCHECK.METRIC-UNITS"
    "TRUCHECK.MOD-GRADE"
    "TRUCHECK.OPERATOR-NAME"
    "TRUCHECK.REPORT-SECTION"
    "TRUCHECK.REPORT-SECTION-ENABLED"
    "TRUCHECK.RM-GRADE"
    "TRUCHECK.SC"
    "TRUCHECK.SC-GRADE"
    "TRUCHECK.UEC"
    "TRUCHECK.UEC-GRADE"
    # -- TUNE ------------------------------------------------------------------
    "TUNE.EXCLUDE-AMBIENT-RESULTS"
    "TUNE.STATUS"
    "TUNE.TRAIN-CODE"
    "TUNE.TUNE-FILTERS"
    "TUNE.TUNE-LIGHT-BANKS"
    "TUNE.TUNE-LIGHT-COLORS"
    "TUNE.TUNE-LIGHT-COMBINATIONS"
    # -- UPC-EAN ---------------------------------------------------------------
    "UPC-EAN.SUPPLEMENT"
    "UPC-EAN.SUPPLEMENT-REQUIRED"
    "UPC-EAN.VERIFICATION"
    # -- VERIFICATION ----------------------------------------------------------
    "VERIFICATION.ENABLE"
    # -- WEB HMI / UI ----------------------------------------------------------
    "WEBHMI.ACTIONS"
    "WEBHMI.ENABLE"
    "WEBHMI.LANGUAGE"
    "WEBHMI.MATCH-STRING"
    "WEBHMI.SETTINGS"
    "WEBHMI.VERIFICATION"
    "WEBUI.ENABLE"
)

# -- Output file ---------------------------------------------------------------
if ($OutputFile -eq "") {
    $ts         = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $OutputFile = "DM475V-DPM_866D76_${ts}.txt"
}

# -- DMCC GET helper -----------------------------------------------------------
function Send-DmccGet {
    param($Stream, $Key)
    $cmd   = "||>GET $Key`r`n"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($cmd)
    $Stream.Write($bytes, 0, $bytes.Length)

    # Blocking read — ReadTimeout (500ms) is already set on the stream.
    # DataAvailable was unreliable: it returned false before data arrived.
    $buf = New-Object byte[] 4096
    $raw = ""
    try {
        $n = $Stream.Read($buf, 0, $buf.Length)
        if ($n -gt 0) { $raw = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n) }
    } catch { }

    # Drain any remaining chunks (multi-line responses: DEVICE.LOG, COM.SCRIPT etc.)
    $Stream.ReadTimeout = 80
    while ($true) {
        try {
            $n = $Stream.Read($buf, 0, $buf.Length)
            if ($n -gt 0) { $raw += [System.Text.Encoding]::ASCII.GetString($buf, 0, $n) }
            else { break }
        } catch { break }
    }
    $Stream.ReadTimeout = 500  # restore for next call

    $raw = $raw.Trim()
    if ($raw.Length -eq 0) { return "(no response)" }

    # Strip ACK prefix ||:::N[M] — keep the value that follows on the same or next line
    $raw = $raw -replace '\|\|:::\d+\[\d+\]', '' -replace '^\|\|>', '' -replace '\|\|>', ''
    $raw = $raw.Trim()
    if ($raw -eq "") { return "(no response)" }
    return $raw
}

# -- Connect -------------------------------------------------------------------
Write-Host "Connecting to $DeviceIp`:$Port ..." -ForegroundColor Cyan
$tcp = New-Object System.Net.Sockets.TcpClient
$tcp.Connect($DeviceIp, $Port)
$stream = $tcp.GetStream()
$stream.ReadTimeout = 500

# Drain banner
Start-Sleep -Milliseconds 400
if ($stream.DataAvailable) {
    $drain = New-Object byte[] 512
    $stream.Read($drain, 0, $drain.Length) | Out-Null
}

# Switch to Extended mode so every GET returns a value (default Silent mode returns nothing)
$modeCmd = [System.Text.Encoding]::ASCII.GetBytes("||>SET COM.DMCC-RESPONSE 2`r`n")
$stream.Write($modeCmd, 0, $modeCmd.Length)
Start-Sleep -Milliseconds 400
if ($stream.DataAvailable) {
    $drain = New-Object byte[] 512
    $stream.Read($drain, 0, $drain.Length) | Out-Null
}

$lines  = @()
$header = "# DM475V-DPM  866D76  $DeviceIp  fw:6.1.16_tc9  dump:$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$sep    = "# " + ("-" * 70)
$lines += $header
$lines += $sep

Write-Host $header -ForegroundColor Yellow
Write-Host $sep    -ForegroundColor Yellow

$total = $params.Count
$i     = 0
$currentPrefix = ""

foreach ($key in $params) {
    $i++
    $prefix = $key.Split('.')[0]
    if ($prefix -ne $currentPrefix) {
        $currentPrefix = $prefix
        $grpLine = "`n# -- $prefix " + ("-" * [Math]::Max(1, 60 - $prefix.Length))
        $lines  += $grpLine
        Write-Host $grpLine -ForegroundColor DarkGray
    }

    $val  = Send-DmccGet -Stream $stream -Key $key
    $line = "{0,-52} = {1}" -f $key, $val
    $lines += $line

    $pct = [int](($i / $total) * 100)
    Write-Progress -Activity "Reading DMCC parameters" -Status "$i / $total - $key" -PercentComplete $pct
    Write-Host $line
}

Write-Progress -Activity "Reading DMCC parameters" -Completed
$tcp.Close()

$lines | Out-File -FilePath $OutputFile -Encoding UTF8
Write-Host ""
Write-Host "Done. $total parameters queried." -ForegroundColor Green
Write-Host "Saved: $OutputFile"               -ForegroundColor Green
Write-Host ""
Write-Host "To diff against a future capture:"  -ForegroundColor Cyan
Write-Host "  Compare-Object (Get-Content '$OutputFile') (Get-Content 'future-dump.txt')" -ForegroundColor Cyan
