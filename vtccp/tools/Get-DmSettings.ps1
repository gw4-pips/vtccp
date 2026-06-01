param(
    [string]$DeviceIp   = "10.10.10.7",
    [int]   $Port       = 23,
    [string]$OutputFile = ""
)

$params = @(
    "LIVEIMG.MODE",
    "DATA.IMAGE-TYPE",
    "DATA.RESULT-TYPE",
    "DATA.RESULT-ENCODING",
    "DATA.RESULT-ALWAYSSEND",
    "COM.DMCC-RESPONSE",
    "COM.DMCC-CHECKSUM",
    "COM.DMCC-HEADER",
    "IMAGE.FORMAT",
    "IMAGE.SIZE",
    "TRIGGER.TYPE",
    "CAMERA.MIRROR-HORIZONTAL",
    "CAMERA.MIRROR-VERTICAL",
    "DEVICE.TYPE",
    "DEVICE.FIRMWARE-VER",
    "DEVICE.NAME",
    "DEVICE.SERIAL-NUMBER",
    "TRAIN.AUTO-DISABLE",
    "UPC-EAN.SUPPLEMENT",
    "SYMBOL.DATAMATRIX",
    "SYMBOL.QR",
    "SYMBOL.C128",
    "SYMBOL.C93",
    "SYMBOL.C39",
    "SYMBOL.CODABAR",
    "SYMBOL.I2O5",
    "SYMBOL.UPC-EAN",
    "SYMBOL.PDF417",
    "SYMBOL.DATABAR",
    "TRUCHECK.APPLICATION-STANDARD",
    "TRUCHECK.GRADING-STANDARD",
    "TRUCHECK.APERTURE",
    "TRUCHECK.APERTURE-SIZE",
    "TRUCHECK.DOT-PEEN",
    "TRUCHECK.METRIC-UNITS",
    "TRUCHECK.APPLICATION-CUSTOM-PASS-GRADE",
    "TRUCHECK.APPLICATION-CUSTOM-MINIMUM-X-DIM",
    "TRUCHECK.APPLICATION-CUSTOM-MAXIMUM-X-DIM",
    "TRUCHECK.AUTO-BATCH",
    "TRUCHECK.BATCH-NUMBER",
    "TRUCHECK.OPERATOR-NAME",
    "TRUCHECK.COMPANY-NAME",
    "TRUCHECK.CUSTOM-NOTE",
    "NTP.ENABLE",
    "NTP.SERVER1",
    "NTP.SERVER2",
    "DEVICE.TIMEZONE"
)

if ($OutputFile -eq "") {
    $ts = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $OutputFile = "DM-Settings_$ts.txt"
}

function Send-DmccGet {
    param($Stream, $Key)
    $cmd = "||>GET $Key`r`n"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($cmd)
    $Stream.Write($bytes, 0, $bytes.Length)
    Start-Sleep -Milliseconds 180
    $buf = New-Object byte[] 512
    $n = 0
    if ($Stream.DataAvailable) {
        $n = $Stream.Read($buf, 0, $buf.Length)
    }
    if ($n -gt 0) {
        $raw = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n).Trim()
        $raw = $raw -replace '\|\|:::.*', '' -replace '^\|\|>', '' 
        return $raw.Trim()
    }
    return "(no response)"
}

Write-Host "Connecting to $DeviceIp`:$Port ..." -ForegroundColor Cyan
$tcp    = New-Object System.Net.Sockets.TcpClient
$tcp.Connect($DeviceIp, $Port)
$stream = $tcp.GetStream()
$stream.ReadTimeout = 400

Start-Sleep -Milliseconds 300
if ($stream.DataAvailable) {
    $drain = New-Object byte[] 256
    $stream.Read($drain, 0, $drain.Length) | Out-Null
}

$lines  = @()
$header = "# DM Settings dump — $DeviceIp — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$lines += $header
Write-Host $header -ForegroundColor Yellow
Write-Host ("-" * 60) -ForegroundColor Yellow

foreach ($key in $params) {
    $val  = Send-DmccGet -Stream $stream -Key $key
    $line = "{0,-50} = {1}" -f $key, $val
    $lines += $line
    Write-Host $line
}

$tcp.Close()

$lines | Out-File -FilePath $OutputFile -Encoding UTF8
Write-Host ""
Write-Host "Saved to: $OutputFile" -ForegroundColor Green
Write-Host "Run again after reset and diff the two files." -ForegroundColor Cyan
