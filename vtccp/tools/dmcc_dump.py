"""
dmcc_dump.py — full DMCC parameter dump for DM475V on port 23
Usage: python dmcc_dump.py [ip] [port]
Default: 10.10.10.4  port 23
Output: dmcc_dump_YYYYMMDD_HHMMSS.txt  (same format as baseline reference files)
"""

import socket, time, sys, datetime

IP   = sys.argv[1] if len(sys.argv) > 1 else "10.10.10.4"
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 23

PARAMS = [
    ("AZTEC",       ["TRAINED-MODEL"]),
    ("BEEP",        ["DATAVALID-FAIL","DATAVALID-FAIL-OL","GLOBAL-ENABLE","GOOD","GOOD-OL","NO-READ","NO-READ-OL"]),
    ("BEEPER",      ["ENABLED"]),
    ("C11",         ["CHKCHAR","CHKCHAR-OPTION","CODESIZE","QZ-SIZE","VERIFICATION"]),
    ("C128",        ["CODESIZE","QZ-SIZE","VERIFICATION"]),
    ("C25",         ["CHKCHAR","CODESIZE","QZ-SIZE","VERIFICATION","XMTCHK"]),
    ("C39",         ["ASCII","CHKCHAR","CODESIZE","QZ-SIZE","VERIFICATION","XMTCHK"]),
    ("C93",         ["CODESIZE","VERIFICATION"]),
    ("CODABAR",     ["CHKCHAR","CODESIZE","QZ-SIZE","VERIFICATION","XMTCHK"]),
    ("UCC",         ["CC-C","XMTMODE"]),
    ("CAMERA",      ["AUTO-REGULATION","BURST-LENGTH","EXPOSURE","EXPOSURE-LIMIT-US",
                     "EXPOSURE-US","FOV","FOV-ENABLED","FULL-BLACK-LEVEL-CORRECTION",
                     "GAIN","GAIN-LIMIT","HDR-MODE","INTERVAL-US","MAX-EXPOSURE-US",
                     "MIRROR-HORIZONTAL","MIRROR-VERTICAL","TARGET-BRIGHTNESS",
                     "XPAND-ROI","XPAND-ROI-TYPE"]),
    ("CLIENT",      ["CLIENT-PORT","CLOSE-CONNECTION","CONNECTION-TIMEOUT","ENABLED",
                     "HOST-ADDRESS","HOST-PORT","IDLE-TIMEOUT","LINGER-TIME",
                     "OPEN-CONNECTION","PROTOCOL","RECONNECT-DELAY"]),
    ("CMD",         ["COMMAND-FOOTER","COMMAND-HEADER","ECHO","TRIGGER","TRIGGER-OFF","TRIGGER-ON"]),
    ("COM",         ["BAUD-RATE","BUFFER-AUTO-FLUSH","BUFFER-DATA","BUFFER-MODE",
                     "COMMUNICATIONS-MODULE","DATA-BITS","DMCC-CHECKSUM",
                     "DMCC-DEFAULT-TARGET","DMCC-HEADER","DMCC-RESPONSE","DMCC-TARGET",
                     "MULTI-PORT","PARITY","PROTOCOL","READER-IN-CRADLE",
                     "SCRIPT","SCRIPT-ENABLED","SCRIPT-ERROR","STOP-BITS","XLATEPRNT"]),
    ("CODE",        ["INFO"]),
    ("CQ",          ["CALIB-DATE","ILLUMINATION","METRICS","MINPASS","PROCESSM",
                     "RESET-CALIBRATION","UNITS"]),
    ("DATA",        ["IMAGE-TYPE","RESULT-ALWAYSSEND","RESULT-ENCODING","RESULT-TYPE"]),
    ("DATABAR",     ["EXPANDED","GROUP","LIMITED"]),
    ("DATAMATRIX",  ["ALGORITHM-VERSION","DAMAGE","EXTREME-PRINT-GROWTH",
                     "LEARNING-NUMROWS-NUMCOLS","LEARNING-PARTIAL-MODEL",
                     "LEARNING-PARTIAL-MODEL-GRIDS","LOW-RES-2D","PERSPECTIVE-DISTORTION",
                     "PRINT-DEFORMATION","QUALITY-METRICS","SYMBOL-DIFFICULTY",
                     "TRAINED-FLEX-GRID","TRAINED-IGNORE-MIRROR","TRAINED-IGNORE-POLARITY",
                     "TRAINED-MODEL","USAGE"]),
    ("DECODER",     ["AI-ENHANCER","ALLOW-ALL-MST-RESULTS","CENTERING-WINDOW",
                     "DISPLAY-TARGET","MAC-SCAN-TIMEOUT","NO-READ-FEEDBACK",
                     "REREAD-BASED-ON-SYMBOLOGY","REREAD-MODE","REREAD-NEVER2X",
                     "REREAD-NOT-LAST-N","REREAD-TIME","ROI","TARGET-DECODING",
                     "TIMEOUT","USE-CENTERING"]),
    ("DETECT",      ["MAX-SPEED","MIN-SPEED","PHOTO-EYE-DISTANCE"]),
    ("DETECTION",   ["ACTIVE","ENABLE","SENSITIVITY"]),
    ("DEVICE",      ["DESCRIPTION","FEATURE-KEYS","FIRMWARE-VER","LENS-SERIAL",
                     "MAC-ADDRESS","NAME","SERIAL-NUMBER","TIMEZONE","TYPE","UPTIME"]),
    ("DOTCODE",     ["TRAINED-MODEL","USAGE"]),
    ("ETH",         ["DHCP","GATEWAY","IP-ADDRESS","MAC-ADDRESS","MTU","SUBNET-MASK",
                     "VLAN-ENABLE","VLAN-ID"]),
    ("EVENT",       ["ENABLED","PORT","RESULT","RESULT-ENCODING"]),
    ("FILTER",      ["ENABLE","FILTER1","FILTER2","FILTER3","FILTER4","FILTER5",
                     "FILTER6","FILTER7","FILTER8","MODE"]),
    ("FORMAT",      ["ENABLE","FORMAT"]),
    ("GENERAL",     ["IP-ADDR","IP-PORT"]),
    ("HDR",         ["MODE"]),
    ("I2O5",        ["CHKCHAR","CODESIZE","QZ-SIZE","VERIFICATION","XMTCHK"]),
    ("IMAGE",       ["FORMAT","QUALITY","SEND","SIZE","TRANSFER-FORMAT"]),
    ("IO",          ["DEBOUNCE-TIME","INPUT-LOGIC","LINE-STATE","OUTPUT-LOGIC",
                     "OUTPUT-PULSE-DURATION","OUTPUT-STATE","TREND-ANALYSIS-FAIL",
                     "TREND-ANALYSIS-PASS","TREND-ANALYSIS-WARN"]),
    ("LIVEIMG",     ["MODE","SEND"]),
    ("MATCH",       ["MATCH-STRING","MASTER-DATABASE","MODE","STRING-COUNT"]),
    ("MAXCODE",     ["TRAINED-MODEL"]),
    ("MICRO-QR",    ["TRAINED-MODEL"]),
    ("MST",         ["DMST-LISTEN-PORT","DMST-SEND-HOST","ENABLE","RESULT-ALWAYSSEND","RESULT-DETAIL"]),
    ("NTP",         ["ENABLE","SERVER1","SERVER2"]),
    ("OCR",         ["CHARSET","ENABLE","FONT","LINESCAN-MODE","MODE","NOREAD-THRESHOLD",
                     "REGION","RESULT-DISPLAY","TRAINED-MODEL"]),
    ("PDF417",      ["CODESIZE","MACRO","MICRO","TRAINED-MODEL","VERIFICATION"]),
    ("POWERUP",     ["LINE-STATE","TRIGGER"]),
    ("QR",          ["CODESIZE","MODEL","TRAINED-MODEL","USAGE"]),
    ("RESULT",      ["DETAIL","DISPLAY","EXTERNAL","MULTICODE-MODE","MULTICODE-TIMEOUT",
                     "NOREAD-ENABLED","PARTIALLY-DECODED","SEND","SEPARATOR","SORT","STORE"]),
    ("STATISTICS",  ["CLEAR","ENABLED","FAIL-COUNT","GOOD-COUNT","MATCH","NOREAD-COUNT",
                     "PASS-COUNT","TOTAL-COUNT","TREND-WINDOW","WARN-COUNT"]),
    ("SYMBOL",      ["C128","C39","C93","CODABAR","DATABAR","DATAMATRIX","DOTCODE",
                     "I2O5","PDF417","PHARMACODE","QR","UPC-EAN"]),
    ("TRAIN",       ["AUTO-DISABLE","MAX-CODES","MODE","MULTICODE-ORDERING","PITCH-ANGLE",
                     "POLARITY","POSITION-TOLERANCE","TRAINED"]),
    ("TRIGGER",     ["BURST-LENGTH","DELAY","ENABLED","END-FRAME","GATE-DELAY",
                     "GATE-ENABLED","GATE-LENGTH","IN-QUEUE","MAX-ON-TIME","MIN-INTERVAL",
                     "OVERLAP","START-FRAME","TIMEOUT","TYPE"]),
    ("TRUCHECK",    ["APERTURE","APERTURE-SIZE","APPLICATION-CUSTOM-MAXIMUM-X-DIM",
                     "APPLICATION-CUSTOM-MINIMUM-X-DIM","APPLICATION-CUSTOM-PASS-GRADE",
                     "APPLICATION-GS1-TABLE","APPLICATION-STANDARD","AUTO-BATCH",
                     "BATCH-NUMBER","CALIBRATION-DATE","COMPANY-NAME","CUSTOM-NOTE",
                     "DECODE-GRADE","DOT-PEEN","FPD","FPD-GRADE","GNU","GNU-GRADE",
                     "GRADING-STANDARD","ISO15415-VERSION","ISO15416-VERSION",
                     "ISO29158-VERSION","METRIC-UNITS","MOD-GRADE","OPERATOR-NAME",
                     "REPORT-SECTION","REPORT-SECTION-ENABLED","RM-GRADE","SC",
                     "SC-GRADE","UEC","UEC-GRADE"]),
    ("TUNE",        ["EXCLUDE-AMBIENT-RESULTS","STATUS","TRAIN-CODE","TUNE-FILTERS",
                     "TUNE-LIGHT-BANKS","TUNE-LIGHT-COLORS","TUNE-LIGHT-COMBINATIONS"]),
    ("UPC-EAN",     ["SUPPLEMENT","SUPPLEMENT-REQUIRED","VERIFICATION"]),
    ("VERIFICATION",["ENABLE"]),
    ("WEBHMI",      ["ACTIONS","ENABLE","LANGUAGE","MATCH-STRING","SETTINGS","VERIFICATION"]),
    ("WEBUI",       ["ENABLE"]),
]

# ---------------------------------------------------------------

def connect():
    s = socket.socket()
    s.settimeout(4)
    s.connect((IP, PORT))
    time.sleep(0.2)
    s.sendall(b'||>SET COM.DMCC-RESPONSE 2\r\n')
    time.sleep(0.4)
    return s

def dmcc(s, cmd):
    s.sendall(('||>GET ' + cmd + '\r\n').encode())
    time.sleep(0.35)
    data = b''
    deadline = time.time() + 3
    while time.time() < deadline:
        try:
            chunk = s.recv(4096)
            if chunk:
                data += chunk
            else:
                break
        except socket.timeout:
            break
    if not data:
        return '(no response)'
    raw = data.decode('ascii', errors='replace')
    # Strip the wire header lines (||:::N[0]) — keep the value payload
    lines = raw.splitlines()
    value_lines = [l for l in lines if not l.startswith('||:::')]
    return '\n'.join(value_lines).strip() if value_lines else '(no response)'

# ---------------------------------------------------------------

ts  = datetime.datetime.now()
out = f"DM475V-DPM  866D76  {IP}  fw:?  dump:{ts.strftime('%Y-%m-%d %H:%M:%S')}\n"
out += "# " + "-"*70 + "\n\n"

print(f"Connecting to {IP}:{PORT} …")
s = connect()
print("Connected. Querying parameters …\n")

total = sum(len(v) for _, v in PARAMS)
done  = 0

for group, keys in PARAMS:
    out += f"# -- {group} {'-' * (55 - len(group))}\n"
    for key in keys:
        param = f"{group}.{key}"
        val   = dmcc(s, param)
        line  = f"{param:<52} = {val}"
        out  += line + "\n"
        print(line)
        done += 1
        if done % 20 == 0:
            print(f"  … {done}/{total} done")
    out += "\n"

s.close()

fname = f"dmcc_dump_{ts.strftime('%Y%m%d_%H%M%S')}.txt"
with open(fname, 'w', encoding='utf-8') as f:
    f.write(out)

print(f"\nDone — {done} parameters written to {fname}")
