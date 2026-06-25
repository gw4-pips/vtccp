# Wireshark Protocol Analysis — DM475V ↔ DMST Full Session Capture

**Capture**: "Wireshark - 475V Quit, Open, Connect, Verify"
**Archived**: `vtccp/architecture/gui-reference/wireshark-dmst-full-capture.txt` (7200 lines)
**Session date**: 2026-05-25 (DM475-63530E-PIPS-Verif-Lab, fw 6.1.16_sr4)
**Capture type**: Single TCP stream — Follow TCP Stream export from Wireshark
**Analyzed**: 2026-05-25

---

## 1. What this capture is — and is not

This file is the output of Wireshark's **"Follow TCP Stream"** on **one specific TCP connection**
between the DM475V device and DMST. It shows the DMST-native HTTP result-push channel only.
The DMCC command/response channel (port 44444) is a **separate TCP connection** that is NOT
in this capture.

**Two distinct TCP channels exist simultaneously during a DMST session:**

| Channel | Port | Direction | Protocol | Purpose |
|---|---|---|---|---|
| DMCC command channel | **44444** | DMST → device (request) / device → DMST (response) | Raw DMCC XML over TCP | GET/SET configuration, TRIGGER, GET SYMBOL.RESULT |
| HTTP result-push channel | **44444** (confirmed 2026-05-25) | DMST → device (GET subscribe) / device → DMST (PUT pushes, same TCP connection) | HTTP/1.1 over TCP | Result delivery, periodic status, config sync |

> **Both channels share port 44444.** The device distinguishes them by connection intent:
> a DMCC session sends raw XML; an HTTP session opens a TCP connection to port 44444
> and sends `GET /events?enable HTTP/1.1`. The device responds `204 No Content` and
> then uses the **same Keep-Alive TCP connection** to push `PUT /status.xml`,
> `PUT /codes.xml`, and `PUT /pcm_report.html` back to DMST.
> Source: Wireshark packet 46 — `Dst Port: 44444`, capture filter `host 10.10.10.7 and port 44444`.

This capture covers only the **HTTP result-push channel**.

---

## 2. HTTP result-push protocol — full architecture

### 2.1 Subscription handshake

DMST initiates the session by subscribing to the device's event stream:

```
→ GET /events?enable HTTP/1.1
  Date: Mon, 25 May 2026 13:10:12 GMT
  X-Peer: 42715336

← HTTP/1.1 204 No Content
  Server: DM475/6.1.16 (DeviceID=50)
  Content-Length: 0
```

After the 204 response, the device begins PUTting events to DMST's embedded HTTP server.
The device acts as HTTP CLIENT for all subsequent traffic.

### 2.2 Periodic telemetry — PUT /status.xml

**Frequency**: every ~1 second  
**Size**: ~4,623–4,628 bytes  
**User-Agent**: `DM475/6.1.16 (DeviceID=50)` (device is the sender)

Root element: `<status version="3">`

Contains:
- `<read_stats>`: good_reads, bad_reads, passed/failed validations, trigger counters, buffer stats, decoded_symbols
- `<monitored_values>`: image timing (request→receive ~27.9ms, time-between-images ~333ms), acquisition length, gap
- `<monitored_counters>`: image data buffer counter (~38/hr rate), MST network stats

### 2.3 Configuration sync — PUT /vs.cfg

**Frequency**: irregular, observed once per session in this capture  
**Size**: 288–400 bytes  
**Content-Encoding: aes** — body is AES-encrypted binary, **not readable**

This carries device configuration state. The encryption key is not known and the content
is not useful to VTCCP.

### 2.4 Result push — PUT /codes.xml

**Frequency**: once per scan trigger  
**Size**: ~9,415 bytes (monitoring scan, no JPEG) or ~202,249 bytes (full verification scan, includes JPEG image)  
**Content-Type**: `plain/xml` (mislabeled — it is XML)

This is the primary result delivery mechanism. Full structure documented in §3 below.

### 2.5 HTML verification report — PUT /pcm_report.html

**Frequency**: once per SUCCESSFUL verification scan (not on monitoring scans)  
**Size**: 131,915–202,249 bytes  
**Content-Type**: `text/xml` (mislabeled — it is HTML)  
**Timing**: sent IMMEDIATELY BEFORE the corresponding codes.xml PUT

The HTML body is a self-contained report document (no external dependencies). Content:
- Full CSS with grade color table (see §4)
- Cognex logo embedded as base64 PNG data URI `<img>`
- Verification report tables (same data shown in DMST TruCheck window)
- Modulation value grid (colored squares, per-module grading visualization)

**Filesystem relationship**: DMST also saves this same HTML content to disk at:
```
{Documents}\{DeviceName}\CodeQuality\{timestamp}.html
```
This is the file that `DmstHtmlScraper.ParseHtml()` reads. The file on disk and the
HTTP body are the same document.

---

## 3. codes.xml — full structure

```xml
<result id="4582" image_id="-1" version="3" origin="common">

  <general>
    <status>GOOD READ</status>
    <result_source>DM475-63530E-PIPS-Verif-Lab</result_source>
    <full_string encoding="base64">PD94bW...</full_string>   <!-- base64-encoded push XML -->
    <trigger_index>55</trigger_index>
    <trigger_time>289</trigger_time>
    <decode_time>252</decode_time>
    <symbology>Data Matrix</symbology>
    <generator>DM475-63530E-PIPS-Verif-Lab</generator>
    <bca_result>No Stats</bca_result>
    <bca_parcel>No Stats</bca_parcel>
    <bca_face>No Stats</bca_face>
    <bca_position>{-1, -1, -1}</bca_position>
    <read_setup>0</read_setup>
    <module_size>15.86</module_size>
  </general>

  <statistics>
    <!-- same as status.xml read_stats block -->
    <read_stats>
      <good_reads>28</good_reads>
      <bad_reads>1</bad_reads>
      ...
    </read_stats>
  </statistics>

  <trucheck_verificaiton_result>   <!-- NOTE: "verificaiton" is misspelled in firmware -->
    ...                            <!-- see §4 for full contents -->
  </trucheck_verificaiton_result>

</result>
```

### 3.1 The `origin` field — critical discriminator

| Value | Meaning | codes.xml size | pcm_report.html sent? |
|---|---|---|---|
| `"monitor"` | Continuous-mode monitoring scan; no full verification data | ~9,415 bytes | No |
| `"common"` | Full triggered verification scan with TruCheck data + JPEG | ~202,249 bytes | Yes (sent first) |

Monitor-mode scans (`origin="monitor"`) also have a `<trucheck_verificaiton_result>` block,
but it contains only `<status>valid</status>`, `<CalibrationDate>`, `<OpticalVariant>`, and
`<SymbolData><DecodedData>NO DECODE</DecodedData>` with a single DECODE=F parameter.
The full quality parameter table and general characteristics only appear on `origin="common"`.

### 3.2 The `full_string` field — push script output

The `<full_string encoding="base64">` content is the complete push script output, base64-encoded.
Decoded, it is a `<DMCCResponse><DMSymVerResponse>` document — the same XML that VTCCP's
push parser already processes via ResultsReceived / DmstListener.

This confirms: the push script output reaches VTCCP through the DataMan Network Client
mechanism (raw XML over TCP), and also travels through the DMST-native HTTP path wrapped
in base64. The two paths carry identical push XML content.

**Decoded push XML from successful scan (trigger_index=55, Grade D, 2026-05-25T09:10:19):**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<DMCCResponse>
<DMSymVerResponse>
  <DateTime>2026-05-25T09:10:19</DateTime>
  <SymbologyName>Data Matrix</SymbologyName>
  <DecodedData>[)&amp;gt;0618VD895361P8902AS3122A02965</DecodedData>
  <SymbologyId>]d1</SymbologyId>
  <SymbolQuality>41</SymbolQuality>
  <SymbolAngle>7</SymbolAngle>
  <ModuleSizePx>15.857597351074219</ModuleSizePx>
  <PushScriptDiag>v1.33 q=r.trucheck m=found</PushScriptDiag>
  <Source>DM475-63530E-PIPS-Verif-Lab</Source>
  <FormalGrade>1/D</FormalGrade>
  <OverallGrade>D</OverallGrade>
  <OverallGradeNumeric>1</OverallGradeNumeric>
  <GradingStandard>ISO 15415:2011</GradingStandard>
  <ApplicationStandard>Custom</ApplicationStandard>
  <ApplicationPass>Fail (Quality)</ApplicationPass>
  <ApplicationPassReason>Quality</ApplicationPassReason>
  <ApertureRef>16</ApertureRef>
  <Wavelength>660</Wavelength>
  <Lighting>45Q</Lighting>
  <Standard>ISO 15415:2011</Standard>
  <UECPercent>41.7</UECPercent>
  <UECGrade>C</UECGrade>
  <SCPercent>78.3</SCPercent>
  <SCRlRd>83/4</SCRlRd>
  <SCGrade>A</SCGrade>
  <MinReflectance>4</MinReflectance>
  <MODGrade>A</MODGrade>
  <RMGrade>C</RMGrade>
  <ANUPercent>11.1</ANUPercent>
  <ANUGrade>D</ANUGrade>
  <GNUPercent>7.5</GNUPercent>
  <GNUGrade>A</GNUGrade>
  <FPDValue>4</FPDValue>
  <FPDGrade>A</FPDGrade>
  <DecodeGrade>A</DecodeGrade>
  <MatrixSize>16x36</MatrixSize>
  <HorizontalBWG>11</HorizontalBWG>
  <VerticalBWG>11</VerticalBWG>
  <EncodedCharacters>33</EncodedCharacters>      <!-- WRONG: firmware push gives 33; correct=38 -->
  <TotalCodewords>56</TotalCodewords>
  <DataCodewords></DataCodewords>                <!-- EMPTY: q.symbols null on fw 6.1.16_sr4 -->
  <ErrorCorrectionBudget></ErrorCorrectionBudget><!-- EMPTY: same -->
  <ErrorsCorrected>7</ErrorsCorrected>
  <ErrorCapacityUsed>14</ErrorCapacityUsed>
  <ErrorCorrectionType>ECC200</ErrorCorrectionType>
  <NominalXDim>20.3 mil</NominalXDim>
  <ImagePolarity></ImagePolarity>               <!-- EMPTY: not accessible via push XML — resolved via DmstHtmlScraper HTML report -->
  <ContrastUniformity>74</ContrastUniformity>
  <MRD>67</MRD>
  <ContrastUniformityRow>12</ContrastUniformityRow>
  <ContrastUniformityCol>17</ContrastUniformityCol>
  <OpticsSource>LiveScan</OpticsSource>
  <DDGrade>X</DDGrade>
  <AverageGrade>A</AverageGrade>
  <AverageGradeNumeric>4</AverageGradeNumeric>
  <CustomNote></CustomNote>
  <LLSGrade>A</LLSGrade>
  <BLSGrade>A</BLSGrade>
  <LQZGrade>A</LQZGrade>
  <BQZGrade>A</BQZGrade>
  <TQZGrade>X</TQZGrade>     <!-- X = not applicable for this DM symbol orientation -->
  <RQZGrade>X</RQZGrade>
  <TTRPercent></TTRPercent>
  <TTRGrade>X</TTRGrade>
  <RTRPercent></RTRPercent>
  <RTRGrade>X</RTRGrade>
  <TCTGrade>X</TCTGrade>
  <RCTGrade>X</RCTGrade>
  <!-- New grade fields (empty on 16×36 DM — likely for other DM variants or DMV-8072V) -->
  <ULQZGrade></ULQZGrade>
  <URQZGrade></URQZGrade>
  <RUQZGrade></RUQZGrade>
  <RLQZGrade></RLQZGrade>
  <LLQZGrade></LLQZGrade>
  <LRQZGrade></LRQZGrade>
  <HClockTrackGrade></HClockTrackGrade>
  <VClockTrackGrade></VClockTrackGrade>
  <ULQTTRPercent></ULQTTRPercent>
  <ULQTTRGrade></ULQTTRGrade>
  <URQTTRPercent></URQTTRPercent>
  <URQTTRGrade></URQTTRGrade>
  <LLQTTRPercent></LLQTTRPercent>
  <LLQTTRGrade></LLQTTRGrade>
  <LRQTTRPercent></LRQTTRPercent>
  <LRQTTRGrade></LRQTTRGrade>
  <ULQRTRPercent></ULQRTRPercent>
  <ULQRTRGrade></ULQRTRGrade>
  <URQRTRPercent></URQRTRPercent>
  <URQRTRGrade></URQRTRGrade>
</DMSymVerResponse>
</DMCCResponse>
```

**New push XML fields confirmed from this capture (not previously inventoried):**

| Field | Value (DM 16×36 live, Grade D) | Notes |
|---|---|---|
| `<TQZGrade>` | `X` | Top Quiet Zone — X on 16×36 (orientation-dependent) |
| `<RQZGrade>` | `X` | Right Quiet Zone — X on 16×36 |
| `<ULQZGrade>` | `` (empty) | Upper-Left QZ — empty on standard DM |
| `<URQZGrade>` | `` | Upper-Right QZ |
| `<RUQZGrade>` | `` | Right-Upper QZ |
| `<RLQZGrade>` | `` | Right-Lower QZ |
| `<LLQZGrade>` | `` | Left-Lower QZ |
| `<LRQZGrade>` | `` | Left-Right QZ |
| `<HClockTrackGrade>` | `` | Horizontal clock track |
| `<VClockTrackGrade>` | `` | Vertical clock track |
| `<ULQTTRPercent/Grade>` | `` | Upper-Left QZ Top Transition Ratio |
| `<URQTTRPercent/Grade>` | `` | Upper-Right QZ Top Transition Ratio |
| `<LLQTTRPercent/Grade>` | `` | Lower-Left QZ Top Transition Ratio |
| `<LRQTTRPercent/Grade>` | `` | Lower-Right QZ Top Transition Ratio |
| `<ULQRTRPercent/Grade>` | `` | Upper-Left QZ Right Transition Ratio |
| `<URQRTRPercent/Grade>` | `` | Upper-Right QZ Right Transition Ratio |

These fields are likely populated for DMV-8072V symbols (which have four finder elements) and
for DM variants with different quiet zone topology. All empty on the standard 16×36 DM symbol.

---

## 4. trucheck_verificaiton_result block — complete inventory

**IMPORTANT**: The XML tag name is `<trucheck_verificaiton_result>` — note the misspelling
("verificaiton" with transposed "ai"). This is a firmware literal; parse it exactly.

The block appears in EVERY codes.xml PUT. Monitor-mode scans have a minimal version.
Full data only appears on `origin="common"` (triggered verification) scans.

### 4.1 Top-level fields

```xml
<trucheck_verificaiton_result>
  <status>valid</status>
  <CalibrationDate>5/20/2026 1:14:58 AM</CalibrationDate>
  <OpticalVariant>DM475V</OpticalVariant>   <!-- EXACT device model string -->
  <SymbolData>
    <CalibrationState>0</CalibrationState>
    <SymbologyType>DataMatrix</SymbologyType>
    <DecodedData>[)&gt;&lt;RS&gt;06...</DecodedData>
    <Base64Data>Wyk+HjA2HTE4VkQ4OTUzNh0x...</Base64Data>
    <VerificaitonTime>183</VerificaitonTime>   <!-- misspelled in firmware -->
    <PreDecodeTime>0</PreDecodeTime>
    <BlurTime>64</BlurTime>
    <ThreshTime>13</ThreshTime>
    <StickTime>0</StickTime>
    <LineSearchTime>13</LineSearchTime>
    <CanidateEvaluationTime>30</CanidateEvaluationTime>  <!-- "Canidate" misspelled in firmware -->
    <PostDecodeTime>51</PostDecodeTime>
    <ResultTime>0</ResultTime>
    <!-- ... report sections follow ... -->
  </SymbolData>
</trucheck_verificaiton_result>
```

**`<OpticalVariant>DM475V</OpticalVariant>`**: This is the exact model string as the device
reports it — "DM475V", not "DM470" (which is the family designation shown in the DMST device
list). VTCCP reads `DEVICE.TYPE` via DMCC on connect; confirm whether that returns "DM475V" or
"DM470". If DMCC returns "DM470", the trucheck XML OpticalVariant is the authoritative source.

**`<CalibrationState>0</CalibrationState>`**: Meaning of 0 not yet confirmed. Observed value
on this device. Correlation with `FieldCalibrated` in push XML (always false) is unknown.

### 4.2 ReportSection — Verification Grades (sectionType="GradingInfo")

```xml
<ReportSection sectionType="GradingInfo" sectionTitle="Verification Grades">
  <GradeInfo>
    <Standard>ISO 15415:2011</Standard>
    <Grade>1.0</Grade>            <!-- numeric grade -->
    <ValueGrade>D</ValueGrade>    <!-- letter grade -->
    <Aperture>16</Aperture>
    <Wavelength>660</Wavelength>
    <Lighting>45Q</Lighting>
    <FormalGrade>1.0/16/660/45Q</FormalGrade>  <!-- COMPLETE ISO formal notation -->
  </GradeInfo>
  <GradeInfo>
    <Standard>Custom</Standard>
    <Grade>Fail (Quality)</Grade>
  </GradeInfo>
</ReportSection>
```

**`<FormalGrade>1.0/16/660/45Q</FormalGrade>`**: This is the ISO 15415 formal grade notation
in full form: `numericGrade / aperture / wavelength / lighting`. The push XML `<FormalGrade>`
field contains `1/D` (grade/letter) — a different abbreviated format. The trucheck XML has
the complete notation needed for formal conformance documentation.

### 4.3 ReportSection — Grade History (sectionType="GradeHistory")

```xml
<ReportSection sectionType="GradeHistory" sectionTitle="Grade History">
  <FailWindow>1</FailWindow>
  <MaxFailCount>1</MaxFailCount>
  <CurrentFailCount>1</CurrentFailCount>
  <VerificationOverallPass>0</VerificationOverallPass>
  <Grade>0.0</Grade>
  <Grade>0.0</Grade>
  <Grade>1.0</Grade>
  <!-- ... 27 grade entries total ... -->
</ReportSection>
```

### 4.4 ReportSection — ISO15415 Quality Parameters (sectionType="Table")

Full parameter list confirmed from this capture. Each `<Parameter>` has `<Number>`, `<Name>`,
`<Grade>`, optional `<Value>` and `<Data>`, `<Check>`.

| Number | Name | Grade (this scan) | Value | Check |
|---|---|---|---|---|
| 1 | Unused Error Correction (UEC) | 2.0 | 41.7% | PASS |
| 2 | Symbol Contrast (SC) | 4.0 | 78% (Data: Rl/Rd 83/4) | PASS |
| 3a | Modulation (MOD) | 4.0 | — | PASS |
| 3b | Reflectance Margin (RM) | 2.0 | — | PASS |
| 4 | Axial Nonuniformity (ANU) | 1.0 | 11.1% | FAIL |
| 5 | Grid Nonuniformity (GNU) | 4.0 | 7.5% | PASS |
| 6 | Fixed Pattern Damage (FPD) | 4.0 | 4.0 | PASS |
| 7 | Left 'L' Side (LLS) | 4.0 | — | PASS |
| 8 | Bottom 'L' Side (BLS) | 4.0 | — | PASS |
| 9 | Left Quiet Zone (LQZ) | 4.0 | — | PASS |
| 10 | Bottom Quiet Zone (BQZ) | 4.0 | — | PASS |
| 11a | Upper Left Quiet Zone (ULQZ) | 4.0 | — | PASS |
| 11b | Upper Right Quiet Zone (URQZ) | 4.0 | — | PASS |
| 12a | Right Upper Quiet Zone (RUQZ) | 4.0 | — | PASS |
| 12b | Right Lower Quiet Zone (RLQZ) | 4.0 | — | PASS |
| 13a | Left Top Transition Ratio (LQTTR) | 4.0 | 0% | PASS |
| 13b | Right Top Transition Ratio (RQTTR) | 4.0 | 0% | PASS |
| 14a | Left Right Transition Ratio (LQRTR) | 4.0 | 0% | PASS |
| 14b | Right Right Transition Ratio (RQRTR) | 4.0 | 0% | PASS |
| 15a | Left Top Clock Track (LQTCT) | 4.0 | — | PASS |
| 15b | Right Top Clock Track (RQTCT) | 4.0 | — | PASS |
| 16a | Left Right Clock Track (LQRCT) | 4.0 | — | PASS |
| 16b | Right Right Clock Track (RQRCT) | 4.0 | — | PASS |
| 17 | Average Grade (AG) | 4.0 | 4.0 | PASS |
| 18 | DECODE | 4.0 | — | PASS |

### 4.5 ReportSection — General Characteristics (sectionType="Table") — CRITICAL

**This section resolves previously-unresolvable fields.** All values confirmed from
this capture (DM 16×36 GS1 label, live scan, 2026-05-25):

| Name | Data (this scan) | Notes |
|---|---|---|
| Matrix Size | `16x36 (Data: 14x34)` | Includes inner data region size |
| Horizontal BWG | `11%` | ✓ matches push XML |
| Vertical BWG | `11%` | ✓ matches push XML |
| **Encoded characters** | **`38`** | ★ **CORRECT VALUE** — push XML gives 33 (wrong) |
| **Total Codewords** | **`56`** | ✓ matches push XML |
| **Data Codewords** | **`32`** | ★ **CORRECT VALUE** — push XML gives empty |
| **Error Correction Budget** | **`24`** | ★ **CORRECT VALUE** — push XML gives empty |
| Errors Corrected | `7` | ✓ matches push XML |
| Error Capacity Used | `14` | ✓ matches push XML |
| Error Correction Type | `ECC 200` | Note space (push XML: `ECC200` no space) |
| **Image** | **`Black on white`** | ★ **ImagePolarity resolved** — push XML gives empty |
| Nominal X Dim | `20.3 mil` | ✓ matches push XML |
| Pixels per Module | `15.96` | New field — not in push XML |
| Contrast Uniformity | `74 at module(12,17)` | Location annotation — push XML gives just `74` |
| MRD | `67% (73% - 6%)` | Expanded form — push XML gives just `67` |

**Field resolution status — all symbologies:**

| Field | Applies to | Push XML status | trucheck XML | HTML report (DmstHtmlScraper) |
|---|---|---|---|---|
| `EncodedCharacters` | All | Wrong (33 vs correct 38) | ✓ CORRECT: 38 | ✓ "Encoded characters: 38" |
| `DataCodewords` | All | Empty | ✓ CORRECT: 32 | ✓ "Data Codewords: 32" |
| `ErrorCorrectionBudget` | All | Empty | ✓ CORRECT: 24 | ✓ "Error Correction Budget: 24" |
| `ImagePolarity` | All | Empty | ✓ "Black on white" | ✓ "Image: Black on white" — **RESOLVED** in ParseHtml() |
| `ECLevel` | **QR only** — DM ECC200 has no selectable level; field is inapplicable to DM | Not accessible via push XML | ✗ Not present (FIB={grade,numericGrade} only) | ✓ "Error Correction Level"="M" confirmed in QR HTML — ParseHtml() extension pending |
| `DataMaskPattern` | **QR only** — DM has no data masking; field is inapplicable to DM | Not accessible via push XML | ✗ Not present | ✓ "Data Mask Pattern"="2" confirmed in QR HTML — ParseHtml() extension pending |
| `ECI` | QR and others | Not accessible via push XML | ✗ Not present | ✓ "ECI"="000003" confirmed in QR HTML — ParseHtml() extension pending |

**Confirmed**: The trucheck XML and the HTML report are parallel representations of the same
data. `DmstHtmlScraper.ParseHtml()` reads the same information from the filesystem-saved HTML.
All four resolvable fields above should be extractable from the HTML via the scraper.

### 4.6 Remaining report sections

- **ASCII Values** (`sectionType="TabularData"`): per-character decimal ASCII values. 38 entries for this GS1 symbol. Not currently captured in VTCCP — informational only.
- **Codewords** (`sectionType="TabularData"`): per-codeword decimal values; `*` prefix marks error-corrected codewords. 56 entries (32 data + 24 ECC). `*=Fixed by Error Correction` notation.
- **Encodation Analysis** (`sectionType="Table"`): per-codeword columns Codeword/Mode/Result. Mode examples: `ASCII`. Result shows decoded character or control code.

---

## 5. pcm_report.html — grade color CSS table (confirmed from capture)

The CSS grade-to-color mapping used by DMST's HTML report:

| CSS class | Color | Hex | Numeric grade range |
|---|---|---|---|
| `.g34grade`–`.g45grade` | Green | `#13D232` | ≥ 3.4 (letter A) |
| `.g25grade`–`.g33grade` | Yellow-Green | `#89E232` | 2.5–3.3 (letter B) |
| `.g15grade`–`.g24grade` | Yellow | `#FFFF00` | 1.5–2.4 (letter C) |
| `.g5grade`–`.g14grade` | Orange-Red | `#FF4060` | 0.5–1.4 (letter D) |
| `.g0grade`–`.g4grade` | Red | `#FF3232` | 0.0–0.4 (letter F) |

Note: CSS classes are per-tenth (`.g10grade` = 1.0, `.g25grade` = 2.5, etc.).
Modulation value squares in the HTML use `background-color` inline styles that
differ from this table (e.g., `#5AE070` = passing green for modulation grid).

---

## 6. VTCCP access path analysis

### 6.1 Current VTCCP push architecture (Network Client mode)

VTCCP's `DmstListener` uses the DataMan **Network Client** feature:
- Device configured with VTCCP's host IP and port as Network Client destination
- After each scan, device TCP-connects to VTCCP and sends raw push script XML
- No HTTP wrapper — the push XML arrives directly as `<DMCCResponse>...`
- `DmstListener.HandleClientAsync()` receives this, detects it starts with `<`, parses as XML
- The `trucheck_verificaiton_result` block does NOT travel via this path — it is only in the DMST-native HTTP PUT /codes.xml

### 6.2 DmstHtmlScraper — correct and current access path

`DmstHtmlScraper` watches the filesystem for HTML files written by DMST. DMST saves the
same HTML content it receives via `PUT /pcm_report.html` to:
```
{Documents}\{DeviceName}\CodeQuality\{timestamp}.html
```
This is the approved supplemental-data path for VTCCP. It captures:
- EncodedCharacters (correct value)
- DataCodewords (correct value)
- ErrorCorrectionBudget (correct value)
- ImagePolarity ("Black on white" / "White on black")
- Any ECLevel/DataMaskPattern/ECI if present in the HTML (to be verified on QR scan)

**Prerequisite**: DMST must be running alongside VTCCP, with "Preferred Quality Report File
Extension" set to ".html" in DMST Options → Data Logging → Reporting.

### 6.3 Alternative: HTTP subscriber mode (future architecture)

VTCCP could implement an HTTP server that subscribes to the device event stream by
sending `GET /events?enable`. It would then receive `PUT /codes.xml` directly, including
the full `trucheck_verificaiton_result` XML. This would eliminate the DMST filesystem
dependency for supplemental data.

**Port confirmed 2026-05-25**: The device HTTP event server is on **port 44444** — the same
port as DMCC. A separate TCP connection to port 44444 beginning with `GET /events?enable`
initiates the HTTP push channel. No separate port configuration required.

**Implementation sketch** (when this is eventually built):
1. Open a new `TcpClient` to `_cfg.Host:44444` (separate from the DMCC SDK connection)
2. Send:
   ```
   GET /events?enable HTTP/1.1\r\n
   Date: {RFC1123 timestamp}\r\n
   X-Peer: {session token}\r\n
   \r\n
   ```
3. Read `HTTP/1.1 204 No Content` — then loop reading HTTP PUT requests on the same stream
4. Parse `PUT /codes.xml` body: extract `<full_string encoding="base64">` from `<general>` block
5. Parse `PUT /pcm_report.html` body: same HTML format as DmstHtmlScraper expects

**Decision for now**: Not implementing. DmstHtmlScraper (§6.2) is the current approved path.
Log this as a future architecture option — fully unblocked.

---

## 8. DPM device capture — live mode toggle analysis (2026-06-24)

**Capture**: User-initiated, DM475V-DPM at 10.10.10.4  
**Display filter**: `ip.addr == 10.10.10.4` (no port filter)  
**Capture date**: 2026-06-24  
**Condition**: Verifier idle at capture start; live mode toggled ON → OFF → ON → OFF twice  
**Archived**: `attached_assets/Pasted-11-0-426049100-…_1782347481563.txt` (2256 lines)

### 8.1 Live mode control endpoint — confirmed

DMST uses an HTTP REST call (not raw DMCC text) to toggle live/monitor mode:

| Action | Request | Response |
|---|---|---|
| Enable live mode | `GET /monitormode?enable=true HTTP/1.1` | `HTTP/1.1 204 No Content` |
| Disable live mode | `GET /monitormode?enable=false HTTP/1.1` | `HTTP/1.1 204 No Content` |

**Four complete toggle cycles observed** (all confirmed `204 No Content`):

| Packet | Time (s) | Event |
|---|---|---|
| 575 | 23.676 | `GET /monitormode?enable=true` (ON #1) |
| 590 | 23.683 | `204 No Content` |
| 1779 | 27.009 | `GET /monitormode?enable=false` (OFF #1) |
| 1787 | 27.024 | `204 No Content` |
| 1955 | 31.209 | `GET /monitormode?enable=true` (ON #2) |
| 1971 | 31.216 | `204 No Content` |
| 3022 | 34.305 | `GET /monitormode?enable=false` (OFF #2) |
| 3051 | 34.557 | `204 No Content` |

### 8.2 What is NOT present in the capture

| Expected / hypothesized | Actual |
|---|---|
| `MONITOR-MODE.ENABLE` DMCC raw text command | **Absent** — DMST uses HTTP REST, not raw DMCC |
| `VERIFICATION.ENABLE` command (any form) | **Absent** — no separate verification toggle at all |
| Any raw DMCC text command | **Absent** — zero raw DMCC traffic on port 44444 |
| `GET /events?enable` subscription handshake | Not visible (capture started after connection was established) |

**Conclusion on VERIFICATION.ENABLE**: TruCheck verification is always active during monitor
mode. There is no separate per-session toggle. The DMCC reference lists `VERIFICATION.ENABLE`
as applicable to DM8072V/DM370/DM390/DM470 — but the DM475V does not use it through this path.

### 8.3 Live image streaming — confirmed failure mode

Every `GET /svg_image.img` returns `HTTP/1.1 500 Internal Server Error` for non-DMST clients.
This is consistent across all four live-mode active periods, multiple polls per period:

```
GET /svg_image.img HTTP/1.1   →   HTTP/1.1 500 Internal Server Error
```

DMST's live view works through a different mechanism (likely AES-keyed channel or SDK
integration). Third-party access to the live video stream is not available via this endpoint.
This is consistent with the prior finding that `/svg_image.img` is AES-encrypted.

### 8.4 Device behavior on mode toggle

Immediately after each `GET /monitormode?enable=true`:
1. Device pushes `PUT /vs.cfg` (AES-encrypted config sync, ~490–620 bytes)
2. Device begins pushing `PUT /codes.xml` (~every 300ms, `origin="monitor"`)
3. Device continues `PUT /status.xml` (~every second)

Immediately after each `GET /monitormode?enable=false`:
1. Device pushes final `PUT /codes.xml` and `PUT /status.xml`
2. Device pushes `PUT /vs.cfg` (config sync, ~235–476 bytes)
3. Traffic returns to keep-alive only

### 8.5 Two-connection architecture — confirmed on DPM device

| Connection | Local port | Role |
|---|---|---|
| Subscription channel | 55653 | Long-lived; device pushes `PUT /status.xml`, `PUT /codes.xml` to DMST |
| Command channel | 55654 | DMST sends `GET /monitormode?…` and `GET /svg_image.img` here |

Both connect to port 44444 on the device. Identical architecture to the LBL device capture.

### 8.6 Implications for VTCCP

To control live mode on the DM475V independently of DMST, VTCCP needs:

```
GET /monitormode?enable=true HTTP/1.1\r\n
Host: 10.10.10.4:44444\r\n
\r\n
```

This is a plain HTTP GET on port 44444 — the same port as the DMCC SDK connection. The device
accepts it on any TCP connection that speaks HTTP (not raw DMCC XML). A separate `TcpClient`
should be used to avoid interfering with the SDK connection.

**No `MONITOR-MODE.ENABLE` DMCC SET command is required.** The HTTP endpoint is the confirmed path.

---

## 7. Confirmed facts from this capture

| Finding | Value |
|---|---|
| DMST-native result protocol | HTTP pub/sub over TCP on **port 44444** (same port as DMCC — confirmed from Wireshark packet 46 TCP header) |
| Device HTTP User-Agent | `DM475/6.1.16 (DeviceID=50)` |
| Subscription endpoint | `GET /events?enable` → `204 No Content` |
| Push endpoints | `PUT /status.xml`, `PUT /vs.cfg` (AES), `PUT /codes.xml`, `PUT /pcm_report.html` |
| codes.xml root element | `<result id="N" image_id="-1" version="3" origin="monitor\|common">` |
| Verification discriminator | `origin="common"` = full verification; `origin="monitor"` = monitoring scan |
| Push XML carrier | `<full_string encoding="base64">` in codes.xml `<general>` block |
| OpticalVariant | `DM475V` (exact model string — trucheck XML) |
| FormalGrade format (trucheck XML) | `1.0/16/660/45Q` (numeric/aperture/wavelength/lighting) |
| FormalGrade format (push XML) | `1/D` (abbreviated: grade/letter) |
| Misspelled firmware tags | `<trucheck_verificaiton_result>`, `<VerificaitonTime>`, `<CanidateEvaluationTime>` |
| EncodedCharacters (trucheck XML) | 38 (CORRECT — push XML gives wrong 33) |
| DataCodewords (trucheck XML) | 32 (CORRECT — push XML gives empty) |
| ErrorCorrectionBudget (trucheck XML) | 24 (CORRECT — push XML gives empty) |
| ImagePolarity (trucheck XML) | "Black on white" (push XML gives empty) |
| ECLevel (QR only), DataMaskPattern (QR only), ECI | Not in trucheck XML or push XML — accessible via DmstHtmlScraper QR HTML report (ParseHtml() extension pending). DM has no ECLevel or data masking — these fields are inapplicable to DM. |
| CalibrationState | 0 (meaning not yet confirmed) |
| pcm_report.html timing | Sent BEFORE the codes.xml for the same scan |
| HTML CSS grade color | Green ≥34, Yellow-Green 25-33, Yellow 15-24, Orange-Red 5-14, Red 0-4 |
| NominalXDim per scan | 20.3 mil (DM 16×36 live); 12.6 mil (QR 29×29 loaded) |
| Pixels per Module | 15.96 (this scan — only in trucheck XML, not in push XML) |
| Status.xml rate | ~1 per second, 4.6KB |
| Codewords section | Per-codeword decimal values, *=error-corrected |
| Encodation Analysis | Per-codeword Mode/Result table |

---

## 9. Full cold-start session capture — DPM device (2026-06-24)

**Capture**: DMST cold start → TC connect → Go Live → Trigger scan → Go Live → Cancel  
**Device**: DM475V-DPM at 10.10.10.4  
**Total packets captured**: 3,581 lines  
**Archived**: `attached_assets/Pasted-88-4-052788100-…_1782348465833.txt`

### 9.1 UDP discovery — pre-connection broadcast

Before DMST opens any TCP connection, the device broadcasts/unicasts on **UDP port 1069**:

| Time | Source | Destination | Notes |
|---|---|---|---|
| t=4.05s | 10.10.10.4 | 10.10.10.19 (unicast) | 181-byte payload |
| t=4.85s | 10.10.10.4 | 255.255.255.255 (broadcast) | 181-byte payload |
| t=7.34s | 10.10.10.4 | 10.10.10.19 (unicast) | 181-byte payload |

These precede the TCP connection by ~6 seconds. Likely device presence/heartbeat announcements
that DMST uses to detect and list available devices. Not needed for VTCCP (device IP is known).

### 9.2 TCP connection establishment — t=10.20s

Two TCP connections opened simultaneously to port 44444:

| Connection | Local port | Role | First request |
|---|---|---|---|
| A | **54767** | Event subscription channel | `GET /events?enable` → 204 |
| B | **54768** | Command/request channel | Unknown (Continuation) → 200 / 204 |

### 9.3 Initial handshake on command channel — unknown endpoints

Immediately after connection B (54768) is established, two `Continuation`-labeled HTTP requests
are sent and answered before `GET /events?enable` fires on connection A:

| Packet | Time | Request | Response |
|---|---|---|---|
| 505 | t=10.23s | `Continuation` (130 bytes total) on 54768 | `HTTP/1.1 200 OK` (195 bytes) — returns content |
| 508 | t=10.26s | `Continuation` (131 bytes total) on 54768 | `HTTP/1.1 204 No Content` |

Wireshark labels these "Continuation" because it cannot decode the HTTP method/URL from the
packet boundaries — the actual endpoint names are in the raw TCP payload. To determine the URLs,
use Wireshark → Right-click stream → **Follow TCP Stream** on connection 54768 and read the
first ~300 bytes. The first call returns `200 OK` (content), suggesting an auth or capability
query; the second returns `204 No Content`.

### 9.4 Subscription handshake — confirmed cold-start sequence

```
→ GET /events?enable HTTP/1.1  (conn A, t=10.28s)
← HTTP/1.1 204 No Content
```

Device immediately begins periodic `PUT /status.xml` pushes on this connection.

### 9.5 DMST initialization GET sequence — complete inventory

All on command channel (conn B, 54768):

| Request | Response | Notes |
|---|---|---|
| `GET /vs.cfg` | `200 OK` (~24KB) | First config fetch — AES body, unreadable |
| `GET /parameters.xml` | `200 OK` (large) | **NEW** — never seen before; likely full device parameter list |
| `GET /vs.cfg` | `200 OK` (~24KB) | Second config fetch |
| `GET /status.xml` | `200 OK` (~3.9KB) | Device status snapshot |
| `GET /device_info.xml` | `401 Unauthorized` | Auth required — DMST first attempt fails |
| `GET /device_info.xml` | `401 Unauthorized` | Retry — also fails; DMST cannot authenticate this endpoint |
| `GET /status.xml` | `200 OK` | Second status poll |

**`GET /parameters.xml`** is a newly-discovered endpoint. The large response (~63KB based on
TCP segment count) may be the complete device configuration parameter dump — equivalent to
running `GET ALL` over DMCC. Worth examining in detail.

**`GET /device_info.xml` → 401** is persistent. DMST either does not have credentials for
this endpoint or the DM475V does not support it. Not needed for VTCCP.

### 9.6 Go Live → Sleep — confirmed

```
→ GET /monitormode?enable=true HTTP/1.1   (t=16.64s)
← HTTP/1.1 204 No Content
```

Device immediately pushes: `PUT /status.xml` ×2, `PUT /vs.cfg` (588B), then begins monitor
scanning cycle (`PUT /codes.xml` + `PUT /status.xml` at ~300ms cadence).

### 9.7 Trigger command — endpoint unknown, two-packet pattern

When the user clicks **Go Live** from Sleep mode to trigger a verification scan, two
`Continuation`-labeled HTTP requests fire in rapid succession on the command channel:

| Packet | Time | Request | Response |
|---|---|---|---|
| 2035 | t=21.19s | `Continuation` (133 bytes) on 54768 | `HTTP/1.1 204 No Content` |
| 2038 | t=21.23s | `Continuation` (134 bytes) on 54768 | `HTTP/1.1 204 No Content` |

Both return `204 No Content` within ~2ms. The pattern matches other control commands
(`/monitormode?enable=*`). Same pattern observed in second trigger (pkts 4307, 4310 at t=35.16s).

**Scan result arrives ~550ms after trigger** as:
1. `PUT /vs.cfg` (235B)
2. `PUT /pcm_report.html` (full HTML verification report)
3. `PUT /codes.xml` (verification result + base64 push XML)
4. **`PUT /svg_image.img`** (scan image — see §9.8)
5. `PUT /status.xml` ×2

**To identify the trigger URL**: In Wireshark, right-click on pkt 2035 or 4307 →
**Follow TCP Stream** → read the raw ASCII — the HTTP GET line will be visible at the
start of the Continuation payload. Expected format: `GET /verify HTTP/1.1` or similar.

### 9.8 ★ PUT /svg_image.img — CRITICAL FINDING

**The device PUSHes the scan image to DMST.** It does NOT serve it via GET.

```
← PUT /svg_image.img HTTP/1.1   (device → DMST, on events channel 54767)
```

This appears in the scan result bundle immediately after `PUT /codes.xml`. The device sends
the verification image (and live-mode frames) by PUTting them on the **subscription channel**
(the same keep-alive TCP connection opened with `GET /events?enable`).

| Observed | Explanation |
|---|---|
| `GET /svg_image.img` → `HTTP/1.1 500 Internal Server Error` | DMST polling the wrong direction — device does NOT serve images via GET |
| `PUT /svg_image.img` (device → DMST) | Correct mechanism — device PUSHes image as HTTP PUT on events channel |

**For VTCCP live image capture**: Subscribe via `GET /events?enable`, then receive
`PUT /svg_image.img` events on the same TCP connection. The body is the image (format TBD —
likely JPEG). Parse the HTTP PUT body from the event stream.

### 9.9 Second Go Live (post-scan) and Cancel

```
→ GET /monitormode?enable=true HTTP/1.1   (t=25.39s) → 204 No Content
→ GET /monitormode?enable=false HTTP/1.1  (t=28.09s) → 204 No Content
→ GET /monitormode?enable=true HTTP/1.1   (t=31.29s) → 204 No Content
→ GET /monitormode?enable=false HTTP/1.1  (t=33.56s) → 204 No Content
```

Identical to previous DPM capture. Confirmed stable.

### 9.10 Summary — new findings from this capture

| Finding | Status |
|---|---|
| UDP port 1069 device discovery/heartbeat | ✓ Confirmed |
| `GET /events?enable` cold-start handshake | ✓ First time visible — confirmed |
| Initial Continuation handshake (init endpoints) | ⚠ Present, URLs unknown — Follow TCP Stream needed |
| `GET /parameters.xml` | ✓ New endpoint confirmed; content unknown |
| `GET /device_info.xml` → 401 | ✓ Confirmed — DMST cannot authenticate this endpoint |
| `GET /monitormode?enable=true/false` | ✓ Re-confirmed (×4 cycles) |
| Trigger = Continuation packets (×2) → 204 each | ✓ Pattern confirmed; URL unknown |
| `PUT /svg_image.img` (device → DMST) | ✓ **CRITICAL NEW FINDING** — push model confirmed |
| `GET /svg_image.img` → 500 = wrong direction | ✓ Explained — GET is wrong, PUT is correct mechanism |

