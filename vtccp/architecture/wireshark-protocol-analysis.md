# Wireshark Protocol Analysis — DM475V ↔ DMST Full Session Capture

**Capture**: "Wireshark - 475V Quit, Open, Connect, Verify"
**Archived**: `vtccp/architecture/gui-reference/wireshark-dmst-full-capture.txt` (7200 lines)
**Session date**: 2026-05-25 (DM475-63530E-PIPS-Verif-Lab, fw 6.1.16_sr4)
**Capture type**: Single TCP stream — Follow TCP Stream export from Wireshark
**Analyzed**: 2026-05-25
**Version**: 13 — **2026-06-25**: §9.7, §9.9, §9.10, §10.10 updated — trigger URL fully confirmed (`TRIGGER /on` / `TRIGGER /off`); §11 added (command channel full TCP stream analysis); **§12 added** — full-session unfiltered capture: verification toggle mechanism (`PUT /vs.cfg` DMST→device), config backup endpoints (`GET /config.cfg`, `GET /config.cdc`), device config-change push (`PUT /status.xml` device→DMST on events channel), complete init sequence, external telemetry SNI confirmation, UDP discovery protocol

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

### 9.7 ★ Trigger command — CONFIRMED 2026-06-25

When the user clicks **Go Live** from Sleep mode to trigger a verification scan, two
custom HTTP verbs fire in rapid succession on the command channel (54768):

```
→ TRIGGER /on HTTP/1.1
  Date: Thu, 25 Jun 2026 00:46:46 GMT
  X-Peer: 62476613

← HTTP/1.1 204 No Content
  Content-Length: 0
  Cache-Control: no-cache
  Pragma: no-cache
  Connection: Keep-Alive
  Server: DM475/6.1.16 (DeviceID=50)

→ TRIGGER /off HTTP/1.1
  Date: Thu, 25 Jun 2026 00:46:46 GMT
  X-Peer: 62476613

← HTTP/1.1 204 No Content
  ...
```

Both verified from Follow TCP Stream on command channel (54768) — 22,489-line stream,
confirmed twice (00:46:46 and 00:47:00 within same session).

| Packet | Time | Request | Response |
|---|---|---|---|
| 2035 | t=21.19s | `TRIGGER /on HTTP/1.1` on 54768 | `HTTP/1.1 204 No Content` |
| 2038 | t=21.23s | `TRIGGER /off HTTP/1.1` on 54768 | `HTTP/1.1 204 No Content` |

Both return `204 No Content` within ~2ms. No body, no Content-Length beyond 0.

**Full "Go Live from Sleep" sequence on command channel:**

```
→ GET /monitormode?enable=false   (exit sleep/monitor mode)
← 204 No Content

→ TRIGGER /on                     (fire verification trigger)
← 204 No Content

→ TRIGGER /off                    (release trigger)
← 204 No Content

→ GET /monitormode?enable=true    (return to sleep mode)
← 204 No Content
```

The device cannot be in monitor mode and execute a trigger simultaneously. DMST exits
monitor mode, fires the trigger pair, then re-enters monitor mode.

**Scan result arrives ~550ms after TRIGGER /on** on the events channel (54767) as:
1. `PUT /vs.cfg` (235B, AES-encrypted)
2. `PUT /pcm_report.html` (full HTML verification report)
3. `PUT /codes.xml` (verification result + base64 push XML, `origin="common"`)
4. **`PUT /svg_image.img`** (annotated SVG scan image — see §9.8)
5. `PUT /status.xml` ×2

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

Now that the trigger sequence is confirmed (§9.7), the full command channel sequence
for the second half of the session reads:

```
→ GET /monitormode?enable=true HTTP/1.1   (t=25.39s) → 204  ← return to Sleep after scan #1
→ GET /monitormode?enable=false HTTP/1.1  (t=28.09s) → 204  ┐
→ TRIGGER /on HTTP/1.1                    (t=28.10s) → 204  │ Go Live from Sleep → scan #2
→ TRIGGER /off HTTP/1.1                   (t=28.10s) → 204  │
→ GET /monitormode?enable=true HTTP/1.1   (t=31.29s) → 204  ┘ return to Sleep after scan #2
→ GET /monitormode?enable=false HTTP/1.1  (t=33.56s) → 204  ← Cancel (no trigger follows)
```

"Cancel" = `GET /monitormode?enable=false` with no `TRIGGER /on` following.
"Go Live from Sleep" = `GET /monitormode?enable=false` + `TRIGGER /on` + `TRIGGER /off` + `GET /monitormode?enable=true`.

### 9.10 Summary — new findings from this capture

| Finding | Status |
|---|---|
| UDP port 1069 device discovery/heartbeat | ✓ Confirmed |
| `GET /events?enable` cold-start handshake | ✓ First time visible — confirmed |
| Initial Continuation handshake (init endpoints) | ⚠ Present, URLs unknown — Follow TCP Stream needed |
| `GET /parameters.xml` | ✓ New endpoint confirmed; content unknown |
| `GET /device_info.xml` → 401 | ✓ Confirmed — DMST cannot authenticate this endpoint |
| `GET /monitormode?enable=true/false` | ✓ Re-confirmed (×4 cycles) |
| `TRIGGER /on` + `TRIGGER /off` → 204 each | ✓ **FULLY CONFIRMED 2026-06-25** — custom HTTP verbs, no body |
| `PUT /svg_image.img` (device → DMST) | ✓ **CRITICAL NEW FINDING** — push model confirmed |
| `GET /svg_image.img` → 500 = wrong direction | ✓ Explained — GET is wrong, PUT is correct mechanism |


---

## 10. Events channel (54767) full TCP stream — 2026-06-24 evening

**Source**: Follow TCP Stream on connection A (54767) — `GET /events?enable` subscription channel  
**File**: `attached_assets/Pasted-GET-events-enable-HTTP-1-1-…_1782349614260.txt`  
**Lines**: 17,295

### 10.1 Stream open sequence

```
→ GET /events?enable HTTP/1.1
  Date: Thu, 25 Jun 2026 00:46:35 GMT
  X-Peer: 62476613

← HTTP/1.1 204 No Content
  Server: DM475/6.1.16 (DeviceID=50)
```

Device immediately begins pushing `PUT /status.xml` (~4630 bytes, ~1/sec).

### 10.2 ★ PUT /svg_image.img — format confirmed: SVG

```
PUT /svg_image.img HTTP/1.1
Content-Type: image/svg+xml
User-Agent: DM475/6.1.16 (DeviceID=50)
```

**The verification image is SVG, not JPEG.** The `.img` extension is misleading — the device
delivers a vector SVG file. This is the annotated scan image shown in DMST's verification panel.

Implications for VTCCP:
- SVG is text-based XML — can be rendered in a WebView, embedded in HTML reports, or
  parsed to extract grid/module annotation data
- No JPEG decode needed — SVG can be displayed directly
- The tail of the stream (lines 12805-12854) shows `<GridIntersection>` elements with
  Center, IdealCenter, Mod, Grade, IsBlack — this is the per-module annotation data embedded
  in the SVG body

### 10.3 origin="monitor" vs origin="common" — confirmed in codes.xml

| Value | Meaning | Count in this session |
|---|---|---|
| `origin="monitor"` | Background monitoring scan in sleep mode | 25 |
| `origin="common"` | Full triggered TruCheck verification | 2 |

Monitor scans carry a partial `<trucheck_verificaiton_result>` (DECODE=F only, no grade data).
Common scans carry the full result including all ISO 15415 grade parameters.

### 10.4 ★ `<full_string encoding="base64">` — push XML inside codes.xml

Every `PUT /codes.xml` (both monitor and common) contains:

```xml
<full_string encoding="base64">PD94bWwg…</full_string>
```

This is the **complete DMCC push XML response** (`<DMCCResponse><DMSymVerResponse>…`),
base64-encoded. Decoding it yields the same push XML that the DMCC push listener receives
on a separate connection.

**This means**: VTCCP can receive all grade data from the events channel `codes.xml` body
alone, without needing a separate DMCC push listener TCP connection. The events channel
delivers everything in a single stream.

### 10.5 Decoded push XML from triggered verification (origin="common")

From result id=12341 (trigger #2), decoded from `<full_string>`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<DMCCResponse>
<DMSymVerResponse>
  <DateTime>2026-06-24T20:47:00</DateTime>
  <SymbologyName>Data Matrix</SymbologyName>
  <DecodedData>gibgibgib…</DecodedData>
  <SymbologyId>]d1</SymbologyId>
  <SymbolQuality>42</SymbolQuality>
  <SymbolAngle>359</SymbolAngle>
  <ModuleSizePx>24.218093872070312</ModuleSizePx>
  <PushScriptDiag>v1.37 q=r.trucheck m=found</PushScriptDiag>
  <Source>DM475-DPM-866D76-VCCS-Verif-Lab</Source>
  <FormalGrade>2/C</FormalGrade>
  <OverallGrade>C</OverallGrade>
  <OverallGradeNumeric>2</OverallGradeNumeric>
  <GradingStandard>ISO 15415:2011</GradingStandard>
  <ApplicationStandard>Custom</ApplicationStandard>
  <ApplicationPass>Pass</ApplicationPass>
  <ApertureRef>08</ApertureRef>
  <Wavelength>660</Wavelength>
  <Lighting>45Q</Lighting>
  <UECPercent>42.9</UECPercent>  <UECGrade>C</UECGrade>
  <SCPercent>81.3</SCPercent>    <SCGrade>A</SCGrade>
  <MinReflectance>3</MinReflectance>
  <MODGrade>A</MODGrade>
  <RMGrade>C</RMGrade>
  <ANUPercent>3.9</ANUPercent>   <ANUGrade>A</ANUGrade>
  <GNUPercent>2.3</GNUPercent>   <GNUGrade>A</GNUGrade>
  <FPDValue>4</FPDValue>         <FPDGrade>A</FPDGrade>
  <DecodeGrade>A</DecodeGrade>
  <MatrixSize>26x26</MatrixSize>
  <HorizontalBWG>11</HorizontalBWG>
  <VerticalBWG>10</VerticalBWG>
  <EncodedCharacters>24</EncodedCharacters>   ← BUG #1 MAY BE RESOLVED IN v1.37
</DMSymVerResponse>
</DMCCResponse>
```

### 10.6 Push script v1.37 — confirmed on DPM device (already running on LBL)

`<PushScriptDiag>v1.37 q=r.trucheck m=found</PushScriptDiag>`

The DM475V-DPM (10.10.10.4) is running push script **v1.37**. This is not a new version
discovery — v1.37 was already confirmed on DM475V-LBL (10.10.10.7) at scan #16 (2026-06-20)
and had been running on that device for an extended period before the DPM machine was brought
into use. Both devices are on the same script version.

Key observation: `<EncodedCharacters>24` is present in v1.37 output — if this is now
populated consistently, the v1.30 bug #1 (EncodedCharacters dead path on fw 6.1.16_sr4)
may be resolved in newer push script versions. Verify by comparing v1.37 result vs DMST
HTML report value for the same scan.

### 10.7 codes.xml timing field

```xml
<trigger_time>572</trigger_time>   ← milliseconds from trigger command to scan complete
<decode_time>532</decode_time>     ← ms for decode step
```

Both triggered scans: trigger_time = 572–578ms. Reliable field for scan latency monitoring.

### 10.8 trucheck_verificaiton_result — complete structure (common scans)

```xml
<trucheck_verificaiton_result>
  <status>valid</status>
  <CalibrationDate>6/23/2026 2:12:00 AM</CalibrationDate>
  <OpticalVariant>DM475V</OpticalVariant>
  <SymbolData>
    <CalibrationState>0</CalibrationState>
    <SymbologyType>DataMatrix</SymbologyType>
    <DecodedData>…</DecodedData>
    <Base64Data>…</Base64Data>        ← decoded data as base64
    <VerificaitonTime>455</VerificaitonTime>
    <BlurTime>272</BlurTime>
    <ThreshTime>29</ThreshTime>
    <CanidateEvaluationTime>70</CanidateEvaluationTime>
    <PostDecodeTime>60</PostDecodeTime>
    <ReportSection sectionType="GradingInfo">   ← ISO + Custom grades
    <ReportSection sectionType="GradeHistory">  ← pass/fail window history
    <ReportSection sectionType="Table">         ← per-parameter table
  </SymbolData>
</trucheck_verificaiton_result>
```

### 10.9 status.xml — full body confirmed

4630-byte XML, pushed ~1/sec in sleep mode. Contains:
- `<read_stats>`: good/bad reads, trigger counts, decoded_symbols by read_setup index
- `<monitored_values>`: image timing stats (request time, acquisition length/gap, etc.)
- `<monitored_counters>`: PTP, MST (multi-scanner), buffer overflow counters

Not needed for VTCCP result parsing, but useful for connection health monitoring.

### 10.10 Summary — events channel is sufficient for full VTCCP result delivery

| Data needed | Source on events channel |
|---|---|
| All ISO 15415 grade fields | `PUT /codes.xml` → `<full_string>` → base64 decode → push XML |
| Formal grade / aperture / wavelength / lighting | Same |
| Decoded data string | Same |
| Verification image (SVG) | `PUT /svg_image.img` (`Content-Type: image/svg+xml`) |
| HTML report | `PUT /pcm_report.html` |
| Device status/health | `PUT /status.xml` |
| Monitor vs verification discriminator | `origin="monitor"` vs `origin="common"` in codes.xml |

**★ Trigger URL CONFIRMED 2026-06-25** — `TRIGGER /on` and `TRIGGER /off` are custom HTTP verbs
on the command channel (54768). No body, `204 No Content` response. See §9.7 and §11 for full detail.

**Complete DMST HTTP control protocol is now fully reverse-engineered.** All endpoints and
verbs are known. VTCCP can operate the DM475V independently of DMST.

---

## 11. Command channel (54768) full TCP stream — 2026-06-25

**Source**: Follow TCP Stream on connection B (54768) — DMST command/request channel  
**File**: `attached_assets/Pasted-RESUME-HTTP-1-1-Date-Thu-25-Jun-2026-00-46-35-GMT-X-Pee_1782350561788.txt`  
**Lines**: 22,489  
**Session timestamps**: 00:46:35 – 00:47:00 (25-second session, DM475V-DPM, 10.10.10.4)

### 11.1 Complete verb inventory — all HTTP requests on command channel

```
Line    Verb / Endpoint                         Response
------  --------------------------------------  ------------------
1       RESUME /                                200 OK (0-byte body)
14      ISALIVE /                               204 No Content
27      GET /vs.cfg                             200 OK (24,000B AES)
248     GET /parameters.xml                     200 OK (content unread)
552     GET /vs.cfg                             200 OK (repeat)
747     GET /status.xml                         200 OK
790     GET /device_info.xml                    401 Unauthorized
804     GET /device_info.xml                    401 Unauthorized (retry)
819     GET /status.xml                         200 OK (repeat)
862     GET /monitormode?enable=true            204 No Content  ← Go Live #1 (Sleep)
875     GET /svg_image.img                      500 Internal Server Error
1999    GET /svg_image.img                      500  ┐
3120    GET /svg_image.img                      500  │ DMST polling (wrong direction)
4230    GET /svg_image.img                      500  │ Device doesn't serve via GET
5416    GET /svg_image.img                      500  │ Interval: ~1100 lines / ~10s each
6551    GET /svg_image.img                      500  │
7676    GET /svg_image.img                      500  │
7689    GET /svg_image.img                      500  │ (double-poll at this point)
8851    GET /svg_image.img                      500  ┘
10010   GET /monitormode?enable=false           204  ← exit Sleep
10023   TRIGGER /on                             204  ┐ Go Live from Sleep → scan #1
10036   TRIGGER /off                            204  ┘
10049   GET /monitormode?enable=true            204  ← return to Sleep
10062   GET /svg_image.img                      500  ┐ polling resumes
11200   GET /svg_image.img                      500  │
11213   GET /svg_image.img                      500  │
12346   GET /svg_image.img                      500  │
13478   GET /svg_image.img                      500  │
13491   GET /svg_image.img                      500  │
14595   GET /svg_image.img                      500  │
15703   GET /svg_image.img                      500  ┘
16864   GET /monitormode?enable=false           204  ← Cancel (no trigger)
16877   GET /svg_image.img                      500
16890   GET /monitormode?enable=true            204  ← Go Live #3 (Sleep again)
16903   GET /svg_image.img                      500  ┐
18059   GET /svg_image.img                      500  │
18072   GET /svg_image.img                      500  │
19162   GET /svg_image.img                      500  │
20275   GET /svg_image.img                      500  │
21366   GET /svg_image.img                      500  ┘
22452   GET /monitormode?enable=false           204  ← exit Sleep
22465   TRIGGER /on                             204  ┐ Go Live from Sleep → scan #2
22478   TRIGGER /off                            204  ┘
        [stream ends — session close]
```

### 11.2 Confirmed verb set — complete

| Verb | Path | Purpose | Response |
|---|---|---|---|
| `RESUME` | `/` | Session resume / keepalive init | `200 OK` (0-byte) |
| `ISALIVE` | `/` | Heartbeat / connection test | `204 No Content` |
| `GET` | `/vs.cfg` | Fetch device config (AES-encrypted) | `200 OK` |
| `GET` | `/parameters.xml` | Full device parameter dump | `200 OK` |
| `GET` | `/status.xml` | Device status | `200 OK` |
| `GET` | `/device_info.xml` | Device identity | `401 Unauthorized` |
| `GET` | `/monitormode?enable=true` | Enter Sleep (monitor mode on) | `204 No Content` |
| `GET` | `/monitormode?enable=false` | Exit Sleep (monitor mode off) | `204 No Content` |
| `GET` | `/svg_image.img` | ❌ Wrong direction — device doesn't serve images via GET | `500 Internal Server Error` |
| **`TRIGGER`** | **`/on`** | **Fire verification trigger** | **`204 No Content`** |
| **`TRIGGER`** | **`/off`** | **Release verification trigger** | **`204 No Content`** |

### 11.3 Key observations

**`TRIGGER /on` + `TRIGGER /off` are the only custom non-standard HTTP verbs** that perform
device actions. `RESUME` and `ISALIVE` are init/heartbeat only.

**Interval between GET /svg_image.img polls**: ~10 seconds (every ~1124 stream lines).
Occasionally two polls fire back-to-back (lines 7676/7689, 11200/11213, 13478/13491, 18059/18072).
All return `500`. DMST is permanently polling the wrong direction during live/sleep mode.

**`GET /monitormode?enable=false` discriminates Cancel from Go Live from Sleep**:
- Cancel: `enable=false` alone — no trigger follows
- Go Live from Sleep: `enable=false` → `TRIGGER /on` → `TRIGGER /off` → `enable=true`

**`GET /device_info.xml` → `401 Unauthorized`**: fired twice during init, both fail.
DMST cannot authenticate this endpoint on fw 6.1.16_sr4. Content unknown.

**`GET /parameters.xml`** appears once during init (line 248). Large response, content AES-encrypted
or binary (unreadable in stream dump). Likely the full device parameter manifest — potentially
replaces a DMCC GET ALL scan at connect time.

### 11.4 VTCCP implementation — complete command channel protocol

```csharp
// All requests sent on command channel TCP connection with:
//   X-Peer: {peerId}   (DMST uses 62476613 — arbitrary, VTCCP picks any int)
//   Date: {RFC1123}

// Session open
RESUME / HTTP/1.1          → expect 200 OK
ISALIVE / HTTP/1.1         → expect 204 No Content

// Init reads
GET /vs.cfg HTTP/1.1       → 200 OK, AES body (ignore or decrypt)
GET /parameters.xml HTTP/1.1 → 200 OK (log for future analysis)
GET /status.xml HTTP/1.1   → 200 OK

// Enter Sleep (Go Live)
GET /monitormode?enable=true HTTP/1.1  → 204 No Content

// Fire verification scan
GET /monitormode?enable=false HTTP/1.1  → 204 No Content
TRIGGER /on HTTP/1.1                    → 204 No Content
TRIGGER /off HTTP/1.1                   → 204 No Content
GET /monitormode?enable=true HTTP/1.1   → 204 No Content
// → await PUT /codes.xml (origin="common") on events channel

// Cancel
GET /monitormode?enable=false HTTP/1.1  → 204 No Content
// (no trigger, no monitormode=true)
```

### 11.5 Protocol reversal status — COMPLETE

| Component | Status |
|---|---|
| Events channel subscription (`GET /events?enable`) | ✓ Confirmed §2.1 |
| Result delivery (`PUT /codes.xml`, `PUT /pcm_report.html`) | ✓ Confirmed §2.4, §2.5 |
| Image delivery (`PUT /svg_image.img` — SVG) | ✓ Confirmed §9.8, §10.2 |
| Go Live / Sleep (`GET /monitormode?enable=*`) | ✓ Confirmed §8.1 |
| **Verification trigger (`TRIGGER /on` / `TRIGGER /off`)** | **✓ CONFIRMED 2026-06-25 — this section** |
| Session init (`RESUME /`, `ISALIVE /`) | ✓ Confirmed §9.3 |
| Config fetch (`GET /vs.cfg`, `GET /parameters.xml`) | ✓ Confirmed (content AES-encrypted) |

**The complete DMST HTTP control protocol is now fully reverse-engineered.**
VTCCP requires no DMST process to be running to operate the DM475V for TruCheck verification.

---

## 12. Full-session unfiltered capture — 2026-06-25

**Capture**: Unfiltered — all interfaces, full session  
**Actions captured**: Open DMST → start TC → (no Go Live) → Save backup config → Disable verification → Re-enable × 2 → Stop  
**Lines**: 3,736 packets  
**Analyzed**: 2026-06-25

This capture was not filtered to the device IP, so it contains all network traffic from the DMST host (10.10.10.19) across the full session. It is the most complete capture to date and adds several previously unobserved events.

---

### 12.1 External services contacted by the DMST host — SNI confirmed

Two external TLS connections fire at session start (before DMST ever connects to the device):

| t= | Destination | SNI (from ClientHello — plaintext) | Who |
|---|---|---|---|
| 0.000s | 34.128.165.207:443 | *(already established — ACK only at pkt 1)* | Unknown (Cloudflare) |
| 0.186s | 18.223.69.42:443 | **`lambdaapi.superops.ai`** | SuperOps RMM agent |
| 0.563s | 18.97.138.67:443 | **`ingress.us1.coralogix.com`** | Coralogix log shipping |

**`lambdaapi.superops.ai`** — SuperOps is an IT RMM (Remote Monitoring & Management) / PSA
platform used by managed service providers to remotely monitor and manage endpoints. At t=186ms
DMST fires a ~26KB TLS 1.3 burst upload (packets 15–95 of the capture, rapid 325-byte TLS
records + one large 10,274-byte frame), then closes the connection with FIN at t=434ms.
This is a one-shot telemetry upload, not a persistent connection.

**`ingress.us1.coralogix.com`** — Coralogix is a cloud log analytics / observability platform.
DMST opens this connection at t=563ms and ships logs there.

**Critical timing note**: The device connection does not begin until t=8.97s (ARP + TCP SYN to
10.10.10.4:44444). The SuperOps upload completes and closes in 248ms — well before any device
interaction. This strongly suggests SuperOps is a **Windows background IT management agent**
installed on the DMST workstation, NOT a component of DMST itself. The Coralogix connection
could be either DMST log shipping or the same Windows IT agent.

> **Test to distinguish**: Close DMST and watch Wireshark for ~2 minutes. If SuperOps/Coralogix
> still fire on their ~10-second heartbeat schedule, they are Windows agents independent of DMST.
> If they stop, they originate from within DMST.

**Payload content**: All traffic is TLS 1.3 — payload is encrypted and unreadable without the
server private key. Packet sizes (97/93/82 bytes on heartbeats; 325-byte bursts on uploads) are
consistent with telemetry/log data, not bulk file transfer. No inference about specific data
content is possible.

**Air-gap implication**: If DMST is deployed in a network-isolated environment, these outbound
HTTPS calls will fail silently or after TCP timeout. Whether DMST or the IT agent degrades
gracefully is unknown.

---

### 12.2 Device UDP discovery broadcasts — NEW

Both DM475V devices (LBL = 10.10.10.7, DPM = 10.10.10.4) send UDP broadcasts at approximately
3-second and 6-second marks into the capture — **before DMST connects to them**:

| pkt | t= | Source | Destination | Port | Size |
|---|---|---|---|---|---|
| 237 | 3.41s | 10.10.10.7 | 255.255.255.255 | 1069 → 1069 | 223B |
| 243 | 3.60s | 10.10.10.7 | 10.10.10.19 (DMST) | 1069 → 63589 | 223B |
| 244 | 3.70s | 10.10.10.4 | 10.10.10.19 (DMST) | 1069 → 63589 | 223B |
| 249 | 3.90s | 10.10.10.4 | 255.255.255.255 | 1069 → 1069 | 223B |
| 380 | 6.29s | 10.10.10.7 | 10.10.10.19 | 1069 → 63589 | 223B |
| 381 | 6.39s | 10.10.10.4 | 10.10.10.19 | 1069 → 63589 | 223B |

UDP port 1069 is associated with Cognex device discovery / GigE Vision network presence announcements.
Both devices are broadcasting their presence to the subnet and unicasting directly to the DMST
host. The DMST host is evidently known to them (presumably from a prior connection or ARP cache).
Capture capture does not show DMST responding to these UDP packets.

---

### 12.3 Complete device init sequence (DPM, t=8.60–10.1s)

At t=8.60s the LBL device (10.10.10.7) issues an ARP request for the DMST host; DPM (10.10.10.4)
follows at t=8.71s. DMST then immediately opens two TCP connections to DPM port 44444:

| pkt | t= | Direction | Verb | Response | Notes |
|---|---|---|---|---|---|
| 444/445 | 8.97s | DMST→DPM | TCP SYN (63160→44444) | SYN-ACK | Channel 1 (command) |
| 447/448 | 8.97s | DMST→DPM | TCP SYN (63161→44444) | SYN-ACK | Channel 2 (data) |
| 450/452 | 8.99s | DMST→DPM | `Continuation` (= `RESUME /`) | `200 OK` (195B) | Session resume |
| 453/454 | 9.01s | DMST→DPM | `Continuation` (= `ISALIVE /`) | `204 No Content` (203B) | Heartbeat |
| 456/458 | 9.03s | DMST→DPM | `GET /events?enable HTTP/1.1` | `204 No Content` (203B) | Subscribe to events |
| 459/485 | 9.04s | DMST→DPM | `GET /vs.cfg HTTP/1.1` | `200 OK` (1,023B) | Fetch device config (~24KB body) |
| 497/531 | 9.63s | DMST→DPM | `GET /parameters.xml HTTP/1.1` | `200 OK` (1,113B) | Full parameter dump |
| 533/552 | 9.71s | DMST→DPM | `GET /vs.cfg HTTP/1.1` | `200 OK` (1,023B) | **Second vs.cfg fetch** |
| 559/564 | 9.83s | DMST→DPM | `GET /status.xml HTTP/1.1` | `200 OK` (529B) | Status poll |
| 566/567 | 9.86s | DMST→DPM | `GET /device_info.xml HTTP/1.1` | `401 Unauthorized` (260B) | Auth failure |
| 570/571 | 9.88s | DMST→DPM | `GET /device_info.xml HTTP/1.1` | `401 Unauthorized` (260B) | Retry, also fails |
| 579/583 | 10.08s | DMST→DPM | `GET /status.xml HTTP/1.1` | `200 OK` (529B) | Poll continues |

**Notes:**
- vs.cfg is fetched **twice** during init (t=9.04s and t=9.71s). The second fetch may be triggered
  by TC panel initialization after the first.
- `GET /device_info.xml` fails with 401 both times. DMST cannot authenticate this endpoint on
  fw 6.1.16_sr4, or the endpoint requires credentials DMST does not send. Content unknown.
  Matches §11.3 observation from the prior capture.
- `GET /events?enable` subscription fires before the config fetches — the device can begin
  pushing events immediately after t=9.03s.

---

### 12.4 Save backup config — NEW endpoints discovered (t=30.23s)

When the user clicks "Save backup config" in DMST, two sequential GET requests fire to the device:

```
→ GET /config.cfg HTTP/1.1         (pkt 2160, t=30.23s)
← HTTP/1.1 200 OK                  (pkt 2273, t=30.26s, 249B total)
   Body: ~83B  ← small metadata

→ GET /config.cdc HTTP/1.1         (pkt 2274, t=30.28s)
← HTTP/1.1 200 OK                  (pkt 2391, t=30.31s, 415B Wireshark frame)
   Body: ~136.5KB  ← large response (seq 232809 → ack 372562 = 139,753 bytes)
```

**`GET /config.cfg`** — small metadata file (body ~83B after HTTP headers). Extension `.cfg`
suggests a configuration descriptor or header file. Likely ASCII/XML.

**`GET /config.cdc`** — large 136.5KB response, reassembled across ~97 TCP segments before
Wireshark presents it as a single HTTP frame. The `.cdc` extension is Cognex-proprietary —
this is the full device configuration backup container. Content is binary or AES-encrypted
(unreadable from packet headers alone). This is what DMST saves to disk as the backup file.

**VTCCP implication**: These two endpoints provide a config snapshot path. `GET /config.cdc`
followed by a corresponding `PUT /config.cdc` would presumably restore a saved configuration.
The PUT direction has not been observed; it would only fire on "Restore from backup".
Both endpoints are additive to the Phase 1 plan and require no new plumbing.

---

### 12.5 Verification toggle mechanism — NEW: `PUT /vs.cfg` (DMST → device) (t=37.5–46.0s)

When the user enables or disables verification in TC, DMST fires this 3-step sequence:

```
→ GET /status.xml HTTP/1.1         ← read current device state (529B)
← HTTP/1.1 200 OK

→ PUT /vs.cfg HTTP/1.1             ← write new verification state to device (118B total)
  [small body — AES-encrypted delta or enable/disable flag]

← [device events channel]:
   PUT /status.xml HTTP/1.1        ← device pushes config-change notification (326B XML)

← HTTP/1.1 200 OK                  ← response to the PUT /vs.cfg
  [406B first toggle; 309B subsequent toggles]
```

Four toggles observed — matches user's "disable + re-enable × 2" sequence:

| pkt | t= | Action |
|---|---|---|
| 3374 | 37.51s | PUT /vs.cfg toggle 1 |
| 3463 | 40.48s | PUT /vs.cfg toggle 2 |
| 3583 | 43.74s | PUT /vs.cfg toggle 3 |
| 3653 | 46.04s | PUT /vs.cfg toggle 4 |

**Direction note**: There are now TWO distinct `PUT /vs.cfg` flows in the protocol:

| Direction | Channel | Content | Size | Trigger |
|---|---|---|---|---|
| Device → DMST | Events (long-poll) | AES-encrypted full config | ~288–400B | Periodic config sync |
| **DMST → Device** | **Command (data channel)** | **AES-encrypted delta or flag** | **~30–50B body** | **Verification enable/disable, config change** |

The DMST→device PUT body is estimated at ~30–50 bytes (118B frame − IP/TCP/HTTP header overhead).
Very small — likely a single AES block carrying just the changed parameter(s).

**VTCCP implication**: To implement verification enable/disable without DMST, VTCCP would need to:
1. GET /status.xml (confirm state)
2. PUT /vs.cfg with the appropriate small AES-encrypted body

The body content is unknown (AES-encrypted). **This specific PUT is NOT required for Phase 2**
(VTCCP controls the trigger, not the enable/disable — that is operator-set before a session).
It is logged here for completeness. If VTCCP ever needs to toggle verification programmatically,
the body format will require a targeted Wireshark decode session with TLS key export or
firmware analysis.

---

### 12.6 Device config-change push — `PUT /status.xml` (device → DMST, events channel)

Previously documented (§2.2): the device sends `PUT /status.xml` **every ~1 second** as
periodic telemetry (~4.6KB, on the events long-poll channel).

**NEW from this capture**: the device also sends `PUT /status.xml` **immediately** after any
configuration change, regardless of the 1-second telemetry schedule. Each verification toggle
produces a 326-byte `PUT /status.xml` within ~170ms of the DMST PUT /vs.cfg completing.

This is a **config-change notification** — smaller than the periodic 4.6KB telemetry
(326B vs 4,600B), and fires out-of-schedule. Content likely carries only the changed parameter
subset, not the full status.

**Distinguishing periodic from config-change `PUT /status.xml`**:

| Property | Periodic | Config-change |
|---|---|---|
| Trigger | ~1 second timer | Immediately after config write |
| Size | ~4,600B | 326B |
| XML content | Full `<status version="3">` tree | Likely partial / delta |

VTCCP's HttpEventsChannel should handle both variants — the parser already processes the
full tree; the smaller variant will parse as a partial update and should not cause errors
if all fields are nullable.

---

### 12.7 Updated complete verb inventory (all captures combined)

| Verb | Path | Direction | Response | First seen |
|---|---|---|---|---|
| `RESUME` | `/` | DMST→device | `200 OK` | §9 |
| `ISALIVE` | `/` | DMST→device | `204 No Content` | §9 |
| `GET` | `/events?enable` | DMST→device | `204 No Content` | §2.1 |
| `GET` | `/vs.cfg` | DMST→device | `200 OK` (~24KB AES) | §11 |
| `GET` | `/parameters.xml` | DMST→device | `200 OK` (1,113B) | §11 |
| `GET` | `/status.xml` | DMST→device | `200 OK` (529B) | §11 |
| `GET` | `/device_info.xml` | DMST→device | `401 Unauthorized` | §11 |
| `GET` | `/config.cfg` | DMST→device | `200 OK` (~83B) | **§12.4** |
| `GET` | `/config.cdc` | DMST→device | `200 OK` (~136.5KB) | **§12.4** |
| `GET` | `/monitormode?enable=true` | DMST→device | `204 No Content` | §8.1 |
| `GET` | `/monitormode?enable=false` | DMST→device | `204 No Content` | §8.1 |
| `GET` | `/svg_image.img` | DMST→device | `500` (wrong direction) | §9 |
| `TRIGGER` | `/on` | DMST→device | `204 No Content` | §9.7 |
| `TRIGGER` | `/off` | DMST→device | `204 No Content` | §9.7 |
| **`PUT`** | **`/vs.cfg`** | **DMST→device** | **`200 OK`** | **§12.5** |
| `PUT` | `/status.xml` (periodic) | device→DMST | *(long-poll push)* | §2.2 |
| `PUT` | `/status.xml` (config-change) | device→DMST | *(long-poll push, 326B)* | **§12.6** |
| `PUT` | `/vs.cfg` | device→DMST | *(long-poll push, AES)* | §2.3 |
| `PUT` | `/codes.xml` | device→DMST | *(long-poll push)* | §2.4 |
| `PUT` | `/pcm_report.html` | device→DMST | *(long-poll push)* | §2.5 |
| `PUT` | `/svg_image.img` | device→DMST | *(long-poll push, SVG)* | §9.8 |

**Total: 20 distinct verb/path combinations confirmed across all captures.**

---

### 12.8 Protocol reversal status — updated

| Component | Status |
|---|---|
| Events channel subscription (`GET /events?enable`) | ✓ Confirmed §2.1 |
| Result delivery (`PUT /codes.xml`, `PUT /pcm_report.html`) | ✓ Confirmed §2.4, §2.5 |
| Image delivery (`PUT /svg_image.img`) | ✓ Confirmed §9.8 |
| Go Live / Sleep (`GET /monitormode?enable=*`) | ✓ Confirmed §8.1 |
| Verification trigger (`TRIGGER /on` / `TRIGGER /off`) | ✓ Confirmed §9.7 |
| Session init (`RESUME /`, `ISALIVE /`) | ✓ Confirmed §9.3 |
| Config fetch read (`GET /vs.cfg`, `GET /parameters.xml`, `GET /status.xml`) | ✓ Confirmed §11 |
| Config backup download (`GET /config.cfg`, `GET /config.cdc`) | ✓ Confirmed **§12.4** |
| **Verification toggle write (`PUT /vs.cfg` DMST→device)** | ✓ **Confirmed §12.5** — body AES-encrypted, content unknown |
| Config-change notification (`PUT /status.xml` 326B, device→DMST) | ✓ **Confirmed §12.6** |

**20 verb/path combinations confirmed. Protocol reversal is complete for all DMST actions
observed to date.** The only remaining unknown is the AES body format of `PUT /vs.cfg`
(DMST→device), which is not required for any Phase 1–2 VTCCP operation.

