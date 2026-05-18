# Live-Scan Samples Index

Raw XML output captured from real DM475V verifier scans, organized by
push-script version and date.

## Authoritative samples

| File | Push-script vsn | Symbol | What it demonstrates |
|---|---|---|---|
| `v1.24-2026-05-18-QR-LoadedImage-Email.xml` | v1.24 | QR Code v5 (37×37), email payload | **Third loaded-image capture; confirms image-load fingerprint.** SymbologyId=`]Q1` (no ECI — no `\000026` prefix). SymbolAngle=**0** (not 360 — breaks prior heuristic). ApplicationPass=`Fail (X Dimension out of Range)`: NominalXDim=7.7 mil < 8 mil min. ApertureRef=06 (auto-selected for smaller module). AGValue=−0.5 (loaded-image artifact). JPEG len=19,388 chars. ContrastUniformity=−1, MRD=−1 confirmed as **reliable** discriminators (SymbolAngle is not). |
| `v1.24-2026-05-18-QR-LoadedImage-URL.xml` | v1.24 | QR Code v5 (37×37), URL payload | **First QR + second loaded-image capture.** SymbologyId=`]Q2` (ECI present — `\000026` UTF-8 prefix in DecodedData). SymbolAngle=360°, ContrastUniformity=−1, MRD=−1. JPEG len=22,260 chars. ApplicationPass=`Pass` (NominalXDim=8.7 mil within [8,30] range; quality=100). |
| `v1.24-2026-05-18-Probe-DataMatrix-GS1Format06.xml` | v1.24 | Data Matrix 16x36 ECC200 | **v1.24 live-confirmed.** 9 promoted first-class fields populated (SymbologyId=`]d1`, SymbolQuality=`41`, SymbolAngle=`1`, ModuleSizePx=`16.196`, CalibrationDate=`1/15/2026 3:04:14 PM`, FieldCalibrated/FactoryCalibrated=`false`, MinPassGrade=`NA`). 7 Debug* probes returned (see CHANGELOG v1.25 scope). Same symbol/cal as v1.23 file below. |
| `v1.23-2026-05-18-Probe-DataMatrix-GS1Format06.xml` | v1.23 | Data Matrix 16x36 ECC200 | Full 30-metric enumeration + 12 r-sibling enumeration + `r.symbology` deep structure. GS1 Format 06 payload (`[)>06 18VD89536 1P8902A S3122A02965`). UPCE-44960 cal symbol, June 2023. |
| `v1.23-2026-05-18-DMST26.1-DataMatrix-GS1Format06.xml` | v1.23 | Data Matrix 16x36 ECC200 | DMST 26.1.0 regression-diff baseline. Schema bit-identical to v25-baseline (above); 9 numeric drifts are measurement-level shot-to-shot variance. |
| `production-full-XML-2026-03-28.txt` | (pre-v1) production format | (various) | Full DMST live XML (149 KB) including `<image>` element — valuable for understanding what `r.image` ultimately becomes in output XML |

## DMST PDF Reports

| File | Scan | What it demonstrates |
|---|---|---|
| `v1.24-2026-05-18-QR-LoadedImage-Email-DMSTReport.pdf` | Email QR loaded-image | **First DMST PDF cross-reference.** Confirms QR v3=29×29 (push XML 35×35 is wrong). Shows Unit Serial, 15 QR grade params (ULP/URP/LLP/HCT/VCT/ALP/VIB/FIB absent from push). DMST epoch timestamp bug (31-Dec-1970) for loaded images. Full modulation grid (37×37 incl. quiet zone), codeword table, encodation analysis. See `v1.24-2026-05-18-QR-Email-DMSTReport-catalog.md` for full analysis. |
| `v1.24-2026-05-18-DM-Live-15434-DMSTReport.pdf` | Live DM, ISO 15434 enabled | **First live-scan DMST PDF + push XML pair.** ISO 15434 Format 06 parse table (DIs: 18V/1P/S). DataCodewords/ECBudget/ECCapacity empty for live DM too (not QR-specific). ULQZ/URQZ/RUQZ/RLQZ empty in push; TTR/CTR sub-grades = X in push but individual in DMST. AG discrepancy (push B / DMST A — push is correct). DebugModSize=684=18×38 confirms 1-module DM quiet zone. See `v1.24-2026-05-18-DM-Live-15434-catalog.md`. |

## Probe history

`probe-history/` contains every probe-iteration XML, named by cryptic Replit
timestamps for the current sprint (May 17-18) and by version+date for the
March sprint.

| File | Version | Date | Notes |
|---|---|---|---|
| `v1.10-2026-03-30-FullLive.xml` | v1.10 | 2026-03-30 | First probe iteration (3,551 chars) — vintage that C# parser was written against |
| `v1.11-2026-03-30-FullLive.xml` | v1.11 | 2026-03-30 | Expanded probe (6,961 chars) — parser-alignment reference |
| `Pasted--xml-version-1-0-encoding-UTF-8-0x0D-0x0A-DMCCResponse-_*.txt` | v1.18-v1.23 | 2026-05-17 to 18 | Current sprint probe iterations, 12 files, ordered by timestamp suffix |

## Conventions

- Filenames for renamed/canonical samples: `vX.YY-YYYY-MM-DD-Description.xml`
- `<0x0D><0x0A>` tokens in some files are Replit's literal-render of CRLF
  bytes. To expand: `sed 's/<0x0D><0x0A>/\n/g' file.txt > expanded.xml`
- Raw cryptic-named pastes in `probe-history/` are kept verbatim for audit;
  canonical/labeled versions live in this directory's root.
