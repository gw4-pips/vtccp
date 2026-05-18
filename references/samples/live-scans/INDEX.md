# Live-Scan Samples Index

Raw XML output captured from real DM475V verifier scans, organized by
push-script version and date.

## Authoritative samples

| File | Push-script vsn | Symbol | What it demonstrates |
|---|---|---|---|
| `v1.24-2026-05-18-Probe-DataMatrix-GS1Format06.xml` | v1.24 | Data Matrix 16x36 ECC200 | **v1.24 live-confirmed.** 9 promoted first-class fields populated (SymbologyId=`]d1`, SymbolQuality=`41`, SymbolAngle=`1`, ModuleSizePx=`16.196`, CalibrationDate=`1/15/2026 3:04:14 PM`, FieldCalibrated/FactoryCalibrated=`false`, MinPassGrade=`NA`). 7 Debug* probes returned (see CHANGELOG v1.25 scope). Same symbol/cal as v1.23 file below. |
| `v1.23-2026-05-18-Probe-DataMatrix-GS1Format06.xml` | v1.23 | Data Matrix 16x36 ECC200 | Full 30-metric enumeration + 12 r-sibling enumeration + `r.symbology` deep structure. GS1 Format 06 payload (`[)>06 18VD89536 1P8902A S3122A02965`). UPCE-44960 cal symbol, June 2023. |
| `v1.23-2026-05-18-DMST26.1-DataMatrix-GS1Format06.xml` | v1.23 | Data Matrix 16x36 ECC200 | DMST 26.1.0 regression-diff baseline. Schema bit-identical to v25-baseline (above); 9 numeric drifts are measurement-level shot-to-shot variance. |
| `production-full-XML-2026-03-28.txt` | (pre-v1) production format | (various) | Full DMST live XML (149 KB) including `<image>` element — valuable for understanding what `r.image` ultimately becomes in output XML |

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
