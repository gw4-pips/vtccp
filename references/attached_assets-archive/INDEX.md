# `attached_assets/` Cryptic-Filename Map

Replit auto-prefixes every uploaded/pasted file with `Pasted-` or `image_` plus
a Unix-ms timestamp, producing filenames like
`Pasted--xml-version-1-0-encoding-UTF-8-0x0D-0x0A-DMCCResponse-_1779054767388.txt`
that are impossible to identify by intent.

This index maps the cryptic names to meaningful descriptions. **Originals are
preserved in `attached_assets/`** — this directory just catalogs them. Most
items of long-term value have been *copied* into other `references/` subtrees
under meaningful names; the originals stay for audit trail.

---

## March 25-30, 2026 cluster — original build sprint

### Manuals (PDFs)
| Cryptic name | Meaning | Filed copy |
|---|---|---|
| `DataMan_CommunicationsAndProgramming_Guide,_2025.4.1.1_1774546662030.pdf` | DMCC command reference (2.8 MB) | `manuals/cognex/comms-and-programming-guide/` |
| `DataMan_SetupTool_Reference_Manual,_25.4.1.1_1774546662030.pdf` | DMST manual (6.5 MB, dup of `_1774546935164.pdf`) | `manuals/cognex/DMST/2025-release/` |
| `DataMan_SetupTool_Reference_Manual,_25.4.1.1_1774546935164.pdf` | DMST manual (6.5 MB) | `manuals/cognex/DMST/2025-release/` |
| `DM475_Verifier_Reference_Manual,_25.4.1.1_1774546874064.pdf` | DM475V hardware reference (7.9 MB) | `manuals/cognex/DM475V/` |
| `DM8072_Verifier_Reference_Manual,_25.4.1.1_1774546935164.pdf` | DMV-8072V hardware reference (5.2 MB) | `manuals/cognex/DMV-8072V/` |
| `DM390_Reference_Manual,_25.4.1.2_1774546935163.pdf` | DM390 reader (non-V, SBG context) | `manuals/cognex/reference-manual-DM390-25.4.1.2.pdf` |

### Sample reports (DMST PDF output)
| Cryptic name family | Meaning | Filed |
|---|---|---|
| `DataMatrix-25-*.pdf` | Plain Data Matrix sample scans | `samples/reports/dmst-2D/` |
| `GS1_DataMatrix-25-*.pdf` | GS1 Data Matrix sample scans | `samples/reports/dmst-GS1/` |
| `_F1_*.pdf` | GS1 result reports with `<FNC1>` data | `samples/reports/dmst-GS1/` |
| `OMNI_Wide_Angle_*.pdf` | 1D omnidirectional samples (UPC, EAN, defects, decodability) | `samples/reports/dmst-1D/` |
| `Test_Data_Matrix_for_fixed_pattern_grading_*.pdf` | Fixed-pattern grading test sample | `samples/reports/dmst-2D/` |
| `Test_for_UEC_*.pdf` | UEC test sample | `samples/reports/dmst-2D/` |
| `___00_*.pdf` and `2025-08-11_*.pdf` | Misc DMST scan output | `samples/reports/dmst-2D/` |
| `Webscan_Report--25-09-26_11_58_38_*.pdf` | **Webscan-format report — trade-dress reference** | `samples/reports/Webscan-format-2025-09-26.pdf` |

### Excel
| Cryptic name | Meaning | Filed |
|---|---|---|
| `CalCardProd,_2026-03-17_1774742144062.xls` | Production calibration-card multi-scan TruCheck XLS — canonical schema | `samples/excel/CalCardProd-2026-03-17.xls` |

### Scripts
| Cryptic name | Meaning | Filed |
|---|---|---|
| `TruCheck_Verification_CSV_Results_1779053543516.js` | Cognex-provided push-script reference (CSV template) — origin of our 167-column schema | `samples/scripts/cognex-trucheck-csv-reference.js` |

### Debug captures (VtccpApp dev sessions)
| Cryptic name | Meaning | Filed |
|---|---|---|
| `Debug#1_*.txt`, `Debug#1.9_*.txt`, `Debug#1.9A_*.txt`, `Debug#1.10_*.txt`, `Debug#1.11_*.txt`, `Debug#2_*.txt`, `Debug#3_*.txt` | Live VtccpApp debug dumps during March development | `samples/debug-logs/` |
| `Pasted-pp-exe-CoreCLR-*` and similar | VtccpApp startup / loaded-DLL dumps | `samples/debug-logs/` |

### Live XML pastes (March)
| Cryptic name | Meaning | Filed |
|---|---|---|
| `Pasted--xml-version-1-0-result-id-1939-image-id-1-version-3-or_1774740073000.txt` | Full production DMST live XML (149 KB) — includes `<image>` element, valuable for `r.image` analysis | `samples/live-scans/production-full-XML-2026-03-28.txt` |
| `Pasted--VTCCP-DMST-RawXML-6961-chars-xml-version-1-0-encoding-_1774911918584.txt` | v1.11 raw push-script output (March 30) | `samples/live-scans/probe-history/v1.11-2026-03-30-FullLive.xml` |
| `Pasted--VTCCP-DMST-RawXML-3551-chars-xml-version-1-0-encoding-_1774908915909.txt` | v1.10 raw push-script output (March 30) | `samples/live-scans/probe-history/v1.10-2026-03-30-FullLive.xml` |

### PNG screenshots (March — DMST UI captures)
~40 screenshots. Not individually catalogued — most are point-in-time DMST UI
captures useful for visual reference but not for ongoing work. Browse
`attached_assets/image_177473*.png` through `image_177494*.png` if you need a
March-sprint UI snapshot.

---

## May 17-18, 2026 cluster — current JS probe sprint

### DMCCResponse XML pastes (push-script probe iterations)
All 12 filed to `samples/live-scans/probe-history/` with original cryptic names
preserved (the timestamp ordering shows probe sequence). The v1.23 final paste
is also filed under a meaningful name:
- `samples/live-scans/v1.23-2026-05-18-Probe-DataMatrix-GS1Format06.xml`

The two byte-identical paste files
(`1.23_xml_version=1.0_encoding=UTF-80x0D0_1779094651799.txt` and
`Pasted--xml-version-1-0-encoding-UTF-8-0x0D-0x0A-DMCCResponse-_1779094062589.txt`)
are the same v1.23 scan — Replit's preview pane truncated the first one
visually, prompting a re-paste that turned out to be identical.

### DMST PDF reports (current sprint)
4 PDFs (`2026-05-17_*.pdf`) filed to `samples/reports/dmst-current-sprint/`.
These are the reports DMST generated during our current probe session.

### PNG screenshots (May)
~45 screenshots from current sprint. Mostly DMST UI captures, scan-result
dialogs, and Excel column views. Same policy: browse if needed, not
individually catalogued.

---

## What's NOT here

These were requested but never uploaded — still outstanding:
- 8072V older firmware manuals (pre-2025.4.1.1)
- DataMan SDK API documentation
- Pre-acquisition Webscan TruCheck PC software docs
- 2026 DMST release notes / manuals
- ISO standard PDFs (15415, 15416, 18004, 29158, 16022)
