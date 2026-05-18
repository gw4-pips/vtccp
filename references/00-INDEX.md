# VTCCP Reference Library — Index

Front door for every reference asset, sample, manual, probe result, and design
note for the VCCS DMV TruCheck Command Pilot (VTCCP) project.

Built incrementally. If you're looking for something and can't find it, check
the unfiled-pile catalog at `attached_assets-archive/INDEX.md`.

---

## Directory map

| Path | Contains | Status |
|---|---|---|
| `manuals/cognex/DM475V/` | Reference manual for the production device (DM475V, fw 6.1.16_sr4) | Filed (PDF + extracted .txt + README) |
| `manuals/cognex/DMV-8072V/` | Reference manual for the Webscan-lineage 8072V verifier | Filed (PDF + digest) |
| `manuals/cognex/DMST/2025-release/` | DataMan Setup Tool reference manual (2025.4.1.1) | Filed (PDF + digest) |
| `manuals/cognex/DMST/2026-release/` | 2026 DMST manuals (when migrated) | Empty — pending (install in progress 2026-05-18) |
| `manuals/cognex/comms-and-programming-guide/` | DMCC command reference (2025.4.1.1) | Filed (PDF + digest) |
| `manuals/cognex/SDK-docs/` | DataMan SDK API docs (if acquired) | Empty — pending |
| `manuals/cognex/reference-manual-DM390-25.4.1.2.pdf` | Bonus — non-V reader, useful for SBG context | Filed |
| `manuals/webscan-legacy/` | Pre-acquisition Webscan TruCheck docs | Empty — to acquire |
| `standards/` | ISO standard PDFs/DOCX + per-standard markdown digests | Filed: 15415 (Ed.3 DRAFT), 15426-2 (3rd ed) + digests + INDEX. Pending: 15416, 29158, 16022, 18004 |
| `samples/live-scans/` | Raw XML output from real verifier, by date + push-script version | Filed (12 from current sprint + 3 from March) |
| `samples/live-scans/probe-history/` | Every probe-iteration XML | Filed |
| `samples/reports/` | DMST + Webscan PDF reports, for trade-dress reference + parser regression | Filed (16 PDFs) |
| `samples/excel/` | Real TruCheck XLS output — canonical schema source | Filed (1, more available in attached_assets) |
| `samples/images/` | Captured barcode images, organized by symbology + 2D/1D/DPM | Empty — pending categorization |
| `samples/scripts/` | Reference scripts (Cognex TruCheck CSV) | Filed |
| `samples/debug-logs/` | VtccpApp dev-time debug captures | Filed (9 from March sprint) |
| `push-script-history/` | Every shipped push-script version + CHANGELOG | Filed (v1.10, v1.11, v1.16, v1.23) |
| `probe-results/` | Parsed insights from JS probe sprint (metrics, siblings, etc.) | Empty — pending writeup |
| `architecture/` | Our own design notes (optics model, reverse-report, feature matrix) | Empty — pending writeup |
| `attached_assets-archive/INDEX.md` | Map of cryptic Replit-pasted filenames → meaningful descriptions | Filed |

---

## Authoritative sources by topic

| Topic | Primary source |
|---|---|
| 2D barcode print quality methodology (15415) | `standards/ISO-IEC-15415-ed3-2024-digest.md` + source PDF |
| 2D verifier conformance + tolerances (15426-2) | `standards/ISO-IEC-15426-2-2023-digest.md` + source DOCX |
| DM475V hardware behavior, calibration, SBG limits | `manuals/cognex/DM475V/reference-manual-25.4.1.1.pdf` |
| DMCC command syntax | `manuals/cognex/comms-and-programming-guide/comms-and-programming-guide-25.4.1.1.pdf` |
| DMST scripting environment + JS API | `manuals/cognex/DMST/2025-release/setup-tool-reference-manual-25.4.1.1.pdf` |
| Webscan-lineage column semantics (LLS/BLS/LQZ/etc) | `manuals/cognex/DMV-8072V/reference-manual-25.4.1.1.pdf` |
| Current shipped push script | `push-script-history/v1.23.js` (also live in `vtccp/DeviceInterface/Dmst/DmstPushScript_v1.js`) |
| Current 30-metric / 12-sibling enumeration | `samples/live-scans/v1.23-2026-05-18-Probe-DataMatrix-GS1Format06.xml` (raw); writeup pending |
| Real-device 167-column schema | `samples/excel/CalCardProd-2026-03-17.xls` + `samples/scripts/cognex-trucheck-csv-reference.js` |
| Webscan-format report (trade-dress reference) | `samples/reports/Webscan-format-2025-09-26.pdf` |
| Current DMST report format (what we're differentiating from) | `samples/reports/dmst-current-sprint/` + `samples/reports/dmst-2D/` |

---

## File-move policy

- **Originals preserved.** Files in `attached_assets/` are not deleted; only copied. The Replit asset pile is the audit trail.
- **`/tmp/` originals not preserved** — those are ephemeral. Files moved out of `/tmp/` to `references/` are now canonical here.
- **`vtccp/DeviceInterface/Dmst/DmstPushScript_v1.js`** is the live source of truth for the build. `push-script-history/v1.23.js` is the archived copy. They are identical at v1.23. When v1.24 ships, update both.

---

## See also

- `vtccp/README.md` — project README, phase status, build instructions
- `replit.md` — workspace overview + brief VTCCP section
- `VTCCP_Phase1_Tasks.docx` (top-level) — original Phase 1 plan, kept for historical record
