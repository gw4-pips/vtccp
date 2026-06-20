# CP UI Strategy Log — 2026-06-13
Version 1.0 | Session: 2026-06-13

---

## 1. Compendium Rewrite — COMPLETE

`references/screenshot-compendium.html` fully rewritten with W/D/V/S structure.

| Letter | Content |
|---|---|
| W | Webscan TruCheck UI (W1 TC0374 companion · W2 WTC Menus/Dialogs · W3 Settings · W4 Single Results · W5 Multi Mode · W6 Reports) |
| D | DMST/Cognex TruCheck UI (D1 App/Connect · D2 TC Config · D3 Results) |
| V | VTCCP Command Pilot UI (V1 Session Launcher · V2 Job Templates · V3 Excel Output) |
| S | Supporting Context (S1 Network · S2 DMST HTML Report · S3 Calibration · S4 Connect/Logging · S5 Barcode Samples · S6 Grade Params) |

---

## 2. CP Operational Modes — Agreed

Five discrete modes. Feature levels kept separate.

| Mode | Name | Description | Status |
|---|---|---|---|
| 1 | Live RDC | Real-time DataMan capture — Push (DMST) and Manual Trigger sub-modes | Implemented |
| 2 | Offline Image Grading | Batch IMAGE.LOAD sessions; device as pure grading engine; no live scan required | Logged, unimplemented |
| 3 | Device Administration | DMCC-native full device config, audit log, template enforcement | Blocked on TC window screenshots |
| 4 | Session Archive & Report | Ex-post-facto report generation from archived XML; no device needed | Architecture in place; UI not built |
| 5 | WTC USB Capture | Capture from Webscan TruCheck USB verifier into VTCCP schema | Design phase (this session) |

---

## 3. WTC Compatibility Strategy — Agreed

### The Bridge Play
WTC mode is a migration path for existing Webscan TruCheck USB users who have resisted DataMan because TC0374 Excel compliance reporting has no DataMan equivalent. CP solves that resistance and provides a staged migration:

```
Step 1: Webscan hardware + CP → operator learns CP, keeps Webscan, gets VTCCP output
Step 2: DataMan hardware + CP → hardware swap is invisible at the software level
Step 3: Full CP feature set — ex-post-facto reports, device config, batch grading
```

Cognex blessing recommended before investing in the full WTC UI clone — Cognex's sales org becomes a distribution channel if they endorse it.

### Three Excel Output Tiers
Configured per Job Template, not per scan. Set once, sticky.

| Tier | Name | Description |
|---|---|---|
| 1 | WTC Exact | Column-for-column identical to TC0374 output. CP's additional columns hidden (present in file, column-width zero). Zero operator retraining. |
| 2 | WTC+ | WTC columns intact in original order + CP enhancement columns appended right: Formal Grade, GS1 GTIN/Lot/Expiry, ISO Standard Edition, Sidecar Reference. Recommended default for WTC mode. |
| 3 | CP Full | Current VTCCP TruCheck-Compatible Schema. Default for DataMan mode. |

### Corrected CP-vs-TC0374 Delta
**TC0374 DOES produce a permanent saved Excel file — that is its entire purpose.**
Prior table was wrong to list this as a TC0374 gap. Corrected delta:

| Feature | TC0374 | CP WTC Mode |
|---|---|---|
| Permanent Excel session record | ✓ | ✓ |
| Report reprint without re-scanning | ✗ session ends = no reprint ever | ✓ XML sidecar makes any past session reprintable |
| GS1 AI parsing | ✓ parsed data in TC0374's own result table | ✓ same parsing **plus** dedicated Excel columns per AI (GTIN, Lot, Expiry, Serial) — individually addressable in formulas and pivot tables |
| Job template enforcement | ✗ | ✓ |
| DataMan hardware support | ✗ | ✓ (Step 2) |

---

## 4. WTC Capture Mode — Architecture Settled

### Core Principle
CP in WTC mode is a **silent parallel capture layer**. The operator workflow is unchanged.

```
Operator: opens WTC → configures in WTC → scans in WTC → sees results in WTC
                                       ↓
                   CP runs alongside, captures each result as it arrives,
                   writes to VTCCP Excel schema (selected tier)
                   Operator never needs to look at CP during a session
```

CP does NOT replicate WTC's settings UI, scan control, or result display. Identical relationship to how CP handles DMST — DMST is the operator's primary interface; CP is the capture layer.

### Capture Channels (priority order)

| Channel | Mechanism | Analogy | Status |
|---|---|---|---|
| Primary | ASCII/XML stream over COM port | DMST HTTP push → codes.xml | Needs Portmon investigation |
| Secondary | XML/CSV file save (user configures in WTC settings) | DMST pcm_report.html filesystem | Clean fallback; proven pattern |
| Deprioritized | RAM scrape via IWebBrowser2 COM automation | — | Technically possible; fragile; skip |

**Portmon investigation needed**: run Portmon (Sysinternals, free) + Process Monitor during a live WTC + TC0374 session. Capture: (a) raw COM port bytes during a scan, (b) any file writes from TC0374 process. Expected protocol: pure ASCII, likely XML-framed. ~90% confidence from user.

### WTC Capture Screen — Minimal by Design

```
Job: [name]     Operator: [id]     Roll: [number]

        [ START CAPTURE ]

  Status: Waiting for WTC scan...

  ● Record 7  UPCA  Grade B  14:23:11
  ● Record 6  DM    Grade A  14:21:44
  ● Record 5  DM    Grade A  14:20:02

  7 records captured → TC-2026-06-13.xlsx
```

- Job/Operator/Roll metadata: pre-entered in CP if not present in COM stream; auto-populated if stream includes them (Portmon will confirm)
- Output tier selector: in Job Template settings, not on capture screen
- No WTC settings replication anywhere in CP

---

## 5. COM Port Intercept Concept — Pending Verification

**Hypothesis (user-raised):** WTC may not seize the COM port until the operator clicks the Excel Functions icon in TC0374. If true, CP has a window to intercept.

**Proposed architecture if hypothesis is confirmed:**

CP runs as a **system tray app**. On detecting WTC process launch, CP watches for COM port claim. If port is still free:

> *"WTC detected — use TC0374 Excel or CP Capture?"*
> ☐ Remember this choice (sticky)

- Sticky setting lives in CP preferences
- Tray icon right-click always overrides for next session
- If "always CP" is set: CP silently takes port, no prompt, TC0374's Excel function will error if invoked (tells operator CP is active)

**Verification step**: Portmon — open WTC, note exactly when COM port claim appears. Before or after Excel icon click is the only question.

**If hypothesis is wrong** (WTC seizes port at launch): the choice is simply which app the user opens. CP detects TC0374 holding the port and surfaces a conflict message. No intercept possible; no tray mechanism needed.

---

## 6. Open Questions — Deferred

| Question | What We Need |
|---|---|
| When does WTC seize COM port? | Portmon session |
| WTC COM protocol format | Portmon output — first 200 bytes of a scan record |
| WTC file save format options | User to check WTC Settings > Save dialog |
| Founder protocol knowledge | Optional, not blocking Architecture B |
| CP top-level navigation structure | Future session |
| Mode entry point design (Job Template vs app-level) | Future session |
| DataMan capture screen — convergence with WTC capture screen? | Future session |
| Report Center / Archive screen design | Future session |
| Both-connected (WTC + DataMan simultaneously) — launch or future? | Future session |

---

## 7. Pending Tasks (Logged)

1. **WTC capabilities/features/shortcomings synopsis** — formal write-up; reference for D1 report design and CP feature decisions
2. **CP high-level UI concepts** — main screen layout exploration (Canvas recommended for side-by-side variants)
3. **Mode 3 (Device Administration)** — blocked until user provides detailed DMST TC window panel screenshots

---

*Next session picks up at: CP top-level navigation and main screen layout.*

---

## 8. Open Issues — Images Sheet (logged 2026-06-20)

### Issue: Col B base64 truncation + sidecar comment is wrong

**Col B** in the Images sheet label row holds the raw base64-encoded JPEG string of the scan image — identical to the visual embed in the row below it. Its purpose is D2 reverse-report reconstruction (re-generate the report from the Excel file without re-scanning).

**Problem 1 — Truncation**: Excel enforces a hard 32,767-character cell limit. `ImagesSheetWriter` caps at 32,000 and appends `[TRUNCATED]`. Any image whose base64 exceeds that limit is unrecoverable from col B. Larger DM symbols and QR codes at higher versions are at risk.

**Problem 2 — Wrong sidecar comment**: `ImagesSheetWriter.cs` line 26 states *"the sidecar (SessionSidecar) always stores the full payload and is the authoritative source for D2 round-trip."* This is incorrect — `SessionSidecar.cs` stores only session metadata (job name, operator, device info, counters). Image bytes are not written to any sidecar. If col B is truncated, the full image is currently lost.

**Required fix**: Either (a) write the full base64 to per-scan sidecar files alongside the Excel output (e.g. `scan_001.b64` or `scan_001.jpg`), or (b) remove col B and rely solely on the embedded visual image (losing D2 round-trip capability). Decision needed before D2 is implemented.

**Do not implement until the approach is chosen.**
