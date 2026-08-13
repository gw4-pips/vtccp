---
name: Hybrid Report Generator
description: Self-contained HTML report merging Webscan TruCheck barcode grades + VCCS FlexWedge RFID validation. Fire-and-forget from SessionViewModel.
---

# Hybrid Report Generator

## Layout — must match Webscan PDF 2-column structure

The Webscan PDF renders with this layout (confirmed from D61423AFE1783 PDF + screenshot):

**Full-width (span="all") sections:**
1. Report Summary
2. Verification Grades
3. GradingInfoSection

**Two-column side-by-side sections (explicit `<table>` with two `<td>` cells):**
4. Image (left, `col width="3.5in"`) | ISO Quality Parameters (right)
5. Data Format Check (left) | RFID Validation (right)

The Webscan XSLT achieves this via CSS CDO/CDC tokens activating `display:inline-block`/`float:left` on `<li>` and `-moz-column-count:2` on `body`, with `span="all"` for full-width overrides. Our generator replicates the visual result using explicit 2-column `<table>` layout — simpler and more reliable.

**Why:** Using an explicit table avoids dependence on browser-specific CSS column behavior and works identically in all browsers and PDF print.

**How to apply:** Any new "side by side" section must be added as a new row in either the Image+Params table or the Format+RFID table, or as a new `<li>` with a 2-col `<table>`. Never add them as full-width `<div span="all">` blocks unless they genuinely need full width.

## Section builder token map (v2.1)

| Token | Builder method | Width |
|---|---|---|
| `VCCS:GRADING_INFO_SECTION` | `GradingInfoSection(r)` | full |
| `VCCS:IMAGE_AND_PARAMS` | `ImageAndQualityTable(r)` | 2-col table |
| `VCCS:FORMAT_AND_RFID` | `FormatAndRfidTable(r)` | 2-col table |

## Mode-aware write paths (HybridReportMode)

Two public write methods exist:
- `SaveAsync(record, outputDir, ct)` — writes a timestamped filename to `outputDir` (Alongside mode)
- `SaveToPathAsync(record, outputPath, ct)` — writes to an exact full path including `.html` extension (Replace mode)

## Replace mode ordering constraint (critical)

In `DmstHtmlScraper.OnFileCreated`, the file **must be deleted before** `_pending.Add(report)`. If the pending entry were added first, `TryMergeAsync` could consume it and trigger a hybrid write while the watcher's `File.Delete` was still queued — silently deleting the freshly-written replacement. The ordering "read → parse → diagnostic copy → delete → _pending.Add" is intentional and must be preserved.

**Why:** This prevents a race where the hybrid write and the original delete execute concurrently on the same path, with the delete winning and leaving no file.

**How to apply:** Any future change to `OnFileCreated` must keep `_pending.Add` as the last step, after all file I/O on `e.FullPath` is complete.

## Alongside vs Replace scraper behaviour

- `DmstHtmlScraper.DeleteAfterParse = true` (Replace): original Webscan HTML is deleted; hybrid writes to same path via `SaveToPathAsync`
- `DmstHtmlScraper.DeleteAfterParse = false` (Alongside): original stays on disk; hybrid written to session/custom dir via `SaveAsync`
- `RegisterOwnedPath(path)` suppresses the FileSystemWatcher `Created` event for VTCCP-written hybrid files (one-shot per path)

**Why:** Both `StartHttpSubscriberAsync` and `StartPushListenerAsync` call `TryMergeAsync` when `_scraper` is active; `WebscanSourcePath` travels on `VerificationRecord` so it is atomic and pipeline-safe through all subsequent awaits.

## XSLT source encoding

`attached_assets/html_stylesheet_1786632621804.xslt` is **UTF-16 LE**. All grep/sed on it must use `iconv -f UTF-16 -t UTF-8` first — plain grep returns no output on this file.
