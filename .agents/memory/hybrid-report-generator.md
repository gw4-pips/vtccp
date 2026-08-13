---
name: Hybrid Report Generator
description: Self-contained HTML report merging Webscan TruCheck barcode grades + VCCS FlexWedge RFID validation. Fire-and-forget from SessionViewModel.
---

# Hybrid Report Generator

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
