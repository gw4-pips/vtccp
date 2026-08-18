---
name: Device timestamp timezone rule
description: DM475V timestamp format and correct parse strategy — device is set to local timezone via DMCC, NOT UTC.
---

## Rule — confirmed 2026-08-18

The DM475V clock is set to the **operator's local timezone** (America/New_York) via DMCC `DEVICE.TIMEZONE`. It is NOT UTC.

The timestamp string in the push XML `<DateTime>` element has no timezone suffix (e.g., `2026-08-18T15:59:08`). It is already in local time.

**Correct parse:** `DateTime.TryParse(s, out dt)` — no `AssumeUniversal`, no `ToLocalTime()` call needed.

The DM TC HTML filename prefix (e.g., `2026-08-18_15-59-08`) is also in local time (device clock). Parse with `DateTimeStyles.None`.

## Why this matters

An earlier session (2026-08-18) applied `AssumeUniversal | AdjustToUniversal + ToLocalTime()` after assuming UTC (NTP). This was **wrong** and caused 4-hour offset in the other direction. The user clarified that `DEVICE.TIMEZONE` is set to America/New_York, making the device clock local Eastern. The fix was reverted.

## How to apply

- `DmstResultParser.cs` — `DateTime.TryParse(dtStrDirect, out verifyDt)` — no style flags.
- `DmstHtmlScraper.cs` filename parse — `DateTimeStyles.None`.
- Do NOT add `AssumeUniversal` unless the user confirms the device timezone has been reset to UTC.

## Verification

If timestamps in the VCCS report appear 4–5 hours ahead of local time: check `DEVICE.TIMEZONE` via DMCC. If it reads `UTC` or `0`, the device was factory-reset — restore to `America/New_York` (or the local TZ string) and REBOOT. Do NOT fix in code by adding `AssumeUniversal`.
