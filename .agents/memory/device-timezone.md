---
name: Device timestamp timezone rule
description: DM475V timestamp format and correct parse strategy — device is set to local timezone via DMCC, NOT UTC.
---

## Rule — confirmed 2026-08-18

The DM475V clock is set to the **operator's local timezone** (America/New_York) via DMCC `DEVICE.TIMEZONE`. It is NOT UTC.

The timestamp string in the push XML `<DateTime>` element has no timezone suffix (e.g., `2026-08-18T15:59:08`). It is already in local time.

**Correct parse:** `DateTime.TryParse(s, out dt)` — no `AssumeUniversal`, no `ToLocalTime()` call needed.

DMST filenames may contain the local device timestamp anywhere in the name (or no
timestamp at all). If present, use it as-is with `DateTimeStyles.None`; do not
assume it is a prefix. If absent, the HTML `Verified:` header is always present
and is the authoritative device-local report time.

## Why this matters

An earlier session (2026-08-18) applied `AssumeUniversal | AdjustToUniversal + ToLocalTime()` after assuming UTC (NTP). This was **wrong** and caused 4-hour offset in the other direction. The user clarified that `DEVICE.TIMEZONE` is set to America/New_York, making the device clock local Eastern. The fix was reverted.

## How to apply

- `DmstResultParser.cs` — `DateTime.TryParse(dtStrDirect, out verifyDt)` — no style flags.
- `DmstHtmlScraper.cs` — accept an HTML report when either its filename timestamp
  or its `Verified:` header is available. Prefer exact `Verified:` equality for
  Manual-mode file correlation; only use filename-versus-record comparison when
  the incoming record lacks `Verified:`.
- VCCS PDF — display filename time only when present; otherwise preserve raw
  HTML `Verified:` text. Do not derive report time from a generated PDF name.
- Do NOT add `AssumeUniversal` unless the user confirms the device timezone has been reset to UTC.

## Verification

If timestamps in the VCCS report appear 4–5 hours ahead of local time: check `DEVICE.TIMEZONE` via DMCC. If it reads `UTC` or `0`, the device was factory-reset — restore to `America/New_York` (or the local TZ string) and REBOOT. Do NOT fix in code by adding `AssumeUniversal`.
