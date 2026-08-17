---
name: PowerShell 5.1 encoding trap
description: PS 5.1 reads .ps1 files as Windows-1252, not UTF-8; non-ASCII chars inside strings silently close them
---

## Rule
**All PowerShell scripts in this project must be pure ASCII.**

Non-ASCII characters in string literals cause parse failures on Windows PowerShell 5.1 (the version that ships with Windows 10/11). PS 5.1 reads `.ps1` files using the system code page (Windows-1252 on English Windows), NOT UTF-8.

**Why:** UTF-8 byte `0x94` (third byte of the em-dash `—`, E2 80 94) decodes as a RIGHT DOUBLE QUOTATION MARK in Windows-1252. This silently closes any double-quoted string mid-line, causing cascade parse errors — typically reported as "token '||' is not a valid statement separator" or "unexpected token" on lines well after the actual offender.

Confirmed offenders:
- `—` (em-dash, U+2014) → UTF-8 E2 80 **94** → `"` in Win-1252 → closes strings
- `✓` (checkmark, U+2713) → UTF-8 E2 9C **93** → `"` in Win-1252 → opens strings

**How to apply:**
- Write all string content with ASCII only. Use ` - ` instead of `—`, `(OK)` instead of `✓`, `--` instead of `—` in comments.
- If Unicode is required, save the file with a UTF-8 BOM (`\xEF\xBB\xBF` as first three bytes) — PS 5.1 honours the BOM and switches to UTF-8.
- To check a script before committing: `python3 -c "d=open('f.ps1','rb').read(); print([hex(b) for b in d if b>127])"`

**Fix applied to:** `vtccp/tools/Deploy-PushScript.ps1` — em-dash and checkmark replaced with ASCII equivalents.
