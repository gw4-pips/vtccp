---
name: Version bump rule
description: Always bump VtccpApp <Version> in the .csproj with every code-change commit so the user can confirm the rebuild via the title bar.
---

## Rule
Every commit that changes VtccpApp runtime behaviour must increment `<Version>` in `vtccp/VtccpApp/VtccpApp.csproj`.

**Why:** The user reads the version string in the title bar to confirm that `git pull` + rebuild picked up the new binary. Without a bump they cannot distinguish the old binary from the new one on Windows.

**How to apply:** After any edit to VtccpApp source (ViewModels, Views, DeviceInterface wiring, etc.), open `vtccp/VtccpApp/VtccpApp.csproj` and increment the patch component (e.g. 1.2.2 → 1.2.3). Do this in the same commit as the code change.

## Extension: HTML preview files (`dist/vccs-pdf-preview-vN.html`)
Every round of edits to a PDF preview HTML file must save under the next version number — never edit in place and leave the filename unchanged.

**Why:** The user tracks iterations by version number. Editing v5 twice without creating v6 breaks that audit trail. User explicitly corrected this on 2026-08-14.

**How to apply:** When making any change to `dist/vccs-pdf-preview-vN.html`, copy to `vccs-pdf-preview-v(N+1).html`, update the internal title tag and print-hint string, then make edits to the new file only. Leave the prior version untouched as a reference.
