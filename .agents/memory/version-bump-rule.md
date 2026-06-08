---
name: Version bump rule
description: Always bump VtccpApp <Version> in the .csproj with every code-change commit so the user can confirm the rebuild via the title bar.
---

## Rule
Every commit that changes VtccpApp runtime behaviour must increment `<Version>` in `vtccp/VtccpApp/VtccpApp.csproj`.

**Why:** The user reads the version string in the title bar to confirm that `git pull` + rebuild picked up the new binary. Without a bump they cannot distinguish the old binary from the new one on Windows.

**How to apply:** After any edit to VtccpApp source (ViewModels, Views, DeviceInterface wiring, etc.), open `vtccp/VtccpApp/VtccpApp.csproj` and increment the patch component (e.g. 1.2.2 → 1.2.3). Do this in the same commit as the code change.

**Current version as of last change:** 1.2.3 (Live View TRIGGER.TYPE 5 / SetTriggerTypeAsync / WriteAndDrainAsync / OnResultReceived guard)
