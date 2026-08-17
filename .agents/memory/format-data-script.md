---
name: Format Data script is the push script
description: The v1.37 Format Data script is the only script the user ever deploys; COM.SCRIPT is irrelevant to VtccpApp.
---

# Format Data Script = The Push Script

**Canonical file:** `artifacts/script-viewer/src/v137.txt` (272 lines, v1.37)

**Install path in DMST:** Format Data → Script-Based Formatting → Scripting tab → paste → Save Settings → Write to device

**What it does:** Runs on the device after each decode; formats output as `<DMCCResponse><DMSymVerResponse>…</DMSymVerResponse></DMCCResponse>` XML. VtccpApp receives this via its HTTP event subscription (`GET /events?enable` on port 44444 → device pushes `PUT /codes.xml`).

**COM.SCRIPT is irrelevant:** COM.SCRIPT is a separate device feature for pushing to FTP/email/external HTTP endpoints. VtccpApp does not use it. All prior work on deploying DmstPushScript_v1.js to COM.SCRIPT via raw TCP was misdirected and has been removed.

**Why:** Confirmed by reading HttpEventSubscriber.cs (port 44444, GET /events?enable) and the v1.37 install header. The user has always only pasted the Format Data script and the system worked.

**How to apply:** Never reference COM.SCRIPT in the context of VtccpApp result delivery. If auto-deploy is ever implemented, the target is the Format Data (Script-Based Formatting) slot, not COM.SCRIPT.

**Confirmation tag:** `<PushScriptDiag>v1.37 q=r.trucheck m=found</PushScriptDiag>` — look for this in exported records to confirm correct version is running.
