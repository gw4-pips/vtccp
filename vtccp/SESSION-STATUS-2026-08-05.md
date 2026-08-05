# Session Status — 2026-08-05 (end of day)

## What was completed this session

### 1. TRUCHECK.APPLICATION-STANDARD — mapping confirmed and fixed
The DMCC Reference (rev 26.1.0.27) has the integer values for this parameter
**inverted** relative to actual firmware behaviour on fw 6.1.16_tc9.

| Integer | Reference claims | **Actual (empirically confirmed)** |
|---|---|---|
| 4 | Auto | **Custom** |
| 5 | Custom | **Auto** |

Test method: set UI to Custom → DMCC GET returned 4. UI was at Auto → returned 5.

**Fixed:** `vtccp/DeviceInterface/Dmcc/DmccCommand.cs` comment corrected.
**Filed:** `vtccp/references/cognex-bug-reports/TRUCHECK-APPLICATION-STANDARD-mapping-error.md`
— ready to send to Cognex support.

---

### 2. vtccp GitHub auto-push hook — created, one item outstanding

`vtccp/.githooks/post-commit` now exists and is wired to git
(`core.hooksPath = .githooks` confirmed in vtccp git config).

The hook tries these secret names in order:
1. `GITHUB_PAT2`
2. `GITHUB_TOKEN2`
3. `GITHUB_PAT`
4. `GITHUB_TOKEN`

**Outstanding:** The user created a new PAT secret whose name ends in "2" but
the exact name was not confirmed before leaving. The hook will find it
automatically on the next environment restart **as long as the name is one of
the four above**.

**Action needed on return:**
- Tell the agent the exact secret name used, so it can tighten the hook
  (remove the wrong candidates).
- Restart the Replit environment so the new secret is injected into the shell.
- Make one test commit in vtccp and confirm "[post-commit] GitHub push OK" appears.

---

### 3. DM475V-DPM housekeeping params — still pending (Task #54 cancelled)
Six params not yet set on the DPM unit (10.10.10.4):

| Parameter | Current | Target |
|---|---|---|
| DEVICE.NAME | DM475-866D76 | DM475-DPM-866D76-PIPS-Verif-Lab |
| DEVICE.TIMEZONE | UTC | America/New_York |
| NTP.ENABLE | OFF | ON |
| NTP.SERVER1 | — | time.nist.gov |
| NTP.SERVER2 | — | 132.163.96.1 |
| TRUCHECK.OPERATOR-NAME | — | GW4 |
| TRUCHECK.COMPANY-NAME | — | Product Identification and Processing Systems, Inc. |

Task #54 was cancelled this session. Re-raise when ready.

---

## Replit environment note
The `artifacts/script-viewer` workflow was failing at end of session — unrelated
to anything done today. Restart it if the Script Viewer preview is blank.
