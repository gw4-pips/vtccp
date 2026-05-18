# DataMan 26.1.0 Release Notes — Digest

**Source**: `release-notes-26.1.0.txt` (filed alongside, verbatim from Cognex)
**Filed**: 2026-05-18, supplied while user was mid-install

---

## TOP-LINE FINDINGS

1. **26.1.0 is positioned as a refinement / supplement of 25.4.1**, not a
   major version. Language is "further refines and supplements." Suggests
   backward compatibility at the protocol and script level — our v1.23
   push script and `DmccClient`/`DataManSdkClient` should keep working.

2. **DM475V is NOT enumerated in any platform list in these notes.** This
   is the single most important observation for our install:
   - "General updates for fixed-mount platforms **280/80, 380/580, and
     290/390**" — DM475V absent
   - "General updates for platforms **280/80, 380/580, 290/390, and 8700**"
     — DM475V absent
   - **Firmware updates listed**: 6.3.10 (DM80/280), 6.2.9 SR3 (DM8700).
     **No update for our 6.1.x firmware family.**

   This means one of:
   - **(a) DM475V is in a separate verifier product line** not enumerated
     in these notes (the "V" suffix = verifier; the platforms enumerated
     are all readers/decoders, not verifiers). TruCheck-specific notes
     (see #3 below) suggest verifier features are still being actively
     maintained, which supports this reading.
   - **(b) DM475V firmware is in maintenance-only mode** — no new
     features, but compatibility with the new DMST tool is preserved.
   - **(c) DM475V is end-of-life / out of support.** Less likely given
     active TruCheck work.

   **Most plausible reading**: (a). Verifier line is tracked in separate
   notes. Watch for compatibility warnings during DMST 2026.1's first
   connection to the device; if it connects cleanly, we're fine.

3. **TruCheck (the verifier algorithm suite) is actively maintained in
   26.1.0**:
   - "Added ability to toggle the TruCheck graphics overlay."
   - "TruCheck-related data formatting tokens now output 'N/A' when
     TruCheck is disabled."
   - Improved GS1 Digital Link validation for both readers **and verifiers**.

   These confirm the verifier line is alive in 26.1.0. The "N/A token"
   change is **worth a regression check on our v1.23 push script** — if
   the script's data-formatting tokens (e.g. `{Trucheck.Overall}` style
   if any are in use) start emitting "N/A" under conditions where they
   previously emitted empty/null, our `DmstResultParser` may need
   updating. Our v1.23 builds the XML inside `onResult` directly from
   `r.*` properties, not via Format Data tokens, so we're likely
   immune — but worth confirming.

4. **New DMCC command: `DECODER.1D2D-USAGE`** for "Extended decode mode."
   Not directly relevant to verifier flows, but adds to the DMCC namespace
   our `DmccClient` could in principle send. Not enumerated in our 2025
   Comms guide digest — this is a 26.1.0 addition.

5. **Bugfix: "reader crashes when a DMCC command is sent during bootup"**
   — applies to the listed reader platforms. Not necessarily our DM475V
   (not enumerated), but if you ever script-trigger VTCCP startup very
   close to a device reboot, the bug class existed; the fix is in 26.1.0
   for those platforms.

6. **Bugfix: "FTP usage caused a small memory leak"** — applies to
   listed platforms. We don't use FTP push (we use TCP Network Client),
   so not relevant to our config.

7. **New WebUI feature: Diagnostics page** — software-based diagnostic
   tool, view input/output signals. Useful for troubleshooting if the
   2026.1 WebUI is enabled on our device (we connect via DMST, not
   WebUI, so usually moot).

8. **New WebUI feature: Services & Ports control** — can disable
   "major configuration APIs and channels for security purposes." **If
   someone enables this and disables the DMCC TCP service on port 23,
   VTCCP loses connectivity.** Worth documenting as a "do not disable"
   operational note.

9. **System requirements**: Win 10 (32 or 64) / Win 11 (64), .NET 4.7.2+.
   No mention of dropping Win 7/8/8.1 — consistent with 25.4.1's
   requirements as I recall.

---

## What this means for the install in progress

- **Best case**: 2026.1 connects cleanly to the DM475V on 6.1.16_sr4
  via the verifier-line compatibility path. No warnings. You proceed
  with the v1.23 push script untouched and capture a baseline scan
  for diff against your 2025-baseline XML.

- **Yellow flag**: 2026.1 connects but warns about firmware version.
  Record the exact wording for the migration-notes file. Likely
  workable but worth keeping 2025.4.1.1 installable as fallback.

- **Red flag**: 2026.1 refuses to connect or fails to read settings.
  Stop. Two options: (1) confirm 2025.4.1.1 is still installable and
  revert; (2) check for a separate "DataMan Verifier" release matching
  26.1.0 — verifier line may ship its own DMST build.

---

## Regression checks once 2026.1 is connected and reading the device

In rough priority order:

1. **Push pipeline end-to-end**: trigger a scan, confirm v1.23 XML
   arrives at VTCCP `DmstListener` with no shape changes. Same
   `<DMSymVerResponse>` envelope, same field set.

2. **`r.trucheck.jpegImage` accessibility**: probe in script —
   confirms the v1.24 image-emission path is alive on 2026.1.

3. **DPM-mode token behavior**: scan a calibration card (15415 mode);
   confirm that DPM-only fields (`cellDefects`, `finderPatternDefects`,
   etc.) still emit as empty/NA rather than the new "N/A" literal that
   the release notes mention for disabled-TruCheck tokens. **Our parser
   would mishandle the string "N/A" in a numeric field** — that's
   worth a v1.24 parser hardening pass either way.

4. **DMCC TRIGGER from `DmccClient`**: trigger via VTCCP's
   ⚡ Trigger Scan button (DMST disconnected). Confirm response codes
   match our existing 6/8/-1/-2/-3 mapping.

5. **`DataManSdkClient` connection lifecycle**: confirm long-lived
   SDK session still works for the push path.

---

## Action items derivable from these notes (no urgency)

- **(parser-hardening)** Make `DmstResultParser` treat the literal
  string `"N/A"` in numeric fields as null/missing rather than throwing
  parse error. Cheap defensive change against the new
  TruCheck-disabled-token behavior.
- **(operational doc)** Add a note to README about *not* disabling
  the DMCC service via the new WebUI Services & Ports control.
- **(comms digest update)** Add `DECODER.1D2D-USAGE` to the DMCC
  command catalog when we next refresh that digest.

None of these are blocking. All wait until you've completed the install
and shared baseline scan output.
