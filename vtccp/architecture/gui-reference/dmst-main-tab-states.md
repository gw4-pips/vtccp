# DMST TruCheck — Main Tab GUI States

**Device**: DM475-63530E-PIPS-Verif-Lab  
**Firmware**: fw 6.1.16_sr4 (confirmed, Last calibrated: 5/20/2026 1:14:58 AM)  
**Push script active**: v1.33

---

## State 1 — Idle / App Open

**Screenshot**: `dmst-main-idle.png`

The app has just been opened (or a result has been cleared). No scan has been
performed yet. The camera is not in live mode.

- Image pane: **live camera frame visible** — physical scene in view (symbol
  on table, environment background)
- Overall Grade: empty
- Format Grade: empty
- Symbology: empty
- Grade Parameters panel: empty
- Button visible: **Go Live**
- Acceptance Criteria / Data fields: empty
- Navigation counter: 1 / 1

---

## State 2 — Go Live (acquisition mode active)

**Screenshot**: `dmst-main-golive.png`

Operator clicked "Go Live". The camera is streaming a live acquisition feed.
DMST is attempting to locate and lock the symbol.

- Image pane: **live feed with locator overlay**
  - Red dashed bounding box drawn around the detected symbol
  - Red `+` crosshair marker below the symbol (focus reference point)
- Focus Adjustment widget visible (up/down arrow control)
- Overall Grade: still empty
- Grade Parameters panel: still empty
- Buttons visible: **Cancel Live Mode** (top) + **Verify** (bottom)
- DMST is waiting for operator to trigger Verify

---

## State 3 — Post-Scan Result (DM, Grade D)

**Screenshot**: `dmst-main-result-dm-grade-d.png`

Operator triggered Verify. Scan has completed and the result is displayed.

- Image pane: **GRAY / EMPTY** — no image displayed ← see note below
- Overall ISO15415/6 Grade: **1.0 (D)** — red background  
  Sub-line: **1.0/16/660/45Q**
- GS1 Format Grade: **FAIL** — red background
- Symbology: **DataMatrix**
- Buttons visible: **Cancel Live Mode** + **Go Live**
- Acceptance Criteria: **Fail (Quality)** — red background
- Data: `]><RS>06<GS>18VD89536<GS>1P8902A<GS>S3122A02965<RS><EO>`
- Grade Parameters: fully populated (UEC 41.7% / Grade 2.0 yellow, ANU 11.1% /
  Grade 1.0 red — two parameters pulling the overall to D)

### Grade parameters confirmed from this scan

| Parameter | % / Value | Grade |
|---|---|---|
| Unused Error Correction (UEC) | 41.7% | **2.0** (yellow) |
| Symbol Contrast (SC) | 78% | 4.0 (green) |
| Modulation (MOD) | — | 4.0 (green) |
| Reflectance Margin (RM) | — | **2.0** (yellow) |
| Axial Nonuniformity (ANU) | 11.1% | **1.0** (red) |
| Grid Nonuniformity (GNU) | 7.2% | 4.0 (green) |
| Fixed Pattern Damage (FPD) | — | 4.0 (green) |
| Left 'L' Side (LLS) | — | 4.0 |
| Bottom 'L' Side (BLS) | — | 4.0 |
| Left Quiet Zone (LQZ) | — | 4.0 |
| Bottom Quiet Zone (BQZ) | — | 4.0 |
| Upper Left Quiet Zone (ULQZ) | — | 4.0 |
| Upper Right Quiet Zone (URQZ) | — | 4.0 |
| Right Upper Quiet Zone (RUQZ) | — | 4.0 |
| Right Lower Quiet Zone (RLQZ) | — | 4.0 |
| Left Top Transition Ratio (LQTTR) | 0% | 4.0 |
| Right Top Transition Ratio (RQTTR) | 0% | 4.0 |
| Left Right Transition Ratio (LQRTR) | 0% | 4.0 |
| Right Right Transition Ratio (RQRTR) | 0% | 4.0 |
| Left Top Clock Track (LQTCT) | — | 4.0 |
| Right Top Clock Track (RQTCT) | — | 4.0 |
| Left Right Clock Track (LQRCT) | — | 4.0 |
| Right Right Clock Track (RQRCT) | — | 4.0 |
| Average Grade (AG) | — | 4.0 |

Overall = min(UEC=2, SC=4, MOD=4, RM=2, ANU=1, GNU=4, FPD=4) = **1.0 → Grade D**.

The 1.0 is driven entirely by **ANU = 1.0 (red, 11.1%)**. UEC and RM are 2.0 (yellow)
but are not the floor. This is the same symbol used for live-scan catalog entries #12
and #14 (DM GS1 16×36, consistently Grade D).

---

## Missing image in post-scan state — analysis

The image pane is gray in State 3. This is **standard DMST behavior**, not caused
by the VTCCP push script.

**Why the image disappears:**

DMST's Main tab image panel is a live camera feed display. It shows the camera output
while the device is in idle-preview or Go Live acquisition mode. Once the verification
cycle completes, the camera exits its live acquisition loop and the feed stops. DMST
does not replay the captured verification JPEG back into the Main tab image pane after
the scan — the result display mode replaces the feed with grade data.

The captured verification image IS present and accessible:
- In the push XML field `<JpegImageBase64>` (the `r.trucheck.jpegImage` value)
- In the DMST **Report** tab (HTML report, as the embedded scan image)
- In DMST-saved `.html` report files

**Push script confirmed not a factor:**
- v1.33 issues zero `IMAGE.LOAD`, `IMAGE.REPLAY`, or `IMAGE.SEND` commands
- `SYMBOL.RESULT.FULL` is not present in v1.33 (dropped after v1.25 campaign)
- All push script DMCC activity is read-only (GET operations on result data)

**Implication for VTCCP UI design:**

VTCCP should always display the captured symbol image after a scan — sourced from
`<JpegImageBase64>` in the push event. This is a concrete UX improvement over DMST:
the operator sees the exact image that was graded in the same panel as the grade result,
without needing to click to the Report tab. The D4 `JpegImage` field in `VerificationRecord`
will feed this display.

---

## Sub-line format: 1.0/16/660/45Q

The second line of the Overall Grade field: `1.0/16/660/45Q`

- `1.0` — numeric grade (same as letter grade floor)
- `16` — unknown (possibly scan number or aperture reference)
- `660` — likely illumination wavelength (660nm = red LED on DM475V-LBL)
- `45Q` — 45-degree quadrant illumination (the -LBL variant's fixed optics)

This format is not parsed by VTCCP. Logged for future reference.
