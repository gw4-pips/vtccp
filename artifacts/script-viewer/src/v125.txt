// ─────────────────────────────────────────────────────────────────────────────
// VTCCP DMST Push Script
//
//   Version   : 1.25
//   Generated : 2026-05-18 UTC
//   Source    : VTCCP Replit Agent  (github.com/gw4-pips/vtccp)
//   Target    : Cognex DataMan firmware 5.x / 6.x  /  DMV475
//
//   v1.25 — Confirmed wires: <JpegImageBase64> full JPEG payload (base64
//             string, ≈9.9 KB live DM / ≈22 KB loaded QR — confirmed safe
//             by v1.24 DebugJpegProbe), <OpticsSource> LoadedImage/LiveScan
//             (ContrastUniformity===-1 AND MRD===-1 discriminator, confirmed
//             across two loaded-image captures), ECI prefix strip from
//             <DecodedData> for QR codes with modifier 2/4/6, MatrixSize fix
//             for QR (QZ offset 8 not 2 — v3: sqrt(1369)=37, 37-8=29 ✓),
//             <ErrorCorrectionType> symbology-aware (QR → ecLevel probe /
//             fallback "QR"; DM → "ECC200"), <ApplicationPassReason> parsed
//             from ApplicationPass suffix (" (" split), <PixelsPerModule>
//             removed (always empty, redundant with <ModuleSizePx>).
//             New probes: DebugSymbols0 (q.symbols[0] full enum — target for
//             QR grade params ULP/URP/LLP/HCT/VCT/ALP/VIB/FIB + DataCodewords
//             + MaskPattern + ECLevel + per-region ULQZ/URQZ/RUQZ/RLQZ),
//             DebugPrintGrowth (m.printGrowth shape — diagnoses empty BWGPercent).
//             Fixed: DebugImageShape wrapped in try/catch (v1.24: no output,
//             silent error). Dropped: DebugANUCase (lowercase confirmed),
//             DebugJpegProbe (replaced), DebugReaderProps (answered).
//
//   v1.24 — Probe release: image-emission path opened; field promotions;
//             unknown-shape probes; case-mismatch resolver.
//             FOUNDATION: per the comms-and-programming-guide 25.4.1.1
//             digest (filed 2026-05-18), several assumptions in v1.23 are
//             formally confirmed and several new wires are unblocked:
//               • r.trucheck.jpegImage = Base64-encoded JPEG (canonical
//                 path to get the verified-symbol image into the push).
//               • r.trucheck.calibrationDate = string timestamp.
//               • r.symbology = {name, id, quality, moduleSize, corners,
//                                center, angle, size, failureCode}.
//               • r.image is METADATA only ({index, FoV, RoI, exposureTime,
//                 gain, illEnabled, id}) — image bytes live on r.trucheck.
//               • Case-mismatch note: comms guide spells axialNonUniformity
//                 (capital U); our v1.15 probe found lowercase 'u' worked on
//                 q-side. v1.24 probes both to nail the truth in scope.
//             v1.23 baseline (DMST 25.4.1.1 + 26.1.0): regression-passed.
//             Schema bit-identical across the DMST upgrade; pipeline is
//             firmware-agnostic. See references/manuals/cognex/DMST/
//             2026-release/release-notes-26.1.0-digest.md for evidence.
//             NEW WIRES (promote from confirmed-via-probe to first-class):
//               • <SymbologyId>     = r.symbology.id      (e.g. "]d1")
//               • <SymbolQuality>   = r.symbology.quality (0-100 decoder conf)
//               • <SymbolAngle>     = r.symbology.angle   (degrees)
//               • <ModuleSizePx>    = r.symbology.moduleSize (pixels/module)
//               • <CalibrationDate> = r.trucheck.calibrationDate (promote
//                                     from the CustomNote workaround)
//               • <FieldCalibrated>   = readerProperties.status3D.fieldCalibrated
//               • <FactoryCalibrated> = readerProperties.status3D.factoryCalibrated
//               • <BWGPercent>      = r.metrics.printGrowth.raw via mmPctAuto
//                                     (was empty in v1.23; comms guide
//                                     confirms printGrowth IS the BWG source).
//               • <MinPassGrade>    = r.metrics.minPassGrade.grade (new field
//                                     per session plan A5; v1.23 enumerated
//                                     minPassGrade=[obj] in DebugMetricsKeys).
//             NEW PROBES (the four unknowns + case resolver + reader-props):
//               • DebugImageShape  — full enum of r.image. Expecting:
//                                    index/FoV/RoI/exposureTime/gain/illEnabled/id
//                                    per comms guide. Also unblocks D4
//                                    (image-load full implementation): if any
//                                    extra property appears under loaded-image
//                                    mode, it'll show here.
//               • DebugJpegProbe   — r.trucheck.jpegImage LENGTH + first 80
//                                    chars. DOES NOT EMIT FULL PAYLOAD this
//                                    round — base64 JPEG could be 50-300 KB
//                                    which is 7-40× the current ~7 KB XML;
//                                    need size data before committing to full
//                                    emission to avoid Network Client buffer
//                                    overrun or listener-side parser stalls.
//                                    v1.25 commits to <JpegImageBase64> once
//                                    we confirm the network client + listener
//                                    handle the payload size cleanly.
//               • DebugGS1         — r.validation.gs1 deep enum (v1.23 saw
//                                    only the shallow [obj] marker).
//               • DebugDodUid      — r.validation.dodUid deep enum (ditto).
//               • DebugBarcodeAsgn — r.barcodeAssignment deep enum. Totally
//                                    unprobed sibling; comms guide doesn't
//                                    document it either. Pure discovery.
//               • DebugANUCase     — presence-marker for axialNonuniformity
//                                    (lower-u, working in v1.23) vs
//                                    axialNonUniformity (capital-U, comms-
//                                    guide spelling). Resolves the
//                                    documented vs empirical mismatch.
//               • DebugReaderProps — top-level enum of readerProperties.
//                                    Looking for status3D (calibration) and
//                                    any other extension props the guide
//                                    didn't enumerate.
//             RETAINED (continuing baselines, cheap):
//               • DebugModSize, DebugECCount       (formula sanity)
//               • DebugMetricsKeys, DebugRSiblings (regression watch — these
//                                                   stayed bit-identical
//                                                   across DMST 25→26.1)
//             DROPPED (questions answered by v1.23 XML — evidence in filed
//             live-scan sample):
//               • DebugSymbology    — answered: full key set known (becomes
//                                      wires above + DebugImageShape sibling)
//               • DebugCellDefects  — answered: raw=-1, grade=NA (DPM-only)
//               • DebugFPDefects    — answered: raw=-1, grade=NA (DPM-only)
//               • DebugDMCellDims   — answered: both W/H raw=-1, grade=NA
//               • DebugValidation   — answered: state=0, method=0; gs1/
//                                      dodUid get their own deep probes in
//                                      v1.24 replacing this shallow one.
//               • DebugMetricShape  — answered: r.metrics.symbolContrast =
//                                      {raw, grade} monolithic. ISO 15415
//                                      metrics in JS scope are flat single-
//                                      symbol grades, no hidden per-region
//                                      breakdown. Per-region grading
//                                      definitively requires a separate
//                                      data path (DMCC report engine).
//             DEFERRED to v1.25:
//               • <JpegImageBase64> full payload — pending size data.
//               • Per-region grading wire-ups — still architecturally
//                 absent in JS scope; revisit only if a DMCC GET command
//                 surfaces the structured-report data.
//
//   v1.23 — Final probe round (last chance for per-region data) + free wins.
//             v1.22 wins confirmed live:
//               • MatrixSize=16x36 on 16×36 rect scan (lookup table hit
//                 len=684 → "16x36" exactly matches PDF "16x36 (Data: 14x34)")
//               • All 8 per-region grade fields back to empty (false data gone)
//             v1.22 probes revealed:
//               • r.metrics has 30 ISO 15415 grade objects (UEC, SC, MOD, RM,
//                 ANU, GNU, FPD + 22 more) — all monolithic single-symbol
//                 grades.  NO per-region keys.  Confirms the firmware doesn't
//                 compute per-region ISO grades in JS scope.
//               • r has 12 unknown siblings: source (DEVICE NAME!),
//                 symbology[obj], image[obj], validation[obj], decoded/content/
//                 decodeTime/triggerTime/timeout/readSetup/annotation/label/
//                 custom_svg/barcodeAssignment.
//               • r-level row/col candidates all missed — dims aren't on r,
//                 likely inside r.symbology.
//             WIRE: <Source> = r.source (DM475 device name; matches PDF
//                   "Device Name" field).  Currently the only Phase 1 column
//                   we could populate but weren't.
//             FINAL PROBES (last per-region attempt before declaring scope
//             dead):
//               • DebugSymbology    — enum r.symbology (expect row/col dims).
//               • DebugCellDefects  — enum r.metrics.cellDefects (per-region?).
//               • DebugFPDefects    — enum r.metrics.finderPatternDefects
//                                     (finder pattern = L-finder + clock track
//                                     = exactly the per-region grading source).
//               • DebugDMCellDims   — enum r.metrics.dataMatrixCellWidth +
//                                     Height (single value vs per-region arr).
//               • DebugValidation   — enum r.validation (app-pass details).
//               • DebugMetricShape  — enum r.metrics.symbolContrast deeply
//                                     (does any standard metric have hidden
//                                     .regions[] / per-quadrant breakdown?).
//             Retained: DebugModSize, DebugECCount (formula sanity).
//             Dropped:  DebugRectDims (definitively miss at r-level).
//             Re-enabled compact: DebugMetricsKeys + DebugRSiblings
//                                 (useful baselines for any future scan).
//             If v1.23 probes also return no per-region data, v1.24 = strip
//             all probes, declare production rev, Phase 1 moves to C# parser/
//             Excel exporter with per-region columns blank for multi-region
//             symbols.  Per-region support deferred to a future phase using
//             a DMCC GET command for the structured report (or PDF parsing).
//
//   v1.22 — Per-region scope pivot + rectangular MatrixSize.
//             v1.21 proved a hard architectural limit: q.upperLeftPattern /
//             upperRightPattern / lowerLeftPattern / horizontalClockTrack /
//             verticalClockTrack / alignmentPatterns ALL return
//             {grade:"F", numericGrade:0} regardless of actual symbol grade.
//             They are inert placeholders.  The device's own PDF report for
//             the same 16×36 scan shows all 12 per-region grades populated
//             (ULQZ/URQZ/RUQZ/RLQZ + 4 TTR + 4 TCT all = A) — proving the
//             data is computed in a separate report engine the JS push
//             script cannot see.  The q.trucheck introspection path for
//             per-region data is dead-ended.
//             ROLLBACK: drop the misleading v1.21 ULQZ/URQZ/LLQZ/
//                       HClockTrackGrade/VClockTrackGrade wires that
//                       printed "F" everywhere.  Empty > false data.
//             DROP PROBES: DebugULP/URP/LLP/HCT/VCT/AlignPat/LRPSearch
//                          (all definitively dead — keys exist but never
//                          populate).
//             NEW PROBE 1: DebugMetricsKeys — full enum of r.metrics object.
//                          PushScriptDiag has long confirmed m=found but the
//                          metrics tree was never enumerated.
//             NEW PROBE 2: DebugRSiblings — every r.* key we haven't named.
//                          Looking for r.regions / r.report / r.dmsvReport /
//                          anything else that might carry per-region data.
//             NEW PROBE 3: DebugRectDims — try r.rowCount / r.columnCount /
//                          r.numRows / r.numColumns / r.symbolWidth /
//                          r.symbolHeight to find authoritative row×col
//                          dimensions (instead of inferring from modArray).
//             WIRE: rect-aware MatrixSize via modArray-length → ECC200
//                   rect-size lookup table:
//                       200=8x18 / 340=8x32 / 392=12x26 / 532=12x36 /
//                       684=16x36 / 900=16x48 (+ QZ on both axes).
//                   Square branch unchanged: side = sqrt(len) − 2.
//                   Scan 2 (len=684) should now emit "16x36" instead of "".
//             If v1.22's new probes also yield nothing useful, v1.23 will
//             either pivot to a DMCC GET command for the structured report
//             or ship Phase 1 with per-region columns empty for multi-region
//             symbols (the legacy top-level fields cover 1-region symbols
//             completely).
//
//   v1.21 — Per-region (≥32x32 / 2-region rectangular) support pass.
//             The v1.20 32x32 scan proved that top-level TQZ/RQZ/TCT/RCT/
//             TTR/RTR all flip to "X" (NA) on multi-region symbols — the
//             firmware moves that data into the per-region tree:
//                upperLeftPattern / upperRightPattern / lowerLeftPattern
//                horizontalClockTrack / verticalClockTrack / alignmentPatterns
//             Each *Pattern object is {grade, numericGrade} (confirmed v1.19).
//             WIRE 1: ULQZGrade / URQZGrade / LLQZGrade now sourced from
//                     upperLeftPattern.grade / upperRightPattern.grade /
//                     lowerLeftPattern.grade.  Preemptive — verified next scan.
//             WIRE 2: LRQZGrade left empty for now; firmware exposes only
//                     3 of 4 pattern slots.  Probe v1.21 tries fallback keys
//                     (lowerRightPattern, bottomRightPattern, alignmentPatterns
//                     index 3) to find where the 4th region lives, if at all.
//             PROBE: re-added pattern + clockTrack probes (dropped in v1.20)
//                    + new alignmentPatterns array probe + lowerRight key
//                    enumeration.  Retained DebugModSize / DebugECCount.
//                    Dropped DebugModSample (mystery solved — `grade` field
//                    on modulationArray entries is a single paren-coded
//                    modulation-bucket char, only relevant to Phase 2 PDF
//                    heatmap reproduction).
//           After v1.22 wires the per-region columns from probe data, v1.23
//           becomes the production rev and Phase 1 moves to C# parser/Excel.
//
//   v1.20 — Final wirings from v1.19 introspection — Phase 1 feature-complete:
//             FIX 1: reflectanceLight / reflectanceDark are PRIMITIVES (numbers),
//                    not objects with .raw — `_qRl["raw"]` returned undefined.
//                    Rebuilt SCRlRd to handle either shape.
//             WIRE 1: MatrixSize from sqrt(modulationArray.length) - 2.
//                     (22×22 ECC200 → 484 data + 2-cell quiet zone padding
//                     wrapper = 576 = 24² → 24-2 = 22.  Confirmed 22×22.)
//             WIRE 2: EncodedCharacters = encodationAnalysisArray.length.
//             WIRE 3: TotalCodewords = codewordArray.length.
//             WIRE 4: ErrorsCorrected = count of codewordArray[i].isCorrected.
//             WIRE 5: CustomNote = q.customNote (device calibration metadata).
//             COSMETIC: AGValue rounded to 1 decimal place (was 14-digit float).
//           Most v1.19 probes dropped — those territories are mapped.
//           Three retained for cross-validation: DebugModSize (matrix
//           computation sanity), DebugModSample (modulationArray[0] full
//           enumeration — to verify the "grade=(" oddity), DebugECCount
//           (errorsCorrected verification).
//           After v1.20 confirmation, v1.21 will strip all remaining probes
//           and become the production rev; Phase 1 then turns to wiring the
//           C# parser + Excel exporter.
//
//   v1.19 — Three bug fixes + four major new wirings from v1.18 introspection:
//             FIX 1: ISO branch matched "ISO 15415" exactly but device emits
//                    "ISO 15415:2011" — caused SCGrade/MODGrade to fall through
//                    to cellContrast/cellModulation (= "X").  Now uses indexOf.
//             FIX 2: MRD key is uppercase "MRD" on q.general, not "mrd".
//             FIX 3: distributedDamageGrade now emitted (new <DDGrade>).
//             WIRE 1: TTR/RTR now from q.topTransitionRatio /
//                     q.rightTransitionRatio (TrucheckMetric — has .raw + .grade)
//                     instead of being aliased to TCT/RCT.
//             WIRE 2: SC Rl/Rd now from q.reflectanceLight / q.reflectanceDark
//                     (not m.extremeReflectance/m.reflectMin which gave -1).
//             WIRE 3: AverageGrade emitted separately (q.averageGrade).
//             WIRE 4: MinimumReflectance emitted (q.minimumReflectance.raw).
//           New introspection probes for the unmapped territory: arrays
//           (codewordArray, encodationAnalysisArray, modulationArray,
//           applicationStdArray, asciiArray) and per-region patterns
//           (upperLeftPattern, upperRightPattern, lowerLeftPattern).
//           Removed v1.18 probes for q.overall and q.general — fully mapped.
//
//   v1.18 — BREAKTHROUGH: Cognex's official CSV-results template revealed that
//           r.trucheck has NESTED sub-objects we never traversed:
//             trucheck.overall.{gradingStandard,gradeLetter,gradeValue,
//                               applicationStandardName,applicationStandardPass}
//             trucheck.general.{xDimension,contrastUniformity,
//                               horizontalBWG,verticalBWG,...}
//             trucheck.<param>.raw          ← TRUE ISO percentages
//                                             (q-side, NOT r.metrics)
//           Previously we read percentages from r.metrics.<param>.raw which
//           are firmware-internal half-cooked values (SC came out 71/82.7
//           instead of the correct 79).  Switching to q.<param>.raw fixes
//           SC, ANU, GNU, FPD percentages; q.overall fixes overall/formal
//           grades; q.general fixes BWG, X-Dim, contrast uniformity.
//           DPM-vs-ISO branch added (cellContrast/cellModulation when
//           gradingStandard != "ISO 15415").
//           Introspection probes for q.overall and q.general added to
//           discover remaining keys (aperture, wavelength, lighting,
//           matrixSize, encodedChars, codewords, MRD).
//           DMCC GET path abandoned (firmware exposes no GET-able
//           verification namespace — DMSV.* returns [101] invalid).
//
//   v1.17 — Production cleanup.  All debug probes removed (DebugMetric_*).
//
//   v1.16 — UEC scaling fix (10000 → 100) via new mmPctAuto helper.
//           ANUPercent/GNUPercent/HBW/VBW switched to mmPctAuto.
//           MatrixSize, TTR/RTR wired to dataMatrixCell* and markMisplacement
//           (later found to be -1 sentinel — see v1.17 note).
//
//   v1.1 — Fix: use outputResults.content (firmware 6.x parameter) with
//           fallback to global output (firmware 5.x) for cross-version
//           compatibility.  Previously the script crashed with
//           "ReferenceError: output is not defined" on firmware 6.x,
//           causing the device to fall back to plain-text Basic Formatting.
//
//   v1.2 — Fix: firmware 6.x exposes symbology as a typed object, not a
//           plain string; added multi-path fallback (symbologyName →
//           symbology string → symbology.name → String(symbology)).
//           Fix: firmware 6.x may expose quality data at a path other than
//           r.quality; added fallback chain (quality → verificationResult →
//           symbolVerificationResult → symVerResult).
//
//   v1.3 — Fix: removed Object.keys() calls introduced in v1.2; the
//           firmware's embedded JS engine does not support Object.keys(),
//           causing a ReferenceError that crashed the script before
//           outputResults.content was set, reverting every scan to plain-
//           text Basic Formatting mode.  Replaced with a typeof probe loop
//           over a known property list (ES3-compatible).
//
//   v1.4 — Fix: switched quality-path lookup from truthy || chain to
//           explicit _pick() with !== null/undefined, so a quality object
//           that evaluates as 0/false is not skipped.  Added 5 more quality
//           path candidates (verificationResults, qualityResult, gradeResult,
//           rp.quality, rp.verificationResult).  Added _DbgRPKeys probe on
//           readerProperties.  Expanded _rProbe list.
//
//   v1.5 — Cleanup: removed all _Dbg* firmware introspection elements now
//           that quality-path mapping is confirmed on DM475 fw 6.1.16_sr4.
//           (NOTE: turned out to be premature — quality path still not found.)
//
//   v1.6 — Diagnostics re-added: quality-path has never been confirmed on
//           DM475 fw 6.1.16_sr4 (all scans arrive with empty grade columns).
//           Added <PushScriptDiag> element reporting which quality source
//           was resolved, plus 10 additional property-path candidates on
//           decodeResults[0] and readerProperties.  Also probe decodeResults
//           itself (array-level) and try nested sub-paths (.result, .data,
//           .isoResult).  Confirmed ES3-compatible (no Object.keys, no
//           const/let, no arrow functions, no template literals).
//
//   v1.7 — Diagnostics: all 34 quality-path probes returned null on DM475
//           fw 6.1.16_sr4 (r exists as object, r.decoded=true, but no quality
//           property found).  This version adds a full property-presence scan
//           using typeof on 80 candidate names on both r and rp, reporting the
//           defined ones in <DebugRFound> / <DebugRPFound> XML elements.
//           This will identify the exact property name used by this firmware.
//
//   v1.8 — Fix: property scan revealed quality object is at r.trucheck
//           (all lowercase — all previous probes used camelCase variants).
//           Added "trucheck" and "metrics" to _rCandidates so q resolves.
//           Added <DebugQFound> inner-property scan on the resolved q object
//           to identify the exact names for overallGrade, uec, sc, etc.
//
//   v1.9 — Diagnostics: <DebugQFound> revealed only "modulation" and
//           "decode" because the scan used abbreviations (uec, sc, mod, …)
//           while Cognex DMST scripting uses full English camelCase names
//           (uniformEdgeContrast, symbolContrast, modulation, …).
//           Also both properties resolved to [object TrucheckMetric] —
//           every grade lives one level deeper inside those sub-objects.
//           This version:
//           1. Replaces _qScanNames with full Cognex camelCase property
//              names to find all TrucheckMetric entries on r.trucheck.
//           2. Adds <DebugModFound> inner-property scan on q.modulation
//              to reveal the exact field names inside a TrucheckMetric
//              (e.g. grade/value, letter/numeric, gradeValue/percent, …).
//           Once TrucheckMetric structure is known, all 167 grade columns
//           can be wired correctly.
//
//   v1.15 — FIX: for-in scan (v1.14) revealed correct property names on q.
//           All 18 DataMatrix parameters are now accessible; we were simply
//           using wrong camelCase for 5 of them.  Fixed q-side bindings:
//             unusedErrorCorrection → UEC grade  (was 'uniformEdgeContrast')
//             axialNonuniformity    → ANU grade  (was 'axialNonUniformity' cap-U)
//             gridNonuniformity     → GNU grade  (was 'gridNonUniformity' cap-U)
//             leftLSide             → LLS grade  (was 'leftL'/'lls')
//             bottomLSide           → BLS grade  (was 'bottomL'/'bls')
//           Fixed m-side bindings (r.metrics):
//             'UEC'                 → UEC%       (was 'uniformEdgeContrast')
//             extremeReflectance    → Rl (max reflectance 0–1)
//             reflectMin            → Rd (min reflectance 0–1)
//           True ISO SC% = (Rl − Rd) × 100 now emitted (gives 71% vs 72.5%)
//           SCRlRd built as "Rl_pct/Rd_pct" (e.g. "75/4")
//           HorizontalBWG / VerticalBWG wired to horizontalMarkGrowth /
//             verticalMarkGrowth (raw unit TBD; emitting mmVal for now).
//           DebugQEnum removed (q property list fully known).
//           DebugMEnum kept + DebugMEnum2 added (skip first 16) to see
//             truncated tail — looking for TTR%/RTR% Metric objects.
//
//   v1.14 — PROBE: Name-based probes cannot find LLS/BLS grades or TTR%/RTR%.
//           Switch to for-in enumeration of both r.trucheck and r.metrics so
//           every property is visible regardless of name.  Also enumerate
//           r.content and r.image for dimension data.  Targeted fix:
//           SCPercent uses mmPct which returns 72.5% but device shows 71%
//           (device uses integer pixel math; ours is float ratio); tolerable
//           until a better source is found.
//
//   v1.13 — CORRECTION: v1.12 incorrectly suppressed "NA" in mmGrade().
//           OverallGrade, ANU, and GNU ARE measured and reported by the
//           DM475 for this standard — "NA" is a valid device-returned grade
//           when a parameter could not be computed for a specific scan.
//           mmGrade() now returns the grade string as-is (including "NA").
//           mmVal()/mmPct() continue to suppress -1 (numeric sentinel;
//           -1 × 100 = -100% would be meaningless in the XLSX).
//           SCPercent raw = 0.7254... is a 0–1 ratio; ×100 → 72.5%.
//           UEC confirmed absent; symbol dimension fields not in DMST API.
//
//   v1.12 — CLEANUP: Metric confirmed as .grade + .raw; added mmPct();
//           SCPercent/ANUPercent/GNUPercent switched to mmPct (×100);
//           all diagnostic emit blocks removed (paths confirmed v1.11).
//
//   v1.11 — FIX: v1.10 revealed r.metrics is also a nested-object container
//           — each named property is a [object Metric] sub-object, exactly
//           mirroring the r.trucheck / TrucheckMetric pattern.  Confirmed
//           Metric properties: overallGrade, symbolContrast, modulation,
//           reflectanceMargin, axialNonUniformity, gridNonUniformity,
//           fixedPatternDamage, printGrowth, contrastUniformity.
//           Adds mmVal(metric)/mmGrade(metric) helpers that cascade through
//           likely sub-property names (value, percent, grade, measurement…)
//           to extract the scalar from any Metric object.  Adds inner probe
//           <DebugMetricOGFound> on m.overallGrade to confirm exact names.
//           Fixes all [object Metric] outputs for OverallGrade, SCPercent,
//           ANUPercent/Grade, GNUPercent/Grade, AGValue, ContrastUniformity.
//           UEC (uniformEdgeContrast) absent from both r.trucheck and
//           r.metrics — likely not measured for this standard/config.
//
//   v1.10 — WIRING: v1.9 confirmed all 12 TrucheckMetric properties on
//           r.trucheck and that every TrucheckMetric has exactly .grade
//           (letter) and .numericGrade (integer). Adds tmGrade/tmNum
//           helper functions.  Wires all 12 confirmed grade columns:
//             symbolContrast → SCGrade
//             modulation     → MODGrade
//             reflectanceMargin → RMGrade
//             fixedPatternDamage → FPDGrade
//             decode         → DecodeGrade
//             printGrowth    → AGGrade
//             leftQuietZone  → LQZGrade
//             bottomQuietZone→ BQZGrade
//             rightQuietZone → RQZGrade
//             topQuietZone   → TQZGrade
//             topClockTrack  → TCTGrade (also used for TTRGrade as alias)
//             rightClockTrack→ RCTGrade (also used for RTRGrade as alias)
//           Missing: UEC, ANU, GNU, overallGrade, all % values, dimensions.
//           These are expected on r.metrics. Adds <DebugMetricsFound>
//           probe on r.metrics to identify its exact property names.
//           Removes DebugRFound/DebugRPFound (those questions are answered).
// ─────────────────────────────────────────────────────────────────────────────
//
// HOW TO INSTALL
//   DataMan Setup Tool → Format Data (click it in the Application Steps sidebar)
//   1. On the BASIC tab: select the "Script-Based Formatting" radio button
//      (it is below the "Basic Formatting" radio button — DMST will warn you
//       that scripting is being enabled; confirm/OK any prompt)
//   2. Click the SCRIPTING tab (top of the Format Data panel)
//   3. Paste this entire script into the editor pane on the Scripting tab
//   4. Click Save Settings → Write to device
//
// WHAT IT DOES
//   After every scan the device calls onResult().  This script builds a
//   <DMCCResponse><DMSymVerResponse>…</DMSymVerResponse></DMCCResponse> XML
//   document and assigns it to output.content.  That content is then pushed
//   over the Network Client TCP connection to VTCCP (10.10.10.19:9004).
//   VTCCP's DmstListener detects the closing </DMCCResponse> tag, parses the
//   XML, maps every element to a VerificationRecord, and writes a fully-
//   populated XLSX row.
//
// TROUBLESHOOTING
//   If a column stays blank in VTCCP: the property name on the right-hand side
//   of the corresponding elem() call returned undefined.  Cognex exposes
//   slightly different property names across firmware revisions.  The property
//   names are annotated below; check the DataMan Scripting API Reference
//   (Settings → Help in DMST) for your exact firmware if a field is empty.
//
// COMPATIBILITY
//   ECMAScript 5 only — no const/let, no arrow functions, no template literals.
// ─────────────────────────────────────────────────────────────────────────────

function onResult(decodeResults, readerProperties, outputResults) {

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Safely stringify a value; returns "" for undefined/null.
    function s(v) {
        return (v === undefined || v === null) ? "" : String(v);
    }

    // Safely read a nested property path, e.g. prop(q, "lqz") → q.lqz ?? "".
    function prop(obj, key) {
        if (!obj) return "";
        var v = obj[key];
        return (v === undefined || v === null) ? "" : String(v);
    }

    // XML-safe string — replaces the five XML special characters.
    function esc(v) {
        return s(v)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

    // Emit one XML element.  Empty value → element present but empty.
    function elem(tag, val) {
        return "<" + tag + ">" + esc(val) + "</" + tag + ">\r\n";
    }

    // ISO 8601 timestamp from current wall-clock (device local time).
    function isoNow() {
        var t = new Date();
        function p(n) { return n < 10 ? "0" + n : String(n); }
        return t.getFullYear() + "-" + p(t.getMonth() + 1) + "-" + p(t.getDate())
             + "T" + p(t.getHours()) + ":" + p(t.getMinutes()) + ":" + p(t.getSeconds());
    }

    // ── Inputs ────────────────────────────────────────────────────────────────

    var r  = decodeResults[0];
    var rp = readerProperties;   // some firmware puts quality here instead of r

    // Null-safe property fetch — returns the value or null; never throws.
    // Using !== null/undefined (not truthy) so a quality object that happens
    // to stringify as 0 / false is still captured.
    function _pick(obj, key) {
        if (!obj) { return null; }
        var v = obj[key];
        return (typeof v !== "undefined" && v !== null) ? v : null;
    }

    // TrucheckMetric helpers (v1.10) — every sub-property on r.trucheck is a
    // TrucheckMetric with exactly two fields: .grade (letter) and .numericGrade.
    function tmGrade(tm) { return (tm && typeof tm["grade"]        !== "undefined") ? s(tm["grade"])        : ""; }
    function tmNum(tm)   { return (tm && typeof tm["numericGrade"] !== "undefined") ? s(tm["numericGrade"]) : ""; }

    // Metric helpers (v1.13) — Metric objects confirmed to have exactly:
    //   .grade  — letter string ("A"–"F") or "NA" (valid device output)
    //   .raw    — measurement number, or -1 (numeric sentinel = no value)
    // mmGrade() returns grade as-is — "NA" is a valid reported grade and
    //           is NOT suppressed (device genuinely reports it per scan).
    // mmVal()   returns raw as string; suppresses -1 (sentinel) → "".
    // mmPct()   returns raw×100, 1 dp; suppresses -1 → "" (-100% invalid).
    //           Use mmPct for percentage columns (SCPercent, ANUPercent…).
    function mmGrade(met) {
        if (!met) { return ""; }
        var _v = met["grade"];
        if (typeof _v === "undefined" || _v === null) { return ""; }
        return s(_v);  // "NA" is a valid device-returned grade — pass through
    }
    function mmVal(met) {
        if (!met) { return ""; }
        var _v = met["raw"];
        if (typeof _v === "undefined" || _v === null) { return ""; }
        return (_v === -1 || _v === "-1") ? "" : s(_v);
    }
    function mmPct(met) {
        if (!met) { return ""; }
        var _v = met["raw"];
        if (typeof _v === "undefined" || _v === null || _v === -1) { return ""; }
        var _n = parseFloat(_v);
        if (isNaN(_n)) { return ""; }
        // Raw is 0–1 ratio; convert to percentage with 1 decimal place.
        var _pct = Math.round(_n * 1000) / 10;
        return s(_pct);
    }
    // mmPctAuto (v1.16) — UEC.raw arrives already as 0–100 (giving 10000 with mmPct);
    // others (symbolContrast, ANU, GNU) arrive as 0–1 ratio.  Auto-detect: if the
    // raw value is > 1, assume it's already a percent; otherwise treat as ratio.
    function mmPctAuto(met) {
        if (!met) { return ""; }
        var _v = met["raw"];
        if (typeof _v === "undefined" || _v === null || _v === -1) { return ""; }
        var _n = parseFloat(_v);
        if (isNaN(_n)) { return ""; }
        var _pct = (_n > 1) ? _n : (_n * 100);
        return s(Math.round(_pct * 10) / 10);
    }

    // ── Quality-object discovery ──────────────────────────────────────────────
    // Try every known property path in priority order.  _qSource records which
    // one succeeded so <PushScriptDiag> can report it.  If every probe returns
    // null the <PushScriptDiag> element will say "none" and all grade columns
    // will be empty — paste the VS Output [VTCCP-DMST] RawXML line to the
    // VTCCP agent so it can identify the correct path.

    var q       = null;
    var _qSrc   = "none";

    // --- decodeResults[0] first-level candidates ---
    var _rCandidates = [
        "trucheck",             // DM475 fw 6.1.16_sr4 confirmed (v1.8)
        "metrics",              // also present on r — may carry measurement data
        "quality",
        "verificationResult",
        "symbolVerificationResult",
        "symVerResult",
        "verificationResults",
        "qualityResult",
        "gradeResult",
        "isoResult",
        "truCheckResult",
        "verResult",
        "isoVerResult",
        "verificationData",
        "gradeData",
        "verification",
        "grade",
        "gradeInfo",
        "result",
        "data"
    ];

    for (var _ri = 0; _ri < _rCandidates.length; _ri++) {
        var _v = _pick(r, _rCandidates[_ri]);
        if (_v !== null) {
            q     = _v;
            _qSrc = "r." + _rCandidates[_ri];
            break;
        }
    }

    // --- readerProperties candidates (only if r-level search failed) ---
    if (q === null) {
        var _rpCandidates = [
            "quality",
            "verificationResult",
            "symbolVerificationResult",
            "symVerResult",
            "qualityResult",
            "verificationData",
            "truCheckResult",
            "isoResult",
            "gradeResult",
            "verResult",
            "gradeData",
            "verification"
        ];
        for (var _rpi = 0; _rpi < _rpCandidates.length; _rpi++) {
            var _pv = _pick(rp, _rpCandidates[_rpi]);
            if (_pv !== null) {
                q     = _pv;
                _qSrc = "rp." + _rpCandidates[_rpi];
                break;
            }
        }
    }

    // --- decodeResults array-level (some firmware puts quality on the array) ---
    if (q === null) {
        var _arrCandidates = [
            "quality",
            "verificationResult",
            "qualityResult",
            "gradeResult"
        ];
        for (var _ai = 0; _ai < _arrCandidates.length; _ai++) {
            var _av = _pick(decodeResults, _arrCandidates[_ai]);
            if (_av !== null) {
                q     = _av;
                _qSrc = "decodeResults." + _arrCandidates[_ai];
                break;
            }
        }
    }

    // --- nested sub-paths on r.result and r.data (last resort) ---
    if (q === null && _pick(r, "result") !== null) {
        var _rResult = _pick(r, "result");
        var _subKeys = ["quality", "verificationResult", "gradeResult", "isoResult"];
        for (var _si = 0; _si < _subKeys.length; _si++) {
            var _sv = _pick(_rResult, _subKeys[_si]);
            if (_sv !== null) {
                q     = _sv;
                _qSrc = "r.result." + _subKeys[_si];
                break;
            }
        }
    }

    // ── XML assembly ──────────────────────────────────────────────────────────

    var o = '<?xml version="1.0" encoding="UTF-8"?>\r\n'
          + '<DMCCResponse>\r\n'
          + '<DMSymVerResponse>\r\n';

    // ── Identity / timing ─────────────────────────────────────────────────────
    //   r.decoded        — true / false
    //   r.content        — decoded string (empty when NoRead)
    //   r.symbologyName  — plain string in firmware 6.x  (preferred)
    //   r.symbology      — plain string in firmware 5.x, typed object in 6.x

    o += elem("DateTime",    isoNow());

    // Firmware-compatibility: r.symbology may be a typed object (6.x) or a
    // plain string (5.x).  Try property variants before falling back to String().
    var _symbStr = "";
    if (r) {
        if      (typeof r.symbologyName === "string" && r.symbologyName) { _symbStr = r.symbologyName; }
        else if (typeof r.symbology     === "string" && r.symbology)     { _symbStr = r.symbology; }
        else if (r.symbology && r.symbology.name)                        { _symbStr = String(r.symbology.name); }
        else                                                             { _symbStr = s(r.symbology); }
    }
    o += elem("SymbologyName", _symbStr);

    // ── v1.24+: r.symbology fields (moved before DecodedData for v1.25 ECI strip)
    //   Comms guide §4 documents: r.symbology = {name, id, quality, moduleSize,
    //   corners, center, angle, size, failureCode}. v1.23 DebugSymbology
    //   confirmed all populated on DM475 fw 6.1.16_sr4 (id="]d1" = AIM
    //   Symbology ID for Data Matrix ECC200; quality=41 (0-100 decoder conf);
    //   moduleSize=16.18 px; angle=1° (essentially horizontal placement)).
    var _sym   = (r && r.symbology && typeof r.symbology === "object") ? r.symbology : null;
    var _aimId = _sym ? s(_sym["id"]) : "";

    // v1.25: QR branch flag — SymbologyId[1]==='Q' for any QR Code Model 2 variant.
    //   Used for ECI strip, MatrixSize QZ offset, and ErrorCorrectionType below.
    var _qrBranch = (_aimId.length >= 2 && _aimId.charAt(1) === "Q");

    // v1.25: Strip ECI prefix from DecodedData for QR with ECI present.
    //   AIM modifier digit: '2','4','6' = ECI present → strip \x00\x00\x26 (3 bytes).
    //   '1','3','5' = no ECI → pass through as-is.
    //   ECI 26 = UTF-8. Bytes: 0x00 0x00 0x26 = charCode 0, 0, 38.
    //   Ref: references/standards/AIM-identifiers-honeywell-KB000063426.md
    var _decodedStr = (r && r.decoded) ? (r.content || "") : "";
    if (_qrBranch && _aimId.length >= 3) {
        var _aimMod = _aimId.charAt(2);
        if (_aimMod === "2" || _aimMod === "4" || _aimMod === "6") {
            if (_decodedStr.length >= 3
                    && _decodedStr.charCodeAt(0) === 0
                    && _decodedStr.charCodeAt(1) === 0
                    && _decodedStr.charCodeAt(2) === 38) {
                _decodedStr = _decodedStr.substring(3);
            }
        }
    }
    o += elem("DecodedData", esc(_decodedStr));

    o += elem("SymbologyId",   _sym ? s(_sym["id"])         : "");
    o += elem("SymbolQuality", _sym ? s(_sym["quality"])    : "");
    o += elem("SymbolAngle",   _sym ? s(_sym["angle"])      : "");
    o += elem("ModuleSizePx",  _sym ? s(_sym["moduleSize"]) : "");

    // ── Diagnostic elements (v1.7) ───────────────────────────────────────────
    // Reports which quality-object path resolved (or "none"), plus a comma-
    // separated list of every property name that is defined on r and rp.
    // Visible in VTCCP VS Output as [VTCCP-DMST] RawXML.
    // Remove this section once the correct property names are confirmed.

    // r.metrics — sibling of r.trucheck on decodeResults[0].
    // Expected to carry: overall grade, UEC/ANU/GNU, SC%/MOD%/RM%, dimensions.
    var m = _pick(r, "metrics");

    o += elem("PushScriptDiag", "v1.25 q=" + _qSrc + " m=" + (m ? "found" : "null"));
    // v1.23: device name from r.source (matches PDF "Device Name" field;
    //        scan confirmed value "DM475-63530E-PIPS-Verif-Lab").
    o += elem("Source",         (r && typeof r.source === "string") ? esc(r.source) : "");

    // ── Grade emission (v1.10) ────────────────────────────────────────────────
    //
    //   Data sources:
    //     q  = r.trucheck  — TrucheckResult; each sub-property is a TrucheckMetric
    //                        with exactly .grade (letter) and .numericGrade (int)
    //     m  = r.metrics   — sibling object; expected to carry % values, overall
    //                        grade, verification conditions, and symbol dimensions
    //
    //   Confirmed TrucheckMetric properties on q (v1.14 for-in scan) — 18 params:
    //     unusedErrorCorrection (UEC), symbolContrast, modulation (TrucheckMetricGrade),
    //     reflectanceMargin (TrucheckMetricGrade), axialNonuniformity, gridNonuniformity,
    //     fixedPatternDamage, decode, printGrowth, leftLSide (TrucheckMetricGrade),
    //     bottomLSide (TrucheckMetricGrade), leftQuietZone (TrucheckMetricGrade),
    //     bottomQuietZone (TrucheckMetricGrade), topQuietZone, rightQuietZone,
    //     topClockTrack, rightClockTrack
    //     Also: UII, batch, calibrationDate (device/label metadata, not quality grades)
    //   CAMELCASE TRAPS: q uses lowercase 'u' (axialNon*u*niformity),
    //                    m uses capital 'U' (axialNon*U*niformity).
    //   Confirmed Metric properties on m (v1.14 scan, page 1):
    //     symbolContrast, cellContrast, axialNonUniformity, printGrowth, UEC,
    //     modulation, fixedPatternDamage, gridNonUniformity, extremeReflectance (Rl),
    //     reflectMin (Rd), edgeContrastMin, singleScanInt, multiScanInt,
    //     signalToNoiseRatio, horizontalMarkGrowth (HBW), verticalMarkGrowth (VBW)
    //     Page 2 (v1.15 scan) — TTR%/RTR% Metric objects TBD.

    if (q) {

        // ── v1.18 NESTED SUB-OBJECTS (per Cognex CSV template) ────────────────
        //   q.overall   → grading standard, overall/formal grade, app standard
        //   q.general   → x-dimension, contrast uniformity, horizontal/vertical BWG,
        //                 plus probably matrix size, codewords, MRD (probed below)
        var qOv = _pick(q, "overall");
        var qGn = _pick(q, "general");

        // Helper: pull any sub-object property as string (handles undefined/null)
        function ovProp(key) { return qOv ? prop(qOv, key) : ""; }
        function gnProp(key) { return qGn ? prop(qGn, key) : ""; }

        // ── q (r.trucheck) TrucheckMetric bindings — confirmed names from v1.14 scan
        var _uec = _pick(q, "unusedErrorCorrection");
        var _sc  = _pick(q, "symbolContrast");
        var _cc  = _pick(q, "cellContrast");           // v1.18: DPM branch
        var _mod = _pick(q, "modulation");
        var _cm  = _pick(q, "cellModulation");         // v1.18: DPM branch
        var _rm  = _pick(q, "reflectanceMargin");
        var _anu = _pick(q, "axialNonuniformity");     // lowercase 'u' on q-side
        var _gnu = _pick(q, "gridNonuniformity");      // lowercase 'u' on q-side
        var _fpd = _pick(q, "fixedPatternDamage");
        var _lls = _pick(q, "leftLSide");
        var _bls = _pick(q, "bottomLSide");
        var _lqz = _pick(q, "leftQuietZone");
        var _bqz = _pick(q, "bottomQuietZone");
        var _rqz = _pick(q, "rightQuietZone");
        var _tqz = _pick(q, "topQuietZone");
        var _tct = _pick(q, "topClockTrack");
        var _rct = _pick(q, "rightClockTrack");
        var _dec = _pick(q, "decode");
        var _ag  = _pick(q, "printGrowth");

        // ── Grading standard / DPM branch ─────────────────────────────────────
        //   Per Cognex template: when overall.gradingStandard != "ISO 15415"
        //   the contrast/modulation parameters come from cellContrast /
        //   cellModulation instead of symbolContrast / modulation.
        var _gradeStd = ovProp("gradingStandard");
        //   v1.19: device emits "ISO 15415:2011" (with suffix); use prefix match.
        var _isIso    = (_gradeStd.indexOf("ISO 15415") === 0);
        var _scSrc    = _isIso ? _sc  : _cc;   // for SCGrade + SCPercent
        var _modSrc   = _isIso ? _mod : _cm;   // for MODGrade

        // ── Grading summary (v1.18: from q.overall) ───────────────────────────
        var _ogLetter  = ovProp("gradeLetter");
        var _ogNumeric = ovProp("gradeValue");
        o += elem("FormalGrade",         _ogLetter ? (_ogNumeric + "/" + _ogLetter) : "");
        o += elem("OverallGrade",        _ogLetter);
        o += elem("OverallGradeNumeric", _ogNumeric);
        o += elem("GradingStandard",     _gradeStd);
        o += elem("ApplicationStandard", ovProp("applicationStandardName"));
        o += elem("ApplicationPass",     ovProp("applicationStandardPass"));
        // v1.25: Parse reason suffix for parser routing.
        //   "Pass" → reason="".  "Fail (Quality)" → reason="Quality".
        //   "Fail (X Dimension out of Range)" → reason="X Dimension out of Range".
        //   (Format) failures appear when Data Format Check ≠ None.
        //   Parser must route (Quality) to grade summary; (Format) to app-syntax validator.
        var _apStr    = ovProp("applicationStandardPass");
        var _apIdx    = _apStr.indexOf(" (");
        var _apReason = (_apIdx !== -1 && _apStr.charAt(_apStr.length - 1) === ")")
                        ? _apStr.substring(_apIdx + 2, _apStr.length - 1) : "";
        o += elem("ApplicationPassReason", _apReason);

        // ── Verification conditions ───────────────────────────────────────────
        //   Names TBD — probes below enumerate qOv to reveal actual keys.
        //   Trying the most likely candidates from PDF report layout.
        o += elem("ApertureRef", ovProp("aperture")   || ovProp("apertureRef")  || prop(m, "aperture"));
        o += elem("Wavelength",  ovProp("wavelength") || ovProp("waveLength")   || prop(m, "wavelength"));
        o += elem("Lighting",    ovProp("lighting")   || ovProp("lightingType") || prop(m, "lighting"));
        o += elem("Standard",    _gradeStd            || prop(m, "standard"));

        // ── 2D ISO 15415 quality parameters (v1.18: from q.<param>.raw) ───────
        //   q-side TrucheckMetric has BOTH .grade AND .raw (confirmed by
        //   Cognex template using decodeResults[0].trucheck.axialNonuniformity.raw).
        //   mmVal / mmPctAuto work on any object with .raw — reusable.
        o += elem("UECPercent", mmPctAuto(_uec));
        o += elem("UECGrade",   tmGrade(_uec));

        //   SC — TRUE ISO SC% from q.symbolContrast.raw
        //   Rl/Rd from q.reflectanceLight / q.reflectanceDark
        //   v1.20: device emits these as PRIMITIVE NUMBERS, not objects with .raw.
        //   Accept either shape: if object → use .raw; if number → use directly.
        function _refNum(_x) {
            if (_x === null || typeof _x === "undefined") { return NaN; }
            if (typeof _x === "number") { return _x; }
            if (typeof _x === "object" && typeof _x["raw"] !== "undefined") {
                return parseFloat(_x["raw"]);
            }
            return parseFloat(_x);
        }
        function _refToPct(_v) { return (_v > 1) ? _v : (_v * 100); }
        var _rlRaw  = _refNum(_pick(q, "reflectanceLight"));
        var _rdRaw  = _refNum(_pick(q, "reflectanceDark"));
        var _rlInt  = (!isNaN(_rlRaw) && _rlRaw !== -1) ? String(Math.round(_refToPct(_rlRaw))) : "";
        var _rdInt  = (!isNaN(_rdRaw) && _rdRaw !== -1) ? String(Math.round(_refToPct(_rdRaw))) : "";
        var _scRlRd = (_rlInt && _rdInt) ? (_rlInt + "/" + _rdInt) : "";
        o += elem("SCPercent",  mmPctAuto(_scSrc));
        o += elem("SCRlRd",     _scRlRd);
        o += elem("SCGrade",    tmGrade(_scSrc));
        //   MinReflectance: firmware returns raw=0/grade=F on most scans —
        //   suppress when grade is F+raw=0 (firmware NA sentinel).
        var _minR    = _pick(q, "minimumReflectance");
        var _minRStr = "";
        if (_minR) {
            var _mrG = (typeof _minR["grade"] !== "undefined") ? String(_minR["grade"]) : "";
            var _mrR = (typeof _minR["raw"]   !== "undefined") ? parseFloat(_minR["raw"]) : NaN;
            if (!(_mrG === "F" && _mrR === 0)) {
                _minRStr = mmPctAuto(_minR);
            }
        }
        o += elem("MinReflectance", _minRStr);

        o += elem("MODGrade",   tmGrade(_modSrc));
        o += elem("RMGrade",    tmGrade(_rm));

        //   ANU / GNU — percentages from q.<param>.raw (TRUE ISO values)
        o += elem("ANUPercent", mmPctAuto(_anu));
        o += elem("ANUGrade",   tmGrade(_anu));
        o += elem("GNUPercent", mmPctAuto(_gnu));
        o += elem("GNUGrade",   tmGrade(_gnu));

        //   FPD — grade + raw from q.fixedPatternDamage
        o += elem("FPDValue",   mmVal(_fpd));
        o += elem("FPDGrade",   tmGrade(_fpd));

        o += elem("DecodeGrade", tmGrade(_dec));

        //   AG (Print Growth) — v1.20: round raw to 1 decimal place
        function _round1(_x) {
            var _n = parseFloat(_x);
            return isNaN(_n) ? "" : s(Math.round(_n * 10) / 10);
        }
        var _agV = mmVal(_ag);
        o += elem("AGValue",    _agV ? _round1(_agV) : "");
        o += elem("AGGrade",    tmGrade(_ag));

        // ── 2D matrix / general characteristics (v1.18: from q.general) ───────
        //   Names confirmed from Cognex template: xDimension, contrastUniformity,
        //   horizontalBWG, verticalBWG.  Others probed below.
        //   v1.20: MatrixSize derived from modulationArray.length.
        //   Square symbols: modArray indexes ALL cells including a 1-cell QZ
        //   wrapper, so side = sqrt(length) - 2.  (22×22 ECC200: 484+92 frame
        //   cells = 576 = 24² → 24-2 = 22.  Verified live.)
        //   v1.22: rectangular symbols also wrap, so length = (rows+2)*(cols+2).
        //   Without authoritative row/col fields, use a length→size lookup over
        //   the 6 ECC200 rectangular sizes (verified scan: 16×36 → 18×38 = 684).
        var _modArr  = _pick(q, "modulationArray");
        var _modLen  = (_modArr && typeof _modArr.length !== "undefined") ? _modArr.length : 0;
        var _modSide = (_modLen > 0) ? Math.sqrt(_modLen) : 0;
        // v1.25: QR uses 4-module quiet zone per side (8 total subtracted from sqrt).
        //   DM uses 1-module QZ per side (2 total).
        //   QR v3: sqrt(1369)=37, 37-8=29 ✓  (v1.24 wrongly emitted 35 = 37-2)
        //   DM 22×22: sqrt(576)=24, 24-2=22 ✓
        var _qzOffset = _qrBranch ? 8 : 2;
        var _symSide = (_modSide === Math.floor(_modSide) && _modSide > 2) ? (_modSide - _qzOffset) : 0;
        var _msz     = (_symSide > 0) ? (_symSide + "x" + _symSide) : "";
        if (!_msz && _modLen > 0) {
            // ECC200 rectangular table (length-with-QZ → "rowsXcols")
            var _rectMap = {
                "200":  "8x18",
                "340":  "8x32",
                "392":  "12x26",
                "532":  "12x36",
                "684":  "16x36",
                "900":  "16x48"
            };
            var _rectHit = _rectMap[String(_modLen)];
            if (_rectHit) { _msz = _rectHit; }
        }

        //   v1.20: codewordArray.length = total codewords for ECC200.
        //   Iterate to count isCorrected==1 for ErrorsCorrected.
        var _cwArr   = _pick(q, "codewordArray");
        var _cwLen   = (_cwArr && typeof _cwArr.length !== "undefined") ? _cwArr.length : 0;
        var _ecCount = 0;
        if (_cwArr && _cwLen > 0) {
            for (var _i = 0; _i < _cwLen; _i++) {
                var _cw = _cwArr[_i];
                if (_cw && _cw["isCorrected"]) { _ecCount++; }
            }
        }

        //   v1.20: encodationAnalysisArray.length = encoded character count.
        var _eaArr = _pick(q, "encodationAnalysisArray");
        var _eaLen = (_eaArr && typeof _eaArr.length !== "undefined") ? _eaArr.length : 0;

        o += elem("MatrixSize",            _msz);
        o += elem("HorizontalBWG",         gnProp("horizontalBWG"));
        o += elem("VerticalBWG",           gnProp("verticalBWG"));
        o += elem("EncodedCharacters",     _eaLen > 0 ? s(_eaLen) : "");
        o += elem("TotalCodewords",        _cwLen > 0 ? s(_cwLen) : "");
        o += elem("DataCodewords",         "");   // firmware does not split
        o += elem("ErrorCorrectionBudget", "");   // firmware does not expose
        o += elem("ErrorsCorrected",       _cwLen > 0 ? s(_ecCount) : "");
        o += elem("ErrorCapacityUsed",     "");
        // v1.25: ErrorCorrectionType is symbology-aware.
        //   DM → "ECC200" (correct). QR → try ovProp("ecLevel") first, else
        //   probe q top-level; fallback "QR" until DebugSymbols0 reveals path.
        //   Wire fully in v1.26 once symbols[0].ecLevel path confirmed.
        var _ecType = _qrBranch
            ? (ovProp("ecLevel") || (q ? s(_pick(q, "ecLevel")) : "") || "QR")
            : "ECC200";
        o += elem("ErrorCorrectionType",   _ecType);
        o += elem("NominalXDim",           gnProp("xDimension"));
        // v1.25: PixelsPerModule removed — was always empty; <ModuleSizePx>
        //   (from r.symbology.moduleSize, confirmed 16.18 px) is the only source.
        o += elem("ImagePolarity",         gnProp("polarity") || gnProp("imagePolarity"));
        o += elem("ContrastUniformity",    gnProp("contrastUniformity"));
        o += elem("MRD",                   gnProp("MRD") || gnProp("mrd"));
        o += elem("ContrastUniformityRow", gnProp("contrastUniformityRow"));
        o += elem("ContrastUniformityCol", gnProp("contrastUniformityCol"));
        // v1.25: OpticsSource — LoadedImage when ContrastUniformity===-1 AND MRD===-1.
        //   Both must be −1 simultaneously; each alone is insufficient.
        //   Confirmed across two loaded-image captures: URL QR (SymbolAngle=360°)
        //   and email QR (SymbolAngle=0°) are both loaded images — SymbolAngle
        //   is NOT a reliable discriminator. ContrastUniformity and MRD are always
        //   −1 on loaded images (not computed from optics) and always real values
        //   on live scans. Parse as float to handle both "−1" string and −1 number.
        var _cuNum  = parseFloat(gnProp("contrastUniformity"));
        var _mrdNum = parseFloat(gnProp("MRD") || gnProp("mrd"));
        o += elem("OpticsSource",
            (!isNaN(_cuNum) && _cuNum === -1 && !isNaN(_mrdNum) && _mrdNum === -1)
            ? "LoadedImage" : "LiveScan");

        // ── v1.19 new grade emissions ─────────────────────────────────────────
        //   distributedDamageGrade: per TKEnum it's a top-level q property —
        //   probe whether it's a TrucheckMetric (.grade) or a bare string.
        var _dd = _pick(q, "distributedDamageGrade");
        o += elem("DDGrade", (_dd && typeof _dd === "object") ? tmGrade(_dd) : s(_dd));
        //   averageGrade: separate TrucheckMetric carrying the ISO 15415 mean
        //   grade across all parameters (different from overall).
        var _avg = _pick(q, "averageGrade");
        o += elem("AverageGrade",        tmGrade(_avg));
        o += elem("AverageGradeNumeric", mmVal(_avg));
        //   v1.20: customNote = device calibration metadata string
        //   (e.g. "Results from UPCE-44960 Cal. 23 JUN 2023").
        o += elem("CustomNote",          s(_pick(q, "customNote")));

        // ── 2D L-side and quiet zones ─────────────────────────────────────────
        //   LLS/BLS — confirmed on q as leftLSide / bottomLSide (v1.15)
        o += elem("LLSGrade", tmGrade(_lls));
        o += elem("BLSGrade", tmGrade(_bls));
        //   LQZ/BQZ/TQZ/RQZ — from q TrucheckMetric sub-objects (confirmed v1.9)
        o += elem("LQZGrade", tmGrade(_lqz));
        o += elem("BQZGrade", tmGrade(_bqz));
        o += elem("TQZGrade", tmGrade(_tqz));
        o += elem("RQZGrade", tmGrade(_rqz));

        // ── 2D clock track / transition ratio grades ──────────────────────────
        //   v1.16: TTR/RTR percentages sourced from m.{horizontal,vertical}MarkMisplacement
        //   (DebugMEnum2 confirmed these are the ISO 15415 mark misplacement metrics,
        //   which are the DM-firmware equivalent of TTR/RTR in Webscan terminology).
        //   Grades for TTR/RTR continue to use q.topClockTrack/rightClockTrack until
        //   firmware exposes a markMisplacement.grade (the Metric debug probe will tell).
        //   TCT/RCT grades are the same q-side TrucheckMetric.
        //   v1.19: TTR/RTR have their own q-side TrucheckMetrics (revealed by
        //   v1.18 DebugTKEnum) — topTransitionRatio / rightTransitionRatio.
        //   Both have .raw (numeric) and .grade.  No longer aliased to TCT/RCT.
        var _ttr = _pick(q, "topTransitionRatio");
        var _rtr = _pick(q, "rightTransitionRatio");
        o += elem("TTRPercent", mmPctAuto(_ttr));
        o += elem("TTRGrade",   tmGrade(_ttr));
        o += elem("RTRPercent", mmPctAuto(_rtr));
        o += elem("RTRGrade",   tmGrade(_rtr));
        o += elem("TCTGrade",   tmGrade(_tct));
        o += elem("RCTGrade",   tmGrade(_rct));

        // ── Per-region parameters (≥ 32×32 / 2-region rectangular) ───────────
        //   v1.22: ROLLED BACK.  v1.21 probes (DebugULP/URP/LLP/HCT/VCT) all
        //   returned grade=F/numericGrade=0 on BOTH a 32×32 (scan graded C
        //   overall) AND a 16×36 2-region rect (scan graded D overall) — even
        //   though the device's PDF report shows ULQZ/URQZ/RUQZ/RLQZ all = A
        //   for the same 16×36 scan.  q.upperLeftPattern / upperRightPattern /
        //   lowerLeftPattern / horizontalClockTrack / verticalClockTrack /
        //   alignmentPatterns are inert placeholder objects in this firmware's
        //   JS scope.  Emitting "F" for them was false data, so they go back
        //   to empty.  v1.22 probes r.metrics / r-siblings to look for a
        //   different scope that might carry per-region data.
        o += elem("ULQZGrade",        "");
        o += elem("URQZGrade",        "");
        o += elem("RUQZGrade",        "");
        o += elem("RLQZGrade",        "");
        o += elem("LLQZGrade",        "");
        o += elem("LRQZGrade",        "");
        o += elem("HClockTrackGrade", "");
        o += elem("VClockTrackGrade", "");
        o += elem("ULQTTRPercent", "");
        o += elem("ULQTTRGrade",   "");
        o += elem("URQTTRPercent", "");
        o += elem("URQTTRGrade",   "");
        o += elem("LLQTTRPercent", "");
        o += elem("LLQTTRGrade",   "");
        o += elem("LRQTTRPercent", "");
        o += elem("LRQTTRGrade",   "");
        o += elem("ULQRTRPercent", "");
        o += elem("ULQRTRGrade",   "");
        o += elem("URQRTRPercent", "");
        o += elem("URQRTRGrade",   "");
        o += elem("LLQRTRPercent", "");
        o += elem("LLQRTRGrade",   "");
        o += elem("LRQRTRPercent", "");
        o += elem("LRQRTRGrade",   "");
        o += elem("ULQTCTGrade",   "");
        o += elem("URQTCTGrade",   "");
        o += elem("LLQTCTGrade",   "");
        o += elem("LRQTCTGrade",   "");
        o += elem("ULQRCTGrade",   "");
        o += elem("URQRCTGrade",   "");
        o += elem("LLQRCTGrade",   "");
        o += elem("LRQRCTGrade",   "");

        // ── 1D / ANSI summary ─────────────────────────────────────────────────
        var _ansiGrade = prop(m, "ansiGrade") || prop(m, "symbolAnsiGrade");
        o += elem("SymbolAnsiGrade", _ansiGrade);
        o += elem("AvgEdge",   prop(m, "avgEdge"));
        o += elem("AvgRlRd",   prop(m, "avgRlRd"));
        o += elem("AvgSC",     prop(m, "avgSc"));
        o += elem("AvgMinEC",  prop(m, "avgMinEc"));
        o += elem("AvgMOD",    prop(m, "avgMod"));
        o += elem("AvgDefect", prop(m, "avgDefect"));
        o += elem("AvgDcod",   prop(m, "avgDcod"));
        o += elem("AvgDEC",    prop(m, "avgDec"));
        o += elem("AvgLQZ",    prop(m, "avgLqz"));
        o += elem("AvgRQZ",    prop(m, "avgRqz"));
        o += elem("AvgHQZ",    prop(m, "avgHqz"));
        o += elem("AvgMinQZ",  prop(m, "avgMinQz"));
        // v1.24: BWG = Print Growth per comms-and-programming-guide §4. The
        // r.metrics.printGrowth is a Metric object {raw, grade}, not a scalar,
        // so prop() returns "[object Metric]" / undefined. Use mmPctAuto.
        o += elem("BWGPercent",    mmPctAuto(_pick(m, "printGrowth")));
        o += elem("Magnification", prop(m, "magnification"));
        o += elem("Ratio",         prop(m, "ratio"));
        o += elem("NominalXDim1D", prop(m, "nominalXDim1D") || prop(m, "nominalXDim1d"));

    } // end if (q)

    // ── v1.24 + v1.25 wires (post-q-block) ───────────────────────────────────
    //   v1.25: Full JPEG image payload. v1.24 DebugJpegProbe confirmed:
    //   type=string, len≈9,912 for live DM / ≈22,260 for loaded QR — well
    //   within Network Client buffer. r.trucheck.jpegImage is a base64 JPEG
    //   string. Emit as-is; parser stores in VerificationRecord.JpegImageBase64.
    o += elem("JpegImageBase64", q ? s(_pick(q, "jpegImage")) : "");

    //   Promote r.trucheck.calibrationDate to first-class. v1.23 stuffed
    //   calibration metadata into <CustomNote> as a workaround; comms guide §6
    //   confirms r.trucheck.calibrationDate is a string timestamp directly
    //   addressable. CustomNote stays as-is (different data — user-set device
    //   note "Results from UPCE-44960 Cal. 23 JUN 2023").
    o += elem("CalibrationDate", q ? s(_pick(q, "calibrationDate")) : "");

    //   readerProperties.status3D.{fieldCalibrated, factoryCalibrated} — input
    //   to the OpticsCompliant logic per comms guide §6. Booleans.
    var _s3d = _pick(rp, "status3D");
    o += elem("FieldCalibrated",   _s3d ? s(_pick(_s3d, "fieldCalibrated"))   : "");
    o += elem("FactoryCalibrated", _s3d ? s(_pick(_s3d, "factoryCalibrated")) : "");

    //   MinPassGrade — per session plan A5. v1.23 DebugMetricsKeys enumerated
    //   minPassGrade=[obj] on r.metrics. Probe both .grade and .raw shapes.
    var _mpg = _pick(m, "minPassGrade");
    o += elem("MinPassGrade", tmGrade(_mpg));
    o += elem("MinPassRaw",   mmVal(_mpg));

    // ── v1.25 introspection probes ────────────────────────────────────────────
    //   Retained from v1.20: DebugModSize, DebugECCount (formula sanity).
    //   Retained from v1.22: DebugMetricsKeys, DebugRSiblings (regression
    //                        baselines; bit-identical across DMST 25→26.1).
    //   Retained from v1.24: DebugGS1, DebugDodUid (require correct Application
    //                        Settings — set Data Format Check = GS1 then scan a
    //                        GS1 symbol), DebugBarcodeAsgn (+ .stats deep-probe).
    //   NEW v1.25:
    //     • DebugSymbols0   — q.symbols[0] full enum. Target fields:
    //         QR: ulpGrade/urpGrade/llpGrade (finder patterns),
    //             hctGrade/vctGrade (clock tracks), alpGrade (alignment),
    //             vibGrade/fibGrade (version/format info),
    //             ecLevel/maskPattern/rows/cols.
    //         DM: ulqzGrade/urqzGrade/ruqzGrade/rlqzGrade (per-region QZ),
    //             dataCodewords/ecCodewords/ecCapacityUsed,
    //             minLightReflectance/maxDarkReflectance.
    //     • DebugPrintGrowth — m.printGrowth shape to diagnose empty BWGPercent.
    //   FIXED v1.25:
    //     • DebugImageShape — guarded with try/catch (v1.24: silent error,
    //                         element absent from output).
    //   DROPPED v1.25:
    //     • DebugANUCase    — resolved: lowercase 'axialNonuniformity' confirmed.
    //     • DebugJpegProbe  — replaced by <JpegImageBase64> first-class wire.
    //     • DebugReaderProps — answered: status3D paths confirmed in v1.24.

    function _enumKV22(obj, label) {
        if (!obj) { return "(" + label + " null)"; }
        if (typeof obj !== "object") {
            return "(" + label + " " + (typeof obj) + "=" + String(obj).substring(0, 30) + ")";
        }
        var _out = "";
        for (var _k in obj) {
            var _v = obj[_k];
            var _t = typeof _v;
            var _vs;
            if (_t === "object" && _v !== null) {
                _vs = (typeof _v.length !== "undefined") ? ("[arr." + _v.length + "]") : "[obj]";
            } else {
                _vs = String(_v).substring(0, 40);
            }
            _out += _k + "=" + _vs + ";";
        }
        return _out || "(" + label + " empty)";
    }

    // Retained from v1.20 — formula sanity check
    var _v25mod  = _pick(q, "modulationArray");
    var _v25mLen = (_v25mod && typeof _v25mod.length !== "undefined") ? _v25mod.length : 0;
    var _v25mSq  = (_v25mLen > 0) ? Math.sqrt(_v25mLen) : 0;
    o += elem("DebugModSize",
        "len=" + _v25mLen + " sqrt=" + _v25mSq + " sqr=" + (_v25mSq === Math.floor(_v25mSq)));
    var _v25cw   = _pick(q, "codewordArray");
    var _v25cwL  = (_v25cw && typeof _v25cw.length !== "undefined") ? _v25cw.length : 0;
    var _v25ec   = 0;
    if (_v25cw) {
        for (var _v25i = 0; _v25i < _v25cwL; _v25i++) {
            if (_v25cw[_v25i] && _v25cw[_v25i]["isCorrected"]) { _v25ec++; }
        }
    }
    o += elem("DebugECCount", "total=" + _v25cwL + " corrected=" + _v25ec);

    // Retained from v1.22 — regression baseline
    o += elem("DebugMetricsKeys", _enumKV22(m, "r.metrics"));

    var _v25known = { "trucheck": 1, "metrics": 1 };
    var _v25sibStr = "";
    if (r && typeof r === "object") {
        for (var _v25rk in r) {
            if (_v25known[_v25rk]) { continue; }
            var _v25rv = r[_v25rk];
            var _v25rvDesc;
            if (_v25rv === null) { _v25rvDesc = "null"; }
            else if (typeof _v25rv === "object") {
                _v25rvDesc = (typeof _v25rv.length !== "undefined") ? ("[arr." + _v25rv.length + "]") : "[obj]";
            } else {
                _v25rvDesc = (typeof _v25rv) + "=" + String(_v25rv).substring(0, 30);
            }
            _v25sibStr += _v25rk + "=" + _v25rvDesc + ";";
        }
    }
    o += elem("DebugRSiblings", _v25sibStr || "(no unknown r-siblings)");

    // NEW PROBE: q.symbols[0] full enum. Targets QR-specific grade params
    //   and DataCodewords/ECLevel/MaskPattern for both QR and DM.
    //   If symbols array doesn't exist → reports "(q.symbols[0] null)".
    var _v25sym0 = (q && q["symbols"]
                   && typeof q["symbols"].length !== "undefined"
                   && q["symbols"].length > 0)
                  ? q["symbols"][0] : null;
    o += elem("DebugSymbols0", _enumKV22(_v25sym0, "q.symbols[0]"));

    // FIXED PROBE: r.image enum. v1.24 emitted nothing (silent error or
    //   r.image undefined). Try/catch reports any JS error explicitly.
    try {
        o += elem("DebugImageShape", _enumKV22(_pick(r, "image"), "r.image"));
    } catch (_imgErr) {
        o += elem("DebugImageShape", "(error: " + _imgErr.message + ")");
    }

    // NEW PROBE: m.printGrowth shape to diagnose empty <BWGPercent>.
    //   Expected: Metric {raw, grade}. If raw===-1 → sentinel, empty is correct.
    //   If undefined/null → accessor path is wrong on this firmware build.
    //   If number (not object) → use directly instead of mmPctAuto.
    var _v25pg = _pick(m, "printGrowth");
    var _v25pgDesc;
    if (_v25pg === null || typeof _v25pg === "undefined") {
        _v25pgDesc = "null";
    } else if (typeof _v25pg === "object") {
        _v25pgDesc = _enumKV22(_v25pg, "printGrowth");
    } else {
        _v25pgDesc = (typeof _v25pg) + "=" + String(_v25pg).substring(0, 60);
    }
    o += elem("DebugPrintGrowth", _v25pgDesc);

    // Retained from v1.24: requires Data Format Check = GS1 in Application
    //   Settings to populate. All 187 AIs returned undefined in v1.24 because
    //   Data Format Check was set to None — device never ran a GS1 check.
    o += elem("DebugGS1",
        _enumKV22(_pick(_pick(r, "validation"), "gs1"), "r.validation.gs1"));

    // Retained from v1.24: requires DoD UID Application Setting active.
    o += elem("DebugDodUid",
        _enumKV22(_pick(_pick(r, "validation"), "dodUid"), "r.validation.dodUid"));

    // Retained from v1.24 + deep-probe .stats sub-object.
    var _v25ba      = _pick(r, "barcodeAssignment");
    var _v25baStats = _v25ba ? _pick(_v25ba, "stats") : null;
    o += elem("DebugBarcodeAsgn",
        _enumKV22(_v25ba, "r.barcodeAssignment") +
        (_v25baStats ? (" stats:" + _enumKV22(_v25baStats, ".stats")) : ""));

    o += '</DMSymVerResponse>\r\n'
       + '</DMCCResponse>';

    // Firmware 5.x / 6.x compatibility:
    // Older firmware exposes a global 'output' object; newer firmware passes
    // the output as the third parameter 'outputResults'.  Try both so the
    // script works across revisions without modification.
    if (typeof outputResults !== "undefined" && outputResults !== null) {
        outputResults.content = o;
    } else {
        output.content = o;
    }
}
