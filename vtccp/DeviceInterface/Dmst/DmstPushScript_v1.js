// ─────────────────────────────────────────────────────────────────────────────
// VTCCP DMST Push Script
//
//   Version   : 1.20
//   Generated : 2026-05-17 UTC
//   Source    : VTCCP Replit Agent  (github.com/gw4-pips/vtccp)
//   Target    : Cognex DataMan firmware 5.x / 6.x  /  DMV475
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
    o += elem("DecodedData",   (r && r.decoded) ? esc(r.content) : "");

    // ── Diagnostic elements (v1.7) ───────────────────────────────────────────
    // Reports which quality-object path resolved (or "none"), plus a comma-
    // separated list of every property name that is defined on r and rp.
    // Visible in VTCCP VS Output as [VTCCP-DMST] RawXML.
    // Remove this section once the correct property names are confirmed.

    // r.metrics — sibling of r.trucheck on decodeResults[0].
    // Expected to carry: overall grade, UEC/ANU/GNU, SC%/MOD%/RM%, dimensions.
    var m = _pick(r, "metrics");

    o += elem("PushScriptDiag", "v1.20 q=" + _qSrc + " m=" + (m ? "found" : "null"));

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
        //   The array indexes ALL cells including a 1-cell quiet-zone wrapper,
        //   so symbol side = sqrt(length) - 2.  (22×22 ECC200: 484+92 frame
        //   cells = 576 = 24² → 24-2 = 22.  Verified live.)
        var _modArr  = _pick(q, "modulationArray");
        var _modLen  = (_modArr && typeof _modArr.length !== "undefined") ? _modArr.length : 0;
        var _modSide = (_modLen > 0) ? Math.sqrt(_modLen) : 0;
        var _symSide = (_modSide === Math.floor(_modSide) && _modSide > 2) ? (_modSide - 2) : 0;
        var _msz     = (_symSide > 0) ? (_symSide + "x" + _symSide) : "";

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
        o += elem("ErrorCorrectionType",   "ECC200");   // implied by Data Matrix
        o += elem("NominalXDim",           gnProp("xDimension"));
        o += elem("PixelsPerModule",       gnProp("pixelsPerModule") || gnProp("ppm"));
        o += elem("ImagePolarity",         gnProp("polarity") || gnProp("imagePolarity"));
        o += elem("ContrastUniformity",    gnProp("contrastUniformity"));
        o += elem("MRD",                   gnProp("MRD") || gnProp("mrd"));
        o += elem("ContrastUniformityRow", gnProp("contrastUniformityRow"));
        o += elem("ContrastUniformityCol", gnProp("contrastUniformityCol"));

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

        // ── Per-quadrant parameters (matrices ≥ 32×32 only) ──────────────────
        //   No JS property names found yet for per-quadrant data; all empty.
        o += elem("ULQZGrade",     "");
        o += elem("URQZGrade",     "");
        o += elem("RUQZGrade",     "");
        o += elem("RLQZGrade",     "");
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
        o += elem("BWGPercent",    prop(m, "bwgPercent") || prop(m, "bwg"));
        o += elem("Magnification", prop(m, "magnification"));
        o += elem("Ratio",         prop(m, "ratio"));
        o += elem("NominalXDim1D", prop(m, "nominalXDim1D") || prop(m, "nominalXDim1d"));

    } // end if (q)

    // ── v1.20 introspection probes (reduced) ──────────────────────────────────
    //   Three retained for cross-validation of v1.20's MatrixSize / EC count /
    //   modulation-array shape.  All others stripped — territory is mapped.
    //   v1.21 will remove these final probes.
    var _v20mod  = _pick(q, "modulationArray");
    var _v20mLen = (_v20mod && typeof _v20mod.length !== "undefined") ? _v20mod.length : 0;
    var _v20mSq  = (_v20mLen > 0) ? Math.sqrt(_v20mLen) : 0;
    o += elem("DebugModSize",
        "len=" + _v20mLen + " sqrt=" + _v20mSq + " sqr=" + (_v20mSq === Math.floor(_v20mSq)));

    //   Full enumeration of modulationArray[0] with longer value windows
    //   to verify the "grade=(" oddity from v1.19 (was substring truncated
    //   to 1 char — see if grade is actually a letter, paren-wrapped string,
    //   or numeric code).
    var _v20m0 = (_v20mod && _v20mLen > 0) ? _v20mod[0] : null;
    var _v20m0Str = "(null)";
    if (_v20m0 && typeof _v20m0 === "object") {
        _v20m0Str = "";
        for (var _v20k in _v20m0) {
            var _v20v = _v20m0[_v20k];
            _v20m0Str += _v20k + "(" + (typeof _v20v) + ")=" + String(_v20v).substring(0, 40) + ";";
        }
    }
    o += elem("DebugModSample", _v20m0Str);

    //   Cross-check ErrorsCorrected: re-count and report alongside total.
    var _v20cw   = _pick(q, "codewordArray");
    var _v20cwL  = (_v20cw && typeof _v20cw.length !== "undefined") ? _v20cw.length : 0;
    var _v20ec   = 0;
    if (_v20cw) {
        for (var _v20i = 0; _v20i < _v20cwL; _v20i++) {
            if (_v20cw[_v20i] && _v20cw[_v20i]["isCorrected"]) { _v20ec++; }
        }
    }
    o += elem("DebugECCount", "total=" + _v20cwL + " corrected=" + _v20ec);

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
