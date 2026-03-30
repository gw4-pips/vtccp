// ─────────────────────────────────────────────────────────────────────────────
// VTCCP DMST Push Script
//
//   Version   : 1.15
//   Generated : 2026-03-30 UTC
//   Source    : VTCCP Replit Agent  (github.com/gw4-pips/vtccp)
//   Target    : Cognex DataMan firmware 5.x / 6.x  /  DMV475
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

    o += elem("PushScriptDiag", "v1.15 q=" + _qSrc + " m=" + (m ? "found" : "null")
          + " r.decoded=" + s(r && r.decoded)
          + " rType=" + (typeof r));

    // ── v1.15 r.metrics for-in enumeration (pages 1 & 2) ─────────────────────
    // q (r.trucheck) property list fully confirmed by v1.14 scan; no longer probed.
    // m (r.metrics) was truncated in v1.14; two pages to capture the full list.
    // Page 1 (entries 0-19) — confirmed: symbolContrast, cellContrast,
    //   axialNonUniformity, printGrowth, UEC, modulation, fixedPatternDamage,
    //   gridNonUniformity, extremeReflectance, reflectMin, edgeContrastMin,
    //   singleScanInt, multiScanInt, signalToNoiseRatio, horizontalMarkGrowth,
    //   verticalMarkGrowth (v1.14 scan — cut off at #16).
    // Page 2 (entries 16+) — looking for TTR%, RTR%, and any remaining unknowns.
    var _mEnum = "";
    var _mEnum2 = "";
    var _mEnumIdx = 0;
    if (m) {
        for (var _mk in m) {
            var _mkStr = _mk + "=" + String(m[_mk]).substring(0, 20) + ";";
            if (_mEnumIdx < 16) { _mEnum  += _mkStr; }
            else                { _mEnum2 += _mkStr; }
            _mEnumIdx++;
        }
    }
    o += elem("DebugMEnum",  _mEnum  || "(m null)");
    o += elem("DebugMEnum2", _mEnum2 || "(m has ≤16 props)");

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

        // ── q (r.trucheck) TrucheckMetric bindings — confirmed names from v1.14 scan
        var _uec = _pick(q, "unusedErrorCorrection");  // UEC  (v1.15 — was wrong name)
        var _sc  = _pick(q, "symbolContrast");
        var _mod = _pick(q, "modulation");
        var _rm  = _pick(q, "reflectanceMargin");
        var _anu = _pick(q, "axialNonuniformity");     // ANU  lowercase 'u' (v1.15)
        var _gnu = _pick(q, "gridNonuniformity");      // GNU  lowercase 'u' (v1.15)
        var _fpd = _pick(q, "fixedPatternDamage");
        var _lls = _pick(q, "leftLSide");              // LLS  (v1.15 — was 'leftL')
        var _bls = _pick(q, "bottomLSide");            // BLS  (v1.15 — was 'bottomL')
        var _lqz = _pick(q, "leftQuietZone");
        var _bqz = _pick(q, "bottomQuietZone");
        var _rqz = _pick(q, "rightQuietZone");
        var _tqz = _pick(q, "topQuietZone");
        var _tct = _pick(q, "topClockTrack");
        var _rct = _pick(q, "rightClockTrack");
        var _dec = _pick(q, "decode");
        var _ag  = _pick(q, "printGrowth");

        // ── m (r.metrics) Rl/Rd bindings for true ISO SC% and SCRlRd (v1.15) ──
        //   extremeReflectance = Rl (max module reflectance, 0–1 ratio)
        //   reflectMin         = Rd (min module reflectance, 0–1 ratio)
        //   True SC% = (Rl − Rd) × 100; e.g. (0.75 − 0.04) × 100 = 71%.
        var _mRl    = _pick(m, "extremeReflectance");
        var _mRd    = _pick(m, "reflectMin");
        var _rlRaw  = (_mRl && typeof _mRl["raw"] !== "undefined") ? parseFloat(_mRl["raw"]) : NaN;
        var _rdRaw  = (_mRd && typeof _mRd["raw"] !== "undefined") ? parseFloat(_mRd["raw"]) : NaN;
        var _rlOk   = !isNaN(_rlRaw) && _rlRaw !== -1;
        var _rdOk   = !isNaN(_rdRaw) && _rdRaw !== -1;
        var _scPct  = (_rlOk && _rdOk)
                        ? s(Math.round((_rlRaw - _rdRaw) * 1000) / 10)
                        : mmPct(_pick(m, "symbolContrast"));  // fallback
        var _rlInt  = _rlOk ? String(Math.round(_rlRaw * 100)) : "";
        var _rdInt  = _rdOk ? String(Math.round(_rdRaw * 100)) : "";
        var _scRlRd = (_rlInt && _rdInt) ? (_rlInt + "/" + _rdInt) : "";

        // ── Grading summary ───────────────────────────────────────────────────
        //   m.overallGrade is a [object Metric] with .grade (letter/NA) / .raw (number/-1).
        //   mmGrade passes "NA" through; mmVal suppresses -1 (numeric sentinel only).
        var _mOG       = _pick(m, "overallGrade");
        var _ogLetter  = mmGrade(_mOG);
        var _ogNumeric = mmVal(_mOG);
        o += elem("FormalGrade",         _ogLetter);
        o += elem("OverallGrade",        _ogLetter);
        o += elem("OverallGradeNumeric", _ogNumeric);

        // ── Verification conditions ───────────────────────────────────────────
        //   aperture / wavelength / lighting / standard — expected on m; empty until confirmed
        o += elem("ApertureRef", prop(m, "aperture") || prop(m, "apertureRef"));
        o += elem("Wavelength",  prop(m, "wavelength"));
        o += elem("Lighting",    prop(m, "lighting") || prop(m, "lightingType"));
        o += elem("Standard",    prop(m, "standard") || prop(m, "verificationStandard"));

        // ── 2D ISO 15415 quality parameters ───────────────────────────────────
        //   UEC — grade from q.unusedErrorCorrection; % from m.UEC (v1.15)
        o += elem("UECPercent", mmPct(_pick(m, "UEC")));   // 'UEC' key on m (not 'uniformEdgeContrast')
        o += elem("UECGrade",   tmGrade(_uec));

        //   SC — grade from q.symbolContrast; true ISO SC% = (Rl−Rd)×100 (v1.15)
        //        Rl = m.extremeReflectance.raw, Rd = m.reflectMin.raw (both 0–1 ratio)
        o += elem("SCPercent",  _scPct);
        o += elem("SCRlRd",     _scRlRd);
        o += elem("SCGrade",    tmGrade(_sc));

        //   MOD — grade from q.modulation
        o += elem("MODGrade",   tmGrade(_mod));

        //   RM — grade from q.reflectanceMargin
        o += elem("RMGrade",    tmGrade(_rm));

        //   ANU — grade from q.axialNonuniformity (lowercase u); % from m.axialNonUniformity (capital U)
        var _mANU = _pick(m, "axialNonUniformity");   // capital U on m-side
        o += elem("ANUPercent", mmPct(_mANU));
        o += elem("ANUGrade",   tmGrade(_anu));        // from q (v1.15 — was mmGrade(_mANU))

        //   GNU — grade from q.gridNonuniformity (lowercase u); % from m.gridNonUniformity (capital U)
        var _mGNU = _pick(m, "gridNonUniformity");    // capital U on m-side
        o += elem("GNUPercent", mmPct(_mGNU));
        o += elem("GNUGrade",   tmGrade(_gnu));        // from q (v1.15 — was mmGrade(_mGNU))

        //   FPD — grade from q.fixedPatternDamage
        o += elem("FPDGrade",   tmGrade(_fpd));

        //   Decode — grade from q.decode
        o += elem("DecodeGrade", tmGrade(_dec));

        //   AG (Print Growth) — grade from q.printGrowth; measurement from m (Metric)
        o += elem("AGValue",    mmVal(_pick(m, "printGrowth")));
        o += elem("AGGrade",    tmGrade(_ag));

        // ── 2D matrix characteristics ─────────────────────────────────────────
        //   All expected on m (r.metrics) — DebugMetricsFound will confirm names
        var _rows = prop(m, "rows");
        var _cols = prop(m, "columns") || prop(m, "cols");
        var _msz  = prop(m, "matrixSize") || prop(m, "size")
                    || ((_rows && _cols) ? (_rows + "x" + _cols) : "");
        o += elem("MatrixSize",            _msz);
        o += elem("HorizontalBWG",         mmVal(_pick(m, "horizontalMarkGrowth")));  // v1.15 confirmed
        o += elem("VerticalBWG",           mmVal(_pick(m, "verticalMarkGrowth")));
        o += elem("EncodedCharacters",     prop(m, "encodedCharacters") || prop(m, "encodedChars"));
        o += elem("TotalCodewords",        prop(m, "totalCodewords"));
        o += elem("DataCodewords",         prop(m, "dataCodewords"));
        o += elem("ErrorCorrectionBudget", prop(m, "errorCorrectionBudget") || prop(m, "ecBudget"));
        o += elem("ErrorsCorrected",       prop(m, "errorsCorrected")       || prop(m, "ecCorrected"));
        o += elem("ErrorCapacityUsed",     prop(m, "errorCapacityUsed")     || prop(m, "ecCapacityUsed"));
        o += elem("ErrorCorrectionType",   prop(m, "errorCorrectionType")   || prop(m, "ecType"));
        o += elem("NominalXDim",           prop(m, "nominalXDim"));
        o += elem("PixelsPerModule",       prop(m, "pixelsPerModule") || prop(m, "ppm"));
        o += elem("ImagePolarity",         prop(m, "polarity") || prop(m, "imagePolarity"));
        o += elem("ContrastUniformity",    mmVal(_pick(m, "contrastUniformity")));
        o += elem("MRD",                   prop(m, "mrd") || prop(m, "minReflectanceDifference"));

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
        //   topClockTrack    → TTRGrade and TCTGrade (best guess until confirmed)
        //   rightClockTrack  → RTRGrade and RCTGrade
        //   % values expected on m — empty until DebugMetricsFound confirms names
        o += elem("TTRPercent", prop(m, "ttrPercent") || prop(m, "topToTopRatio") || prop(m, "ttr"));
        o += elem("TTRGrade",   tmGrade(_tct));
        o += elem("RTRPercent", prop(m, "rtrPercent") || prop(m, "rightToRightRatio") || prop(m, "rtr"));
        o += elem("RTRGrade",   tmGrade(_rct));
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
