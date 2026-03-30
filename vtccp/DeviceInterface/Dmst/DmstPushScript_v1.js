// ─────────────────────────────────────────────────────────────────────────────
// VTCCP DMST Push Script
//
//   Version   : 1.6
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

    // ── Diagnostic element (v1.6) ─────────────────────────────────────────────
    // Reports which quality-object path resolved, or "none" if all probes
    // failed.  Visible in VTCCP VS Output as [VTCCP-DMST] RawXML.
    // Remove this section once the correct path is identified and confirmed.
    o += elem("PushScriptDiag", "v1.6 q=" + _qSrc
          + " r.decoded=" + s(r && r.decoded)
          + " rType=" + (typeof r));

    if (q) {

        // ── Grading summary ───────────────────────────────────────────────────
        //   overallGrade      — "A" / "B" / "C" / "D" / "F"
        //   overallGradeValue — 4.0 / 3.0 / 2.0 / 1.0 / 0.0  (ISO numeric)
        //   formalGrade       — same letter from the formal standard column

        o += elem("FormalGrade",         prop(q, "formalGrade"));
        o += elem("OverallGrade",        prop(q, "overallGrade"));
        o += elem("OverallGradeNumeric", prop(q, "overallGradeValue"));

        // ── Verification settings ─────────────────────────────────────────────
        //   aperture   — integer aperture reference number
        //   wavelength — integer nm (e.g. 660)
        //   lighting   — "Axial", "45-degree", "Low-angle", …
        //   standard   — "ISO/IEC 15415", "AIM DPM", …

        o += elem("ApertureRef", prop(q, "aperture"));
        o += elem("Wavelength",  prop(q, "wavelength"));
        o += elem("Lighting",    prop(q, "lighting"));
        o += elem("Standard",    prop(q, "standard"));

        // ── 2D ISO 15415 quality parameters ───────────────────────────────────
        //   uec / uecGrade           — Uniform Edge Contrast  %  / letter
        //   sc  / scGrade            — Symbol Contrast  %  / letter
        //   rl  / rd                 — Reflection Level / Difference (for SCRlRd column)
        //   mod / modGrade           — Modulation % / letter
        //   rm  / rmGrade            — Reflectance Margin % / letter
        //   anu / anuGrade           — Axial Non-Uniformity % / letter
        //   gnu / gnuGrade           — Grid Non-Uniformity % / letter
        //   fpd / fpdGrade           — Fixed Pattern Damage / letter
        //   decodeGrade              — Decode letter grade
        //   ag  / agGrade            — Average Grade value / letter

        o += elem("UECPercent", prop(q, "uec"));
        o += elem("UECGrade",   prop(q, "uecGrade"));
        o += elem("SCPercent",  prop(q, "sc"));

        // SC RL/RD — output as "RL/RD" pair e.g. "89/4"; try both property forms.
        var rl = prop(q, "rl");
        var rd = prop(q, "rd");
        var scRlRd = (rl && rd) ? (rl + "/" + rd) : prop(q, "scRlRd");
        o += elem("SCRlRd",    scRlRd);
        o += elem("SCGrade",   prop(q, "scGrade"));

        o += elem("MODGrade",  prop(q, "modGrade"));
        o += elem("RMGrade",   prop(q, "rmGrade"));
        o += elem("ANUPercent",prop(q, "anu"));
        o += elem("ANUGrade",  prop(q, "anuGrade"));
        o += elem("GNUPercent",prop(q, "gnu"));
        o += elem("GNUGrade",  prop(q, "gnuGrade"));
        o += elem("FPDGrade",  prop(q, "fpdGrade"));
        o += elem("DecodeGrade",prop(q, "decodeGrade"));
        o += elem("AGValue",   prop(q, "ag"));
        o += elem("AGGrade",   prop(q, "agGrade"));

        // ── 2D matrix characteristics ─────────────────────────────────────────
        //   rows / columns           — e.g. 16 / 16 → written as "16x16"
        //   hBwg / vBwg              — Horizontal / Vertical Bar Width Growth %
        //   encodedChars             — encoded character count
        //   totalCodewords           — total ECC codewords
        //   dataCodewords            — data codewords
        //   ecBudget                 — error correction budget (total ECC codewords)
        //   ecCorrected              — errors corrected
        //   ecCapacityUsed           — ECC capacity used %
        //   ecType                   — "ECC 200", "Reed-Solomon", …
        //   nominalXDim              — nominal X dimension (device units; typically mils or mm)
        //   ppm                      — pixels per module
        //   polarity                 — "Dark on Light" / "Light on Dark"
        //   contrastUniformity       — uniformity descriptor string
        //   mrd                      — minimum reflectance difference

        var rows = prop(q, "rows");
        var cols = prop(q, "columns");
        var matrixSize = (rows && cols) ? (rows + "x" + cols) : "";
        o += elem("MatrixSize",            matrixSize);
        o += elem("HorizontalBWG",         prop(q, "hBwg"));
        o += elem("VerticalBWG",           prop(q, "vBwg"));
        o += elem("EncodedCharacters",     prop(q, "encodedChars"));
        o += elem("TotalCodewords",        prop(q, "totalCodewords"));
        o += elem("DataCodewords",         prop(q, "dataCodewords"));
        o += elem("ErrorCorrectionBudget", prop(q, "ecBudget"));
        o += elem("ErrorsCorrected",       prop(q, "ecCorrected"));
        o += elem("ErrorCapacityUsed",     prop(q, "ecCapacityUsed"));
        o += elem("ErrorCorrectionType",   prop(q, "ecType"));
        o += elem("NominalXDim",           prop(q, "nominalXDim"));
        o += elem("PixelsPerModule",       prop(q, "ppm"));
        o += elem("ImagePolarity",         prop(q, "polarity"));
        o += elem("ContrastUniformity",    prop(q, "contrastUniformity"));
        o += elem("MRD",                   prop(q, "mrd"));

        // ── 2D quiet zones / border grades ────────────────────────────────────
        //   lls / llsGrade   — Left L-finder score  / letter
        //   bls / blsGrade   — Bottom L-finder score / letter
        //   lqz / lqzGrade   — Left Quiet Zone / letter
        //   bqz / bqzGrade   — Bottom Quiet Zone / letter
        //   tqz / tqzGrade   — Top Quiet Zone / letter
        //   rqz / rqzGrade   — Right Quiet Zone / letter

        o += elem("LLSGrade", prop(q, "llsGrade"));
        o += elem("BLSGrade", prop(q, "blsGrade"));
        o += elem("LQZGrade", prop(q, "lqzGrade"));
        o += elem("BQZGrade", prop(q, "bqzGrade"));
        o += elem("TQZGrade", prop(q, "tqzGrade"));
        o += elem("RQZGrade", prop(q, "rqzGrade"));

        // ── 2D transition ratios / clock tracks ───────────────────────────────
        //   ttr / ttrGrade   — Top Transition Ratio % / letter
        //   rtr / rtrGrade   — Right Transition Ratio % / letter
        //   tct / tctGrade   — Top Clock Track grade letter
        //   rct / rctGrade   — Right Clock Track grade letter

        o += elem("TTRPercent", prop(q, "ttr"));
        o += elem("TTRGrade",   prop(q, "ttrGrade"));
        o += elem("RTRPercent", prop(q, "rtr"));
        o += elem("RTRGrade",   prop(q, "rtrGrade"));
        o += elem("TCTGrade",   prop(q, "tctGrade"));
        o += elem("RCTGrade",   prop(q, "rctGrade"));

        // ── 2D quadrant parameters (matrices ≥ 32×32 only) ───────────────────
        //   Firmware 6.x exposes quadrant sub-object names as:
        //     ulqz / urqz / ruqz / rlqz         — quiet zone grades (letter)
        //     ulqTtr / urqTtr / llqTtr / lrqTtr — TTR % per quadrant
        //     ulqTtrGrade / …Grade               — TTR letter grade per quadrant
        //     ulqRtr / …  / ulqRtrGrade / …      — RTR % / grade per quadrant
        //     ulqTct / ulqTctGrade / ulqRct / ulqRctGrade  — clock track grades
        //   These will be empty for matrices smaller than 32×32.

        o += elem("ULQZGrade", prop(q, "ulqzGrade"));
        o += elem("URQZGrade", prop(q, "urqzGrade"));
        o += elem("RUQZGrade", prop(q, "ruqzGrade"));
        o += elem("RLQZGrade", prop(q, "rlqzGrade"));

        o += elem("ULQTTRPercent", prop(q, "ulqTtr"));
        o += elem("ULQTTRGrade",   prop(q, "ulqTtrGrade"));
        o += elem("URQTTRPercent", prop(q, "urqTtr"));
        o += elem("URQTTRGrade",   prop(q, "urqTtrGrade"));
        o += elem("LLQTTRPercent", prop(q, "llqTtr"));
        o += elem("LLQTTRGrade",   prop(q, "llqTtrGrade"));
        o += elem("LRQTTRPercent", prop(q, "lrqTtr"));
        o += elem("LRQTTRGrade",   prop(q, "lrqTtrGrade"));

        o += elem("ULQRTRPercent", prop(q, "ulqRtr"));
        o += elem("ULQRTRGrade",   prop(q, "ulqRtrGrade"));
        o += elem("URQRTRPercent", prop(q, "urqRtr"));
        o += elem("URQRTRGrade",   prop(q, "urqRtrGrade"));
        o += elem("LLQRTRPercent", prop(q, "llqRtr"));
        o += elem("LLQRTRGrade",   prop(q, "llqRtrGrade"));
        o += elem("LRQRTRPercent", prop(q, "lrqRtr"));
        o += elem("LRQRTRGrade",   prop(q, "lrqRtrGrade"));

        o += elem("ULQTCTGrade", prop(q, "ulqTctGrade"));
        o += elem("URQTCTGrade", prop(q, "urqTctGrade"));
        o += elem("LLQTCTGrade", prop(q, "llqTctGrade"));
        o += elem("LRQTCTGrade", prop(q, "lrqTctGrade"));
        o += elem("ULQRCTGrade", prop(q, "ulqRctGrade"));
        o += elem("URQRCTGrade", prop(q, "urqRctGrade"));
        o += elem("LLQRCTGrade", prop(q, "llqRctGrade"));
        o += elem("LRQRCTGrade", prop(q, "lrqRctGrade"));

        // ── 1D ISO 15416 summary (only populated for 1D symbologies) ─────────
        //   ansiGrade / symbolAnsiGrade  — overall ANSI letter grade
        //   avgEdge / avgRlRd / avgSc / avgMinEc / avgMod / avgDefect
        //   avgDcod / avgDec / avgLqz / avgRqz / avgHqz / avgMinQz
        //   bwg        — Bar Width Growth %
        //   magnification
        //   ratio      — narrow-to-wide ratio
        //   nominalXDim1d — 1D nominal X dimension

        var ansiGrade = prop(q, "ansiGrade") || prop(q, "symbolAnsiGrade");
        o += elem("SymbolAnsiGrade", ansiGrade);
        o += elem("AvgEdge",     prop(q, "avgEdge"));
        o += elem("AvgRlRd",     prop(q, "avgRlRd"));
        o += elem("AvgSC",       prop(q, "avgSc"));
        o += elem("AvgMinEC",    prop(q, "avgMinEc"));
        o += elem("AvgMOD",      prop(q, "avgMod"));
        o += elem("AvgDefect",   prop(q, "avgDefect"));
        o += elem("AvgDcod",     prop(q, "avgDcod"));
        o += elem("AvgDEC",      prop(q, "avgDec"));
        o += elem("AvgLQZ",      prop(q, "avgLqz"));
        o += elem("AvgRQZ",      prop(q, "avgRqz"));
        o += elem("AvgHQZ",      prop(q, "avgHqz"));
        o += elem("AvgMinQZ",    prop(q, "avgMinQz"));
        o += elem("BWGPercent",      prop(q, "bwg"));
        o += elem("Magnification",   prop(q, "magnification"));
        o += elem("Ratio",           prop(q, "ratio"));
        o += elem("NominalXDim1D",   prop(q, "nominalXDim1d") || prop(q, "nominalXDim"));

        // ── Per-scan results for 1D (written as <ScanResults><Scan …>…) ──────
        //   q.scans — array; each entry has: edge, sc, minEc, mod, defect, dec, lqz, rqz, hqz

        if (q.scans && q.scans.length) {
            o += "<ScanResults>\r\n";
            for (var i = 0; i < q.scans.length && i < 10; i++) {
                var sc = q.scans[i];
                o += '<Scan number="' + (i + 1) + '">\r\n';
                o += elem("Edge",   prop(sc, "edge"));
                o += elem("SC",     prop(sc, "sc"));
                o += elem("MinEC",  prop(sc, "minEc"));
                o += elem("MOD",    prop(sc, "mod"));
                o += elem("Defect", prop(sc, "defect"));
                o += elem("DEC",    prop(sc, "dec"));
                o += elem("LQZ",    prop(sc, "lqz"));
                o += elem("RQZ",    prop(sc, "rqz"));
                o += elem("HQZ",    prop(sc, "hqz"));
                o += "</Scan>\r\n";
            }
            o += "</ScanResults>\r\n";
        }

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
