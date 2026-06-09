---
name: Cognex TruCheck Chief Engineer Strategy Intelligence
description: Key strategic facts from direct conversation with Cognex TC chief engineer (2026-06-08); shapes VTCCP roadmap and 395V migration planning
---

Full detail: `vtccp/architecture/cognex-strategy-intel-2026-06-08.md`

**Why this matters for every future session:**

- **DM475V is frozen** — one more firmware update then EOL for feature dev. All 475V probes must complete now. Do not plan features that need new firmware support.
- **DM395V fw7 is the new standard** — everything confirmed on fw 6.1.16_sr4 must be re-validated. First PIPS unit (OMNI model) arriving within weeks of 2026-06-08. Define probe campaign before unit arrives.
- **No CNX standalone web TC app** — TC will be embedded in DST web UI. VTCCP standalone positioning is explicitly distinct and not contested.
- **CNX output direction** — HTML/XML with Excel-friendly headers, NOT native Excel. Shapes D1 priority over further ExcelEngine investment.
- **UPC/EAN supplementals** — business decision at Webscan→Cognex migration, not a technical block. Open probe: firmware-suppressed (a) vs output-ignored (b). If (b), VTCCP can do its own supplemental analysis. Engage new Cognex PM to make supplementals activatable on 395V.
- **CNX moving to browser** — Windows app currently stronger; VTCCP should complete Windows build now, then reassess browser migration timing against next major Win platform hardware refresh.

**Firm scope decisions logged 2026-06-09 (§7 of strategy doc):**
- **USB connection eliminated** — GigE (Ethernet) only across all supported devices; business/eng resource decision, not a technical limit. No USB-COM or USB-Ethernet adapter paths needed.
- **DMV-8072V is fw5, nice-to-have** — deprioritized; do NOT pre-engineer multi-model abstractions speculatively. Revisit only when real customer demand materialises.
- **Priority order**: finish DM475V → explore DM395V on arrival → plan changes → 8072V last.

**How to apply:**
- When scoping new VTCCP features, check against 395V transition timeline.
- When the first 395V arrives, run probe campaign before any integration work.
- Prioritize D1 HTML/XML report over new Excel sheet development.
- Track supplemental probe outcome; update when test results in.
- Do not add USB transport paths or 8072V-specific command set abstractions without explicit go-ahead.
