# Cognex TruCheck Chief Engineer — Strategy Intelligence
**Date**: 2026-06-08  
**Source**: Direct conversation with Cognex TruCheck chief engineer (Webscan founder)  
**Context**: Strategic alignment session; no NDA implications noted

---

## 1. DM475V Platform Status — EOL for Feature Development

- **One more firmware update expected** for the DM475V, then feature development ends.
- Product is approaching end-of-life from Cognex's engineering investment perspective.
- **VTCCP implication**: treat the 475V firmware as a frozen target.
  - All probe campaigns should be completed on current fw 6.1.16_sr4 *now*.
  - Do not plan VTCCP features that depend on firmware fixes or new DMCC keys on this platform.
  - The v1.33 probe campaign being complete before this EOL signal is well-timed.

---

## 2. Platform Direction — DM395V at Firmware Level 7

- **DM395V is the new standard platform**, shipping at firmware level 7.
- First PIPS unit expected: 395V OMNI model, **within weeks** (OPTI model running behind due to complex molding).
- Firmware 7 = potentially new DMCC key space, new push script schema, new connection architecture.
- **Everything confirmed on fw 6.1.16_sr4 requires re-validation on the first 395V**:
  - COM.DMCC-RESPONSE default behaviour
  - Port-23 result push mechanics
  - IMAGE.SEND wire protocol
  - Push script schema and probe campaign (new v-series script needed)
  - HTTP pub/sub channel (port 44444 — confirm same architecture)
- **Action**: Define the 395V probe campaign before the first unit arrives so we can
  execute systematically on day one.

---

## 3. UI Platform Direction — Browser/Web Migration

- Cognex is moving **aggressively away from Windows app toward browser/web interface**.
- The Windows app is currently stronger (more capable), but CNX's investment direction is browser.
- **No plans for a standalone DM TC web app**: TruCheck UI will be embedded as a section
  of the larger DST (DataMan Setup Tool) web interface — not a standalone product.
- Chief engineer has **no objection** to a third party building a standalone DM TC interface
  (VTCCP's exact positioning).
- **VTCCP strategic response**:
  - Build out the Windows app fully now — you have the device, the active user base,
    and firmware access. The Windows app *is* the design specification for any future browser port.
  - Explore the 395V browser interface thoroughly before deciding on the browser migration path.
  - Target transition point: before the next *major* Windows platform hardware refresh cycle.
  - Do not chase the CNX web interface until it is stable enough to design against.

---

## 4. Output Architecture — No Excel-Native Capture Planned

- Cognex has **no plans to build Excel data capture** into TruCheck.
- Direction: export HTML/XML with headers structured to be Excel-import-friendly.
- **VTCCP implication**:
  - ExcelEngine work already shipped for 475V — delivers real value for current users.
  - For 395V integration, prioritize the **D1 HTML/XML report path** over additional
    Excel sheet investment.
  - Align VTCCP output architecture with the CNX export direction: HTML/XML as the
    primary output format, with Excel as a secondary derived output.

---

## 5. UPC/EAN Supplemental Support — Business Decision, Not Technical Constraint

- Lack of UPC/EAN supplemental support is a **business-driven design decision** made at
  the Webscan-to-Cognex migration, not a technical limitation.
- Market assessment at the time: insufficient ROI for engineering investment.
- **Open probe required**: determine whether supplementals are:
  - **(a) Firmware-suppressed**: verifier hardware/firmware cannot decode them at all, or
  - **(b) Output-ignored**: verifier *can* decode supplementals but the grading logic
    ignores them and they are absent from the output
  - If (b): VTCCP could perform its own supplemental analysis on the decoded barcode
    content without any firmware changes.
- **Action**: Run supplemental activation probe on the first available unit (475V now,
  395V when it arrives). Test with a UPC label that has a 2- or 5-digit add-on.
- **Advocacy target**: engage new product manager at Cognex to include supplemental
  support in the 395V as the only realistic path to restoring this capability.
  - Minimum ask: make supplementals *activatable* for third-party analysts even if
    no supplemental grading logic is added to DM TC itself.

---

## 6. Competitive / Ecosystem Context

- The TruCheck product has a clear CNX-stated gap: no standalone web app planned, TC
  is subordinated to the DST web UI. VTCCP's standalone positioning is therefore
  not just tolerated but strategically distinct from anything CNX will ship.
- DM395V OMNI arriving at PIPS first; OPTI model behind due to molding complexity.
- Timeline for 395V adoption: PIPS will be early adopter; use this window to validate
  VTCCP on 395V before broader rollout.

---

## 7. Engineering Scope Decisions — 2026-06-09

Firm decisions from follow-up conversation. These narrow the device matrix and simplify
the multi-model architecture significantly.

### USB Connection — Eliminated Across the Board
- USB (USB-COM and USB-Ethernet adapter) is **not a supported connection interface** for VTCCP.
- **Decision basis**: business and engineering resource triage — not a technical limitation.
- **Impact**: no `SerialDmccTransport` needed; no USB-COM session setup path; no USB variant
  testing required. All device support is GigE (Ethernet) only.
- All connection medium logic in `DeviceConfig.ResolvedConnectionMedium()` can treat USB-Ethernet
  as a GigE variant (same port-23/44444 behaviour) and USB-COM as out of scope.

### DMV-8072V — Nice-to-Have, Not Essential
- **Firmware**: firmware 5 (one major version behind DM475V fw6; two behind DM395V fw7).
- **Priority**: deprioritized — nice to have, not on the critical path.
- **Impact on architecture**: do NOT pre-engineer `IDmccCommandSet` abstraction layers for
  8072V now. Design for 475V + 395V; add 8072V when/if it becomes a real customer requirement.
  Engineering the abstraction speculatively before confirming 8072V's actual DMCC surface would
  be wasted effort — assume less, prove more.
- 8072V fw5 DMCC command surface is unconfirmed. Log what is known (DPM primary use case,
  more limited DMCC set) and revisit when a unit is available.

### Development Priority Order — Confirmed
1. **Finish DM475V** (fw6 / GigE) — complete all outstanding features including IMAGE.SEND
   live view, D1 report, supplemental probe, and any remaining v1.30 bug fixes.
2. **Explore DM395V** (fw7 / GigE) when first unit arrives — run probe campaign,
   confirm connection protocol, push schema, IMAGE.SEND, and DMCC key differences.
3. **Plan enhancements/necessary changes** based on 395V probe results before any 395V build.
4. **DMV-8072V** — revisit only when a real customer or use-case demand materialises.

---

## 8. Recommended Near-Term Actions (Priority Order)

1. **Complete 475V IMAGE.SEND debugging** — get the live view working on current hardware
   before the firmware freezes.
2. **Run supplemental activation probe** on 475V — determine firmware-suppressed vs
   output-ignored.
3. **D1 HTML/XML report** — move up in priority relative to additional Excel investment;
   aligns with CNX output direction.
4. **Define 395V probe campaign** — document all probes needed on first 395V contact
   (connection protocol, push schema, firmware 7 DMCC keys, IMAGE.SEND behaviour).
5. **395V browser interface exploration** — when first unit arrives, document the browser
   UI thoroughly before making browser migration architectural decisions.
6. **PM engagement at Cognex** — supplemental activatability ask for 395V.
7. **DMV-8072V** — defer until real customer demand; revisit fw5 DMCC surface at that time.
