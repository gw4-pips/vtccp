# Standards Library Index

ISO/IEC standards governing barcode symbol print quality, verifier
conformance, and symbology specifications. **VTCCP is a logging/export
utility, not a verifier or a decoder** — these standards inform what the
underlying verifier's reported values *mean*, so we can present and
preserve them correctly.

---

## Filed standards

| File | Standard | Purpose | Digest |
|---|---|---|---|
| `ISO-IEC-15415-ed3-2024-print-quality-2D-DRAFT.pdf` | **ISO/IEC 15415:2024** (3rd ed, DIS) | 2D barcode print quality methodology — the parameters our verifier computes | `ISO-IEC-15415-ed3-2024-digest.md` |
| `ISO-IEC-15426-2-2023-verifier-conformance-2D.docx` | **ISO/IEC 15426-2:2023** (3rd ed) | Verifier conformance — what makes hardware a conforming 2D verifier | `ISO-IEC-15426-2-2023-digest.md` |

Both files are "very close to but not the current published versions" per
project owner. Should be substantively identical to published versions for
our reference purposes; minor wording / table-numbering drift is possible.

---

## Outstanding (not in library yet)

| Standard | Purpose | Priority |
|---|---|---|
| **ISO/IEC 15416** | 1D / linear bar code print quality methodology — the source of all our currently-empty `Avg*` Excel columns | **High** — needed to interpret 1D scan output |
| **ISO/IEC 29158** | Direct Part Marking (DPM) print quality methodology — AIM-DPM mode | **High** — needed for the DPM-only metrics in our 30-key enumeration (`cellDefects`, `finderPatternDefects`, mark misplacement, cell dims) |
| **ISO/IEC 16022** | Data Matrix symbology spec (incl. reference decode algorithm) | Medium — useful for understanding the corners/center/size geometry that v1.24 will probe |
| **ISO/IEC 18004** | QR Code symbology spec (incl. reference decode algorithm) | Medium — needed for QR Code support phase |
| **ISO/IEC 15426-1** | Verifier conformance for 1D symbols (sister to 15426-2) | Low — completeness only |
| **ISO/IEC 15438** | PDF417 symbology spec | Low — niche symbology |
| **ISO/IEC 24778** | Aztec Code symbology spec | Low — niche symbology |
| **ISO/IEC 19762** | AIDC harmonised vocabulary (terminology reference) | Low — definitions reference |

If/when these are acquired, file the same way: PDF/DOCX in this directory
with a descriptive name, plus a sibling `.md` digest using the same format
as the existing digests.

---

## Cross-reference: standards ↔ DMST fields

A summary of how the ISO 15415 mandatory parameters map to the DMST XML
fields we see in v1.23 scans (full detail in `ISO-IEC-15415-ed3-2024-digest.md` §11):

| 15415 parameter | DMST XML | Status |
|---|---|---|
| Decode | `<DecodeGrade>` | wired |
| Symbol Contrast | `<SCPercent>`, `<SCGrade>`, `<SCRlRd>` | wired |
| Modulation | `<MODGrade>` | wired (grade only) |
| Fixed Pattern Damage | `<FPDValue>`, `<FPDGrade>` | wired |
| Axial Nonuniformity | `<ANUPercent>`, `<ANUGrade>` | wired |
| Grid Nonuniformity | `<GNUPercent>`, `<GNUGrade>` | wired |
| Unused Error Correction | `<UECPercent>`, `<UECGrade>` | wired |
| Reflectance Margin | `<RMGrade>` | wired (grade only) |
| Print Growth | `<BWGPercent>` (empty) | **gap — wire from `r.metrics.printGrowth` in v1.24** |
| Contrast Uniformity | `<ContrastUniformity>` | wired |
| Reference grading string | `<FormalGrade>1/D</FormalGrade>` | wired (legacy form); standard prefers `Grade/Aperture/Wavelength/Lighting` (e.g. `1.0/17/660/45Q`) |
| Application minimum pass grade | (not directly emitted) | **gap — wire from `r.metrics.minPassGrade` in v1.24** |

---

## Important framing

**We are not building a verifier.** The standards make it explicit (15426-2
Annex B): primary verification requires NIST-traceable instruments **10x
better** than the commercial verifier under test. The DM475V is the
commercial verifier; VTCCP is the data layer downstream of it.

The standards are reference material to:
1. **Validate** that our schema covers what conformant verifiers must
   report (and we know what's mandatory vs optional)
2. **Document** what each numeric value in our output *means*, in
   vendor-neutral terms
3. **Justify** our `OpticsCompliant` flag and field-suppression rules with
   formal grounding (15426-2 §5 tolerances apply only to verifier-captured
   measurements, not loaded images)
4. **Differentiate** from vendor-specific UIs — using standardized term
   names (SC, MOD, FPD, ANU, GNU, UEC, RM, PG) makes our trade-dress
   posture stronger
