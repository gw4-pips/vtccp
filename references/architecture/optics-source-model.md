# VTCCP Optics Source Model

**Created**: 2026-05-21
**Status**: Authoritative design reference

---

## Purpose

Every verification scan record in VTCCP carries an `OpticsSource` indicator that
describes the origin of the image the DataMan verifier graded. This matters because
ISO/IEC 15426-2 conformance applies only when the verifier's own calibrated optical
chain — controlled illumination, known aperture, known wavelength — produced the
image. Loaded-image scans may or may not meet those conditions, and the distinction
must be disclosed in any formal quality record.

---

## OpticsSource enum values

| Value | Meaning |
|---|---|
| `LiveScan` | The verifier captured and graded a live symbol through its own calibrated optics. Fully ISO/IEC 15426-2 conformant. |
| `LoadedImage` | A JPEG was submitted to the verifier for analysis via IMAGE.LOAD / IMAGE.REPLAY. Origin is undocumented or known to be non-DataMan. |
| `LoadedImageTruCheck` | A JPEG submitted for analysis is documented as having originated from a prior DataMan TruCheck live verification scan. |

---

## Disclosure rules

### LiveScan
No disclaimer required. The result is produced under the verifier's calibrated optical
chain and falls within the ISO/IEC 15426-2 conformance tolerance regime. Full
compliance assertion is valid.

### LoadedImage (undocumented or non-DataMan origin)
**Disclaimer required** in any formal quality record. Rationale: the image most
often will not have originated from a DataMan verifier. It therefore cannot be
assumed to meet ISO/IEC stipulated illumination constraints (controlled 45°
annular or on-axis illumination, calibrated wavelength, known aperture). The
verifier's grading algorithms are operating on optics it did not control.

Suggested disclosure language:
> *"Loaded-image scan — image origin not from a DataMan TruCheck live verification
> event. Results fall outside ISO/IEC 15426-2 optical conformance conditions and
> should not be used as the sole basis for a formal quality declaration."*

### LoadedImageTruCheck (documented prior DataMan TruCheck scan)
**No ISO/IEC disclaimer required.** The image originated from the same class of
calibrated DataMan optics; the optical chain is equivalent to a live scan. However,
an appropriate note must appear indicating this is a **secondary DataMan TruCheck
grading event** — i.e., the symbol was not re-presented to the verifier live; the
original scan's image was re-submitted.

Suggested note language:
> *"Secondary DataMan TruCheck grading event — results derived from a loaded image
> originating from a prior DataMan TruCheck live verification scan."*

---

## OpticsCompliant flag (derived)

`OpticsCompliant` is a boolean derived from `OpticsSource`:

| OpticsSource | OpticsCompliant |
|---|---|
| `LiveScan` | `true` |
| `LoadedImage` | `false` |
| `LoadedImageTruCheck` | `true` (with secondary-event note) |

This flag is what surfaces the disclaimer logic in report generation (D1) and is
the field cited in the ISO 15426-2 digest's OpticsCompliant note.

---

## Firmware observation (DM475V fw 6.1.16_sr4)

**All live QR Code scans on this firmware return `OpticsSource = LoadedImage`.**
This is a firmware behavior, not a stored-image indicator. The push XML field
that drives OpticsSource inference must not be used alone to distinguish a true
stored-image scan from a QR live scan on this firmware version. The discriminator
for a genuine IMAGE.LOAD QR scan remains an open probe item (D4 scope).

---

## Data model wiring

- `VerificationRecord.OpticsSource` — string, values as above
- `VerificationRecord.OpticsCompliant` — bool?, derived at parse time
- Both fields written to Excel schema (Universal columns)
- Report renderer (D1) reads `OpticsCompliant` to conditionally emit disclaimer block

---

## Cross-references

- ISO/IEC 15426-2 digest: `references/standards/ISO-IEC-15426-2-2023-digest.md`
  (Table 1 tolerances apply only under conformant optical conditions)
- D4 (Image-load implementation): discriminator for true IMAGE.LOAD vs. firmware
  QR behavior is an open item
- Session plan: `LoadedImageTruCheck` sub-type must be selectable by the operator
  at scan time or via post-scan tagging (UI design TBD at D4 scope)
