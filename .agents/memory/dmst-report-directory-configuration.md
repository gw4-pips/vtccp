---
name: DMST report directory configuration
description: Rule for locating local DMST HTML reports used in strict PDF provenance.
---

Use an explicitly configured DMST quality-report directory. Do not derive the
directory from `DeviceInfo.Name`, a displayed profile name, the Windows user
documents folder, or another presumed naming convention.

**Why:** The live installation's DMST reporting directory is independent of
the device identity. An inferred path causes the watcher to miss the real HTML
artifact and strict report generation correctly refuses to create a PDF.

**How to apply:** When the report directory changes, obtain the exact path from
the operator and update the explicit configuration before relying on local HTML
provenance. A future user-facing settings control should write that explicit
value; it must not recreate automatic device-name path derivation.

Use one configured session-evidence root per verifier mode by default. For
Webscan, the directory configured in TruCheck for HTML/image exports should also
be VTCCP's monitored input directory and the default destination for generated
VCCS PDFs and Excel logs. Apply the same model to DataMan/DMST: its explicitly
configured quality-report directory is the monitored input and default VTCCP
output root.

**Why:** Operators should not need to search multiple unrelated folders to
assemble one verification record. Co-locating native verifier evidence and
VTCCP outputs makes each session easier to audit and reduces path mismatch
failures.

**How to apply:** First verify that the native report directory is writable and
suitable for VTCCP outputs. If a site requires separation, expose explicit
incoming-evidence and generated-output overrides in VeriWedge settings, ideally
under advanced path controls. Never silently invent or infer a fallback path.

**Artifact retention:** Keep original DMST HTML reports in the configured
directory after parsing and VCCS PDF generation. The report-processing path
must never delete a source HTML artifact.

**Why:** The original HTML is the verifier provenance artifact and must remain
available for audit, troubleshooting, and visual review.