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

**Artifact retention:** Keep original DMST HTML reports in the configured
directory after parsing and VCCS PDF generation. The report-processing path
must never delete a source HTML artifact.

**Why:** The original HTML is the verifier provenance artifact and must remain
available for audit, troubleshooting, and visual review.