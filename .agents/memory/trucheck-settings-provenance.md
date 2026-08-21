---
name: TruCheck settings provenance
description: Rules for sourcing Application Standard, Data Format Check, and Aperture Setting in VCCS reports.
---

Application Standard must never be derived from the TruCheck HTML Data Format Check
heading. TruCheck HTML does not publish the active Application Standard selection; a
heading such as “GS1 Application Data Format” describes the DFC result, not the
Application Standard.

**Why:** DMST permits these configuration settings to change during a live verification
session. A session-start read can therefore be stale, and inferring an unpublished setting
from a DFC heading produces a demonstrably wrong report value (for example, Custom + GS1).

**How to apply:** After every completed verification result, read Application Standard,
Data Format Check selection, and Aperture Setting from TruCheck via DMCC before rendering
or persisting the report. If Application Standard or Aperture lacks a known DMCC value,
display `—`; do not reuse a previous scan’s value or infer one from the DMST screen or HTML.
The published HTML DFC result may be a fallback for the DFC display only, never for
Application Standard; absence of an HTML DFC table is not proof of `None`.