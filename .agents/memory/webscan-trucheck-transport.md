---
name: Webscan TruCheck transport
description: Canonical physical connection model for Webscan TruCheck devices.
---

All Webscan TruCheck devices are USB-connected devices. Do not ask for, infer,
or probe a device IP address or TCP port when working with this verifier family.

The confirmed TC-829 result path is a local native Webscan HTML export, with a
sibling symbol image retained as raw evidence. It is separate from the DataMan
DMST/HTTP integration. The existing DataMan HTTP subscriber, its event
protocol, and its diagnostic capture files are not a Webscan adapter.

**Why:** A controlled TC-829 export established literal Webscan report fields
and a file-based result workflow. Treating it as DataMan transport would test
the wrong interface and can misdirect a controlled validation.

**How to apply:** Validate physical presence through the local Windows USB/PnP
inventory, then use a new native Webscan HTML export for a real scan. Preserve
the HTML and sibling image without editing either. Keep DataMan HTTP scoped to
DataMan; reject incomplete Webscan exports rather than fabricating values.