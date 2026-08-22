---
name: Webscan TruCheck transport
description: Canonical physical connection model for Webscan TruCheck devices.
---

All Webscan TruCheck devices are USB-connected devices. Do not ask for, infer,
or probe a device IP address or TCP port when working with this verifier family.

**Why:** The user explicitly corrected an attempted TC-829 network preflight
on 2026-08-22. Network assumptions would test the wrong interface and can
misdirect a controlled validation.

**How to apply:** Validate physical presence through the local Windows USB/PnP
inventory and use the established DMS-linked result workflow for a real scan.
Treat any HTTP exchange as an integration path around the local software, not
as a claim that the Webscan hardware itself is network-addressable.