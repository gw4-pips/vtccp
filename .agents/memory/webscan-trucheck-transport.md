---
name: Webscan TruCheck transport
description: Canonical physical connection model for Webscan TruCheck devices.
---

All Webscan TruCheck devices are USB-connected devices. Do not ask for, infer,
or probe a device IP address or TCP port when working with this verifier family.

Webscan TruChecks also have a separate result path from the DataMan DMST/HTTP
integration. The existing DataMan HTTP subscriber, its event protocol, and its
diagnostic capture files are not a Webscan adapter. The Webscan result/export
path must be identified and verified independently before implementing against
it.

**Why:** The user explicitly corrected an attempted TC-829 network preflight
on 2026-08-22, then clarified that the DMST/HTTP path is DataMan-only. Network
or result-delivery assumptions would test the wrong interface and can misdirect
a controlled validation.

**How to apply:** Validate physical presence through the local Windows USB/PnP
inventory, then use the installation's Webscan-specific result/export workflow
for a real scan. Keep the DataMan HTTP implementation scoped to DataMan and
never treat it as a Webscan result path.