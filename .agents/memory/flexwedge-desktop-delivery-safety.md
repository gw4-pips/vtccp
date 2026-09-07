---
name: FlexWedge desktop delivery safety
description: Durable safety rules for focused-application wedge output and per-user installation data
---

Track the actual external foreground window while FlexWedge is running. Before emitting a paste, restore that exact window and verify it became foreground; never infer the destination from Z-order.

**Why:** Z-order does not identify the previously focused application, so an inferred handle can send EPC data into the wrong program.

**How to apply:** Any future keyboard, clipboard-paste, or focused-cell adapter must fail closed when the stored window is invalid or focus restoration cannot be confirmed.

Install immutable program files separately from mutable configuration and session journals. Upgrades and default uninstall must preserve operator data; destructive purge requires an explicit option.

**Why:** Mixing binaries and user data makes a routine upgrade or uninstall overwrite profiles and destroy retained RFID records.

**How to apply:** Keep executables under the per-user Programs tree and data under the VCCS application-data tree.