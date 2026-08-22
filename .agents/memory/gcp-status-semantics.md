---
name: GCP validation status semantics
description: Rules for distinguishing registered-prefix validation outcomes in RFID records and reports.
---

# GCP validation status semantics

Treat an absent GCP-table lookup as **NOT FOUND**, not as an invalid prefix.
Use **Invalid** only when a prefix is present in the GS1 table but its registered
length disagrees with the length encoded by the SGTIN partition. When known, display
the partition-derived length alongside the result, for example `Valid (=7)`.

**Why:** An unlisted prefix and an encoded-length mismatch are materially different
operational conditions. Collapsing them into a Boolean `false` makes the report
misleading and prevents operators from distinguishing an incomplete table from an
actual allocation mismatch.

**How to apply:** Preserve the explicit status through RFID validation, record
creation, the RFID log, and report rendering. A legacy Boolean may remain only as
a compatibility projection; it must not turn NOT FOUND into Invalid.