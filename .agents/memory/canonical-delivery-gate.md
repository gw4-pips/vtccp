---
name: Canonical delivery gate
description: Required VTCCP release workflow for preventing locally validated but unpublished builds
---
VTCCP delivery is canonical only after Build → Push → Confirm. The build must succeed, the working tree must be committed, the exact local HEAD must be pushed, and GitHub's fetched branch tip plus application version must be verified.

**Why:** A local checkpoint can contain a valid build while the GitHub branch remains at an older commit, causing downstream Windows workstations to receive stale code.

**How to apply:** Run `bash scripts/build-push-confirm.sh` for VTCCP delivery. Treat any nonzero exit or SHA mismatch as unpublished; do not call the build delivered.