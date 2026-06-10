---
name: Doc versioning rule
description: User preference — always increment version header in synthesis/architecture docs when editing them.
---

## Rule

Any time a `.md`, `.html`, or other synthesis/design document in `references/` (or elsewhere) is modified, the document's version header must be bumped before delivery.

**How to apply:**
- If the doc already has a version line (e.g. `> v1.1 — 2026-05-18`), increment the minor number and update the date.
- If no version header exists, add one on the first edit: `> v1.0 — YYYY-MM-DD` immediately below the title.
- Never deliver an updated doc to the user without this bump.
- When re-zipping or presenting docs, the archive filename may optionally reflect the version (e.g. `vtccp-feature-summary-v1.2.md`) but the in-file header is mandatory.

**Why:** User explicitly stated this requirement. Versioning lets them track which copy they have open vs. what has changed, especially when downloading archives across multiple sessions.
