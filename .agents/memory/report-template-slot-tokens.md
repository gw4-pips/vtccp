---
name: Report template slot tokens
description: HTML report replacement tokens must occur only in live markup, never in comments.
---

## Rule

Keep each report replacement token unique in the embedded HTML template. Do not repeat a slot token in documentation comments because string replacement expands every occurrence.

**Why:** Repeated slot tokens silently duplicate generated report sections while leaving the source generator looking correct.

**How to apply:** When adding or moving a template slot, search the entire HTML file for every token occurrence and ensure only the live insertion point remains.