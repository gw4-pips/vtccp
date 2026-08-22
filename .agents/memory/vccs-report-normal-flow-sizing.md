---
name: VCCS report normal-flow sizing
description: Prevent blank trailing PDF pages while allowing long VCCS reports to flow naturally.
---

For the VCCS HTML report, do not give the page container a fixed or minimum physical paper height when the report must support natural multi-page flow.

**Why:** A flex page container with a Letter-height minimum can push its footer onto an otherwise blank second PDF page, even when the report body itself fits on the first page.

**How to apply:** Keep page sizing content-driven and use the report-section/report-block non-splitting rules to preserve section integrity. Render a representative PDF after changing layout CSS, especially when changing vertical flow or footer placement.