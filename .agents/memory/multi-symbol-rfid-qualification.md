---
name: Multi-symbol RFID qualification
description: Working qualification rule for Webscan exports containing three or more independent native symbol reports
---
For a three-or-more-symbol Webscan event, a qualified RFID result requires a connected successful RFID read that matches at least one linear identity and at least one 2D identity. Additional symbols are retained as native evidence and reported independently; an additional identity mismatch does not by itself invalidate the qualified RFID match.

**Why:** A multi-symbol export can contain independent symbols where one linear and one 2D symbol establish the RFID identity relationship, while other symbols still need transparent per-symbol reporting rather than being silently discarded or allowing source-order selection to decide the primary symbol.

**How to apply:** Keep individual symbol pass/fail rows and native grades/images/DFC data visible. Use an asterisked top result such as “QUALIFIED RFID MATCHED*” when the minimum linear-plus-2D RFID relationship is satisfied, with a note explaining that additional symbols remain independently reported.