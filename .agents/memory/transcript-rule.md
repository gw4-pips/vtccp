---
name: Transcript rule
description: Standing rule to append every conversation turn to transcript/chat-transcript.md, with a date+time timestamp on every single entry
---

**Rule:** At the end of every response, append the user's message and the assistant's reply (text only — no tool calls, no tool results) to `transcript/chat-transcript.md`.

**Format — every entry must include a timestamp:**
```
---

**`YYYY-MM-DD — HH:MM or time-of-day`**

**User:** …

**Assistant:** …

---
```

**Timestamp rule (standing — confirmed 2026-06-24):** Every single entry gets its own `**\`YYYY-MM-DD — time\`**` line — not just at session breaks, not just every few hours. **Every entry. No exceptions.**

- Use HH:MM if known from device timestamps or logs
- Use plain language (morning / afternoon / evening / night) when exact time is unknown
- If multiple consecutive turns happen within the same minute, the same timestamp is fine
- Retrofit missing timestamps when the user points out omissions

**Why:** User wants a running verbatim record with timestamps useful for tracking development pace across days. Agreed standing rule — do not wait to be reminded. Failure to include a timestamp on an entry is a rule violation.

**How to apply:** Every single response, without exception. This rule is saved in MEMORY.md and replit.md so it survives session compression.
