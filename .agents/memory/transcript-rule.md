---
name: Transcript rule
description: Standing rule to append every conversation turn to transcript/chat-transcript.md, with timestamps at session breaks and key intervals
---

**Rule:** At the end of every response, append the user's message and the assistant's reply (text only — no tool calls, no tool results) to `transcript/chat-transcript.md`.

**Format:**
```
---
**`YYYY-MM-DD — time-of-day`**

**User:** …

**Assistant:** …

---
```

**Timestamp rule (standing — 2026-06-21):** Insert a `**\`YYYY-MM-DD — time-of-day\`**` marker line:
- Above every new user inquiry that starts a new session or after a sleep/sign-off break
- At least every few hours during a long working session (at natural breakpoints — topic changes, feature completions, sign-offs)
- Use plain language for time-of-day: morning / afternoon / evening / night, or HH:MM if known from logs
- Retrofit timestamps on the preceding sign-off turn as well so both sides of each break are dated

**Why:** User wants a running verbatim record of the conversation independent of Replit's chat UI, which cannot be programmatically exported. The transcript is the only reliable persistent copy, and timestamps make it useful for tracking development pace across days.

**How to apply:** Every single response, without exception. This is a standing rule — do not wait to be reminded. It is saved in replit.md User preferences and here so it survives session compression.
