---
name: Transcript rule
description: Standing rule to append every conversation turn to transcript/chat-transcript.md
---

**Rule:** At the end of every response, append the user's message and the assistant's reply (text only — no tool calls, no tool results) to `transcript/chat-transcript.md`.

**Format:**
```
**User:** …

**Assistant:** …

---
```

Group turns under a `## YYYY-MM-DD` date heading. Add the heading if it doesn't already exist for today.

**Why:** User wants a running verbatim record of the conversation independent of Replit's chat UI, which cannot be programmatically exported. The transcript file is the only reliable persistent copy.

**How to apply:** Every single response, without exception. This is a standing rule — do not wait to be reminded. It is saved in replit.md User preferences and here so it survives session compression.
