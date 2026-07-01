---
name: Session start date/time rule
description: Standing rule to fetch current date/time at the start of every new session before writing any dated content.
---

## Rule

At the start of every new session, before writing any document, transcript entry, or dated content:
1. Use the web search tool to fetch the current date and time.
2. Apply the confirmed date to all documents and transcript entries in that session.
3. If the user provides a date/time explicitly, that is always authoritative — override any fetched value.

## Why

The system-injected "Today is..." date lags behind reality (demonstrated 2026-07-01: system said June 25, actual date was July 1 — a six-day error). Documents and transcript entries dated incorrectly create a misleading audit trail.

## How to apply

- Session start is recognizable: fresh context window, auto_memory block visible, project goal/progress summary present.
- Fetch date/time immediately — before responding to the user's first message if possible, or as the first action in the first response.
- Once anchored, track time within the session from any explicit timestamps the user provides (e.g., "1 JUL 0948").
- Record the confirmed session-start date in the first transcript entry for that session.
