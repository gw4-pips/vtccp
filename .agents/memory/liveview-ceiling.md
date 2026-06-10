---
name: DM475V live view ceiling
description: Confirmed live view frame rate limits and dead paths on fw 6.1.16_sr4
---

## Rule
LIVEIMG.SEND does not work on DM475V fw 6.1.16_sr4. The only viable unencrypted live
view path is TRIGGER ON + IMAGE.SEND (~1.5-2.5 fps). DMST's 3.5 fps uses an AES-encrypted
HTTP endpoint (GET /svg_image.img on port 44444) whose key is baked into the DMST binary.

## What was tested (2026-06-10, confirmed by direct TCP probe)
- LIVEIMG.MODE=2 + LIVEIMG.SEND 0 1 85, port 23: 0 bytes
- LIVEIMG.MODE=3 + LIVEIMG.SEND 0 1 85, port 23: 0 bytes
- LIVEIMG.MODE=3 + LIVEIMG.SEND 0 1 85, port 44444: 0 bytes
- GET /svg_image.img (HTTP, port 44444): works, returns ~106KB AES body per frame at ~3.5fps
- AES key: not on wire, embedded in DMST binary, not recoverable without binary reverse engineering

## Why
DMST has a privileged encrypted channel. The DMCC reference documents LIVEIMG.SEND
(fw 5.6.3+) but it is non-functional on this firmware/hardware combination regardless
of port or mode. The device accepts SET LIVEIMG.MODE silently without error but LIVEIMG.SEND
returns nothing.

## How to apply
Do not attempt LIVEIMG.SEND in future firmware versions without first confirming it works.
If a Cognex contact offers the AES key for svg_image.img, the HTTP polling path (~3.5fps)
is already understood and straightforward to implement. Without the key, do not invest
further in live view frame rate improvement beyond the current GetFreshFrameAsync approach.
Full analysis: references/architecture/dmst-live-view-protocol-analysis.md
