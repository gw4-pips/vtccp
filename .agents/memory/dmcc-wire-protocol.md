---
name: DMCC raw TCP wire protocol
description: Correct wire format for raw TCP DMCC commands — trigger syntax, header, response modes
---

## Rules

**Command format**: `||>{COMMAND} {ARGUMENT}\r\n`
- `||>` is the bare header (full form: `||checksum:command-id>` — all fields optional)
- Every raw TCP command must include this prefix; bare `COMMAND\r\n` is silently ignored

**Trigger command**: `||>TRIGGER ON\r\n`
- Argument is required: `ON` fires the trigger, `OFF` cancels a held manual trigger
- `TRIGGER\r\n`, `TRIGGER 1\r\n` are wrong — device silently ignores them
- The Cognex SDK also throws `InvalidParameterException` for `SendCommand("TRIGGER")` and `SendCommand("TRIGGER 1")`
- Try `SendCommand("TRIGGER ON")` via SDK first — may work since it's the correct form

**Response modes** (set via `SET COM.DMCC-RESPONSE`):
- Mode 0 (Silent, default): **no ACK for any command** — device sends zero bytes; TRIGGER ON appears to timeout
- Mode 2 (Extended): every command returns `||[status]\r\n`
  - `||[0]\r\n` = OK
  - `||[101]\r\n` = invalid command
  - `||[102]\r\n` = invalid parameter
  - `||[104]\r\n` = parameter rejected due to reader state

**Required raw TCP sequence**:
1. Connect → drain welcome banner (≤600ms read)
2. `||>SET COM.DMCC-RESPONSE 2\r\n` → read ACK `||[0]\r\n` (≤1000ms)
3. `||>TRIGGER ON\r\n` → read ACK `||[0]\r\n` (result arrives via HTTP push)

**Why:** The welcome banner drain and SET COM.DMCC-RESPONSE 2 step were omitted in the first raw TCP implementation, and `TRIGGER` was sent without the `ON` argument and without the `||>` header. All three errors combined to produce zero bytes from the device and no scan.

**How to apply:** Any future raw TCP command sequence must follow steps 1-3. The constants are in `DmccCommand`: `WireHeader`, `SetDmccResponseExtended`, `TriggerOn`, `TriggerOff`. `DmccResponse.Parse` handles both the `||[N]\r\n` wire format and the SDK-synthesised `\r\nN\r\n\r\nbody` format.
