# DMCC Reference — HTML Pages Index

Source: Cognex DataMan Control Commands Reference, fw 6.1.16_sr4
Format: MadCap Flare WebHelp2 HTML export
Filed: 2026-05-29

These are the original HTML pages from the DMCC Reference documentation set.
Uploaded in batches by the user (zipping was not possible on the source machine).

---

## Files received

### Batch 1 — 2026-05-29

| File | Topic | Notes |
|---|---|---|
| `DataMan_Control_Commands_Overview.htm` | DataMan Control Commands — top-level overview | Pre-existing from earlier session |
| `Action.overview.htm` | Action Commands — overview | |
| `Camera.overview.htm` | Camera Commands — overview | |
| `Code_Quality.overview.htm` | Code Quality Commands — overview | |
| `Communication.overview.htm` | Communication Commands — overview | Largest file (1232 lines) — TCP/IP, serial, HTTP settings |
| `Data_Formatting.overview.htm` | Data Formatting Commands — overview | |
| `data-formatting-tokens.htm` | Data Formatting Tokens — full reference | Large file (1589 lines) — output format token definitions |
| `Data_Validation.overview.htm` | Data Validation Commands — overview | |
| `Decoder.overview.htm` | Decoder Commands — overview | |

---

## Files pending (batches 2+)

The full DMCC folder contains 100+ files. Remaining batches to be received.
Expected topics (based on known DMCC Reference structure):

- Input/Output commands
- System commands
- Trigger commands  ← HIGH PRIORITY for trigger investigation
- Image commands (IMAGE.LOAD, IMAGE.SEND, IMAGE.REPLAY)
- Code Quality detail pages (individual command entries)
- Communication detail pages
- Decoder detail pages (TRIGGER.TYPE, GET/SET individual pages)

---

## Priority files to watch for in upcoming batches

| File pattern | Why needed |
|---|---|
| `Trigger.*` or `trigger*` | TRIGGER.TYPE values, TRIGGER command syntax — critical for trigger investigation |
| `IMAGE.*` | IMAGE.LOAD / IMAGE.REPLAY for D4 batch upload feature |
| Any file containing `TRIGGER.TYPE` | Confirm exact SET/GET syntax and value table |
