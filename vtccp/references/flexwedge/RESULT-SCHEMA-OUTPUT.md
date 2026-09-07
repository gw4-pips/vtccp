# Result schema and output adapter

**Document version:** 1.0.0  
**Revision date:** 2026-09-06

`FlexWedgeReadResult` is the destination-independent record. It retains raw EPC/source, decoded scheme, GTIN-14, serial, filter, partition, company prefix, item reference, EPC URI, GCP status and registered length, reader audit (timestamp, COM port, RSSI, PC, TID, lock), warnings, and named bit ranges with start/end offsets. CSV, JSONL, and the automatic per-user JSONL session record retain those audit fields.

`FlexWedgeOutputProfile` selects ordered canonical fields, user-owned headers, prefix, suffix, delimiter, and append key. The Complete decoded result profile is the full release mapping. Version 1 adapters are Clipboard and Previously Focused Application. A foreground watcher retains the actual external window, and delivery confirms that window regained focus before pasting; it has no hardcoded spreadsheet coordinates. CSV and JSONL are session export formats, not interactive destinations.