---
name: ExcelEngine circular dependency rule
description: Which project sub-writers that touch DeviceInterface types must live in — prevents circular reference build failure
---

DeviceInterface.csproj → ExcelEngine.csproj (one-way dependency).

**Rule:** Any Excel sub-writer class that references types from `DeviceInterface.Rfid.*` (or any other DeviceInterface namespace) must be placed in the `DeviceInterface` project, not in `ExcelEngine.Writer`.

**Why:** ExcelEngine cannot reference DeviceInterface — that would create a circular project reference and fail to compile. The dependency arrow is DeviceInterface → ExcelEngine, never reversed.

**How to apply:** Before adding a new sub-writer to `ExcelEngine/Writer/`, check whether it needs to import any `DeviceInterface.*` namespace. If yes, put the writer in `DeviceInterface/Rfid/` (or appropriate DeviceInterface subfolder) instead. It can still use `IExcelAdapter` from `ExcelEngine.Adapters` and schema constants from `ExcelEngine.Schema` — those references go the correct direction.

Example: `RfidTabWriter.cs` was initially scaffolded in `ExcelEngine/Writer/` but had to be moved to `DeviceInterface/Rfid/` because it imports `DeviceInterface.Rfid.Models.RfidValidationResult`.
