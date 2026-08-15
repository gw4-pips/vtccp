# GCP Prefix Table Update Service

Rev 1.0 — 2026-08-15

Azure-gated auto-update for the GS1 GCP Prefix Format List (~200K entries, ~10 MB).
GS1 updates the table irregularly; workstations self-notify on startup and install
with one click — no manual file copying on operator machines.

Why an Azure Function instead of a raw public blob:
- Audit trail of who downloaded which version and when
- Future per-device / per-license gating without client changes
- Storage credentials never ship in the app binary

## Components

| Component | Location |
|---|---|
| Azure Function (dotnet-isolated, .NET 8) | `vtccp/tools/gcp-update-service/` |
| `gcpEncrypt` CLI | `vtccp/tools/gcp-encrypt/` |
| Client update service | `vtccp/DeviceInterface/Rfid/Gcp/GcpUpdateService.cs` |
| Envelope crypto (shared format) | `vtccp/DeviceInterface/Rfid/Gcp/GcpCrypto.cs` |
| App wiring (factory, toast, settings) | `VtccpApp/Services/GcpUpdateServiceFactory.cs`, `MainViewModel`, `SettingsViewModel` |

## Azure side (one-time setup)

Private blob container `gcp-prefix-table/`:
- `gcpPrefixFormatList.xml.enc` — AES-256-GCM encrypted table (GCP1 envelope, below)
- `gcpMeta.json` — `{ "date": "<root date attribute>", "sha256": "<hex of .enc>" }`

Function endpoints (anonymous auth level; gated by pre-shared device token in the
`X-Device-Token` header, validated against the `DeviceTokens` app setting):

| Method | Route | Returns |
|---|---|---|
| GET | `/api/gcpMeta` | `gcpMeta.json` (tiny; version check) |
| GET | `/api/gcpTable` | the encrypted blob |

Every request writes an audit row to Table Storage (`GcpAudit`):
PartitionKey = UTC date, columns DeviceToken, Endpoint, Ip, Version, TimestampUtc.

App settings: `GcpStorageConnection`, `GcpBlobContainer` (default `gcp-prefix-table`),
`GcpAuditTable` (default `GcpAudit`), `DeviceTokens` (semicolon-separated).

## gcpMeta.json schema

```json
{
  "date":   "2026-06-03T11:14:42.028Z",   // "date" attribute of the GCPPrefixFormatList root
  "sha256": "ab12…ef"                      // lower-case hex SHA-256 of gcpPrefixFormatList.xml.enc
}
```

## GCP1 encryption envelope

AES-256-GCM. Key = `gcpKey.bin`, exactly 32 raw bytes, deployed beside `VtccpApp.exe`
(never uploaded to Azure). File layout of `.enc`:

| Offset | Length | Field |
|---|---|---|
| 0 | 4 | magic `GCP1` (ASCII) |
| 4 | 12 | nonce (random per file) |
| 16 | 16 | GCM authentication tag |
| 32 | … | ciphertext |

Implemented identically in `GcpCrypto.cs` (client) and the `gcpEncrypt` CLI — keep in sync.

## Client flow (VtccpApp)

1. **Startup** (background, non-blocking, never throws): `GcpUpdateService.CheckNowAsync()`
   GETs `/api/gcpMeta` and compares `date` to the root `date` attribute of the local table.
   Silent no-op when Settings → Data Sources is unconfigured or the network is down.
2. **Update found** → non-blocking toast in the main window:
   *"GCP prefix table update available ({date}). Install now?"*
3. **Install** (`DownloadAndInstallAsync()`): GET `/api/gcpTable` → verify SHA-256
   against meta → decrypt with `gcpKey.bin` → sanity-parse via `GcpLengthTable` →
   atomic write to `%AppData%\VCCS\gcpPrefixes.xml` → persist `GcpDataPath` +
   `GcpLastModified` in AppSettings.
4. **Reload**: the in-memory table is loaded at session start (SessionViewModel);
   the next session started after install automatically uses the new table. The
   PDF provenance annotation reads `GcpLastModified`, so reports reflect the new date.
5. **Settings → Data Sources**: shows current table date + path, update service URL,
   device token, "Check for update" / "Install update" buttons.

Local table path resolution at session start (first existing wins):
`AppSettings.GcpDataPath` → `%AppData%\VCCS\gcpPrefixes.xml` → bundled seed beside the
EXE (`data\gcp-prefix-format-list.xml`) → `DefaultOutputDirectory\data\...`.

## Publisher workflow (each GS1 release)

1. Download the new XML from GS1.
2. `gcpEncrypt input.xml gcpPrefixFormatList.xml.enc --key gcpKey.bin`
   → writes the `.enc` and `gcpMeta.json` (generates the key on first ever run).
3. Upload both files to the `gcp-prefix-table` container.
4. Done — clients self-notify on next launch.

## Relationship to the earlier interop endpoint

The previous direct-download path (`/tools/gcp/interop/current.xml` on the My2Dir
resolver with `X-GCP-Interop-Key`) is superseded by this Azure Function service,
which adds the audit trail and encrypted-at-rest distribution. The old
`GCP_INTEROP_KEY` env-var path has been removed from `GcpUpdateService`.
