# GCP Update Service — Azure Function

Distributes the encrypted GS1 GCP Prefix Format List to VTCCP workstations with a
device-token gate and an audit trail. Full architecture:
`vtccp/architecture/gcp-update-service.md`.

## One-time Azure setup

1. Storage account → private blob container `gcp-prefix-table`.
2. Function App (dotnet-isolated, .NET 8). Deploy this project:
   `func azure functionapp publish <app-name>` (or CI).
3. App settings (see `local.settings.json.example`):
   - `GcpStorageConnection` — storage connection string
   - `GcpBlobContainer` — default `gcp-prefix-table`
   - `GcpAuditTable` — default `GcpAudit`
   - `DeviceTokens` — semicolon-separated pre-shared tokens, one per workstation/site

## Publishing a new table

1. Download the new XML from GS1.
2. `gcpEncrypt input.xml gcpPrefixFormatList.xml.enc --key gcpKey.bin`
   (tool in `vtccp/tools/gcp-encrypt/`; also writes `gcpMeta.json`).
3. Upload both `gcpPrefixFormatList.xml.enc` and `gcpMeta.json` to the container.
4. Client apps self-notify on next launch.

## Endpoints

| Method | Route | Purpose |
|---|---|---|
| GET | `/api/gcpMeta` | Returns `gcpMeta.json` — `{ "date": "...", "sha256": "..." }` |
| GET | `/api/gcpTable` | Streams the encrypted table blob |

Both require header `X-Device-Token: <token>`; requests are logged to the
`GcpAudit` table (PartitionKey = UTC date, columns: DeviceToken, Endpoint, Ip,
Version, TimestampUtc).

## Local run

Copy `local.settings.json.example` → `local.settings.json`, fill values, then
`func start` (Azure Functions Core Tools v4).
