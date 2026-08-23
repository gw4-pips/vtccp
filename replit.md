# Workspace

## Overview

pnpm workspace monorepo using TypeScript. Each package manages its own dependencies.

## Canonical VTCCP delivery

For VTCCP changes, “done” means **Build → Push → Confirm**, in that order:

```bash
bash scripts/build-push-confirm.sh
```

The command builds the Windows-targeted application, refuses to deliver an
uncommitted tree, pushes the exact local `HEAD`, fetches `origin/main`, and
fails unless GitHub resolves to the same commit. It also prints the application
version confirmed on GitHub. Do not report a VTCCP build as delivered until this
command completes successfully.

## Stack

- **Monorepo tool**: pnpm workspaces
- **Node.js version**: 24
- **Package manager**: pnpm
- **TypeScript version**: 5.9
- **API framework**: Express 5
- **Database**: PostgreSQL + Drizzle ORM
- **Validation**: Zod (`zod/v4`), `drizzle-zod`
- **API codegen**: Orval (from OpenAPI spec)
- **Build**: esbuild (CJS bundle)

## Structure

```text
artifacts-monorepo/
├── artifacts/              # Deployable applications
│   └── api-server/         # Express API server
├── lib/                    # Shared libraries
│   ├── api-spec/           # OpenAPI spec + Orval codegen config
│   ├── api-client-react/   # Generated React Query hooks
│   ├── api-zod/            # Generated Zod schemas from OpenAPI
│   └── db/                 # Drizzle ORM schema + DB connection
├── scripts/                # Utility scripts (single workspace package)
│   └── src/                # Individual .ts scripts, run via `pnpm --filter @workspace/scripts run <script>`
├── pnpm-workspace.yaml     # pnpm workspace (artifacts/*, lib/*, lib/integrations/*, scripts)
├── tsconfig.base.json      # Shared TS options (composite, bundler resolution, es2022)
├── tsconfig.json           # Root TS project references
└── package.json            # Root package with hoisted devDeps
```

## TypeScript & Composite Projects

Every package extends `tsconfig.base.json` which sets `composite: true`. The root `tsconfig.json` lists all packages as project references. This means:

- **Always typecheck from the root** — run `pnpm run typecheck` (which runs `tsc --build --emitDeclarationOnly`). This builds the full dependency graph so that cross-package imports resolve correctly. Running `tsc` inside a single package will fail if its dependencies haven't been built yet.
- **`emitDeclarationOnly`** — we only emit `.d.ts` files during typecheck; actual JS bundling is handled by esbuild/tsx/vite...etc, not `tsc`.
- **Project references** — when package A depends on package B, A's `tsconfig.json` must list B in its `references` array. `tsc --build` uses this to determine build order and skip up-to-date packages.

## Root Scripts

- `pnpm run build` — runs `typecheck` first, then recursively runs `build` in all packages that define it
- `pnpm run typecheck` — runs `tsc --build --emitDeclarationOnly` using project references

## Packages

### `artifacts/api-server` (`@workspace/api-server`)

Express 5 API server. Routes live in `src/routes/` and use `@workspace/api-zod` for request and response validation and `@workspace/db` for persistence.

- Entry: `src/index.ts` — reads `PORT`, starts Express
- App setup: `src/app.ts` — mounts CORS, JSON/urlencoded parsing, routes at `/api`
- Routes: `src/routes/index.ts` mounts sub-routers; `src/routes/health.ts` exposes `GET /health` (full path: `/api/health`)
- Depends on: `@workspace/db`, `@workspace/api-zod`
- `pnpm --filter @workspace/api-server run dev` — run the dev server
- `pnpm --filter @workspace/api-server run build` — production esbuild bundle (`dist/index.cjs`)
- Build bundles an allowlist of deps (express, cors, pg, drizzle-orm, zod, etc.) and externalizes the rest

### `lib/db` (`@workspace/db`)

Database layer using Drizzle ORM with PostgreSQL. Exports a Drizzle client instance and schema models.

- `src/index.ts` — creates a `Pool` + Drizzle instance, exports schema
- `src/schema/index.ts` — barrel re-export of all models
- `src/schema/<modelname>.ts` — table definitions with `drizzle-zod` insert schemas (no models definitions exist right now)
- `drizzle.config.ts` — Drizzle Kit config (requires `DATABASE_URL`, automatically provided by Replit)
- Exports: `.` (pool, db, schema), `./schema` (schema only)

Production migrations are handled by Replit when publishing. In development, we just use `pnpm --filter @workspace/db run push`, and we fallback to `pnpm --filter @workspace/db run push-force`.

### `lib/api-spec` (`@workspace/api-spec`)

Owns the OpenAPI 3.1 spec (`openapi.yaml`) and the Orval config (`orval.config.ts`). Running codegen produces output into two sibling packages:

1. `lib/api-client-react/src/generated/` — React Query hooks + fetch client
2. `lib/api-zod/src/generated/` — Zod schemas

Run codegen: `pnpm --filter @workspace/api-spec run codegen`

### `lib/api-zod` (`@workspace/api-zod`)

Generated Zod schemas from the OpenAPI spec (e.g. `HealthCheckResponse`). Used by `api-server` for response validation.

### `lib/api-client-react` (`@workspace/api-client-react`)

Generated React Query hooks and fetch client from the OpenAPI spec (e.g. `useHealthCheck`, `healthCheck`).

### `scripts` (`@workspace/scripts`)

Utility scripts package. Each script is a `.ts` file in `src/` with a corresponding npm script in `package.json`. Run scripts via `pnpm --filter @workspace/scripts run <script>`. Scripts can import any workspace package (e.g., `@workspace/db`) by adding it as a dependency in `scripts/package.json`.

---

## VTCCP — VCCS DMV TruCheck Command Pilot

Native Windows desktop utility (WPF + C#/.NET 8) located in `vtccp/`. Separate from the pnpm monorepo — built and run directly with the `dotnet` CLI.

### Solution: `vtccp/VTCCP.sln`

| Project | Phase | Status |
|---|---|---|
| `ExcelEngine` | 1 | Complete — 167-column schema, XLS/XLSX adapters, session manager, batch extraction |
| `DeviceInterface` | 2 | Complete — DMCC TCP client, DMST XML parser, DeviceSession, MockDmccServer |
| `ConfigEngine` | 3 | Complete — DeviceProfile, JobTemplate, AppSettings, ConfigStore, ConfigRepository |
| `VtccpApp` | 3+4 | Complete — WPF shell (MainWindow, nav sidebar, DevicesView, TemplatesView, SessionView, HistoryView) |
| `TestHarness` | 1+2+3+4 | Complete — Phase 1 Tasks 1–4 + 7 Phase 2 + 6 Phase 3 + 6 Phase 4 sub-checks, all PASS |

### Build & test

```bash
cd vtccp
dotnet build VTCCP.sln -c Release
dotnet run --project TestHarness/TestHarness.csproj -c Release
```

### Key architectural notes

- `DmccClient.ReadUntilIdleAsync` uses synchronous `Socket.Receive` on a thread-pool thread (not `NetworkStream.ReadAsync`) because on Linux/.NET 8, passing a `CancellationToken` to `ReadAsync` disposes the socket when the token fires.
- `GradingResult.FromLetterAndNumeric(letter, decimal, passFail)` — second param is non-nullable `decimal`; `DmstResultParser` defaults to `0m` when the XML attribute is absent.
- `MockDmccServer` binds to `IPAddress.Loopback, 0` (OS-assigned port) exposed via `Port` property.
- Phase 3 is complete: ConfigEngine (JSON persistence) + VtccpApp WPF shell (MainWindow navigation, DevicesView, TemplatesView, SessionView).
- Phase 4 is complete: HistoryView DataGrid (live-updating, grade/pass-fail/symbology filters, copy-TSV); ScanResultRow + HistoryFilter models; HistoryViewModel wired into SessionViewModel on every trigger.

## Product terminology

- **DataMan TruCheck (DM TC)** — Cognex's built-in verification interface running on the DM475V device. HTML reports written to `Documents\{DeviceName}\CodeQuality\` are produced by DM TC. This is the current test environment.
- **Webscan TruCheck** — a separate, third-party PC verification application (Webscan Inc.). It has its own HTML report format and UI. The project will integrate with Webscan TruCheck next, after DM TC testing is complete.
- **Do not conflate these two products.** They share the "TruCheck" name but are entirely different software from different vendors. Any mention of "TruCheck" must specify which one.

## User preferences

- **Append to transcript after every response**: at the end of every response, append the user's message and my reply (text only, no tool detail) to `transcript/chat-transcript.md`. Format: `**User:** …` then `**Assistant:** …` separated by a blank line, under the current date heading. This is a standing rule — never skip it.

- **Always close VTCCP before Git Pull**: remind the user to close the running VTCCP app before pulling in Visual Studio, to avoid file-lock conflicts on the binaries.

- **Always bump the app version on every code-change commit**: increment `VtccpApp.csproj <Version>` (patch digit) on every commit that changes any C# code. Never commit a code change without a version bump. This is non-negotiable.

- **Commit, Push, Confirm protocol**: after completing any requested code change, commit it, push the commit to the primary GitHub `origin/main`, then verify the remote commit SHA and explicitly confirm the result to the user. Do not wait for a reminder.

- **Always bump the report version whenever the report format changes**: the canonical VCCS RFID Validation Report is the HTML-based v23 format (`dist/vccs-pdf-preview-v23.html`). Its footer version string (e.g. "v1.4.11") must be incremented any time the report layout, content, or logic changes. The report version and the app version are separate numbers — both must be maintained.

- **The canonical VCCS RFID report is v23 HTML rendered silently to PDF**: the generator produces v23 HTML (`dist/vccs-pdf-preview-v23.html` is the design reference), then converts it to a PDF automatically using WebView2 (Edge, primary) with a seamless silent fallback to bundled `wkhtmltopdf.exe`. The user never sees the HTML and takes no manual step — a `.pdf` file is the only output. QuestPDF (`PdfReportGenerator.cs`) has been archived and must not be extended or resurrected. Do not ask the user to open HTML and print.

- **DO NOT BUILD WITHOUT ASKING first**: before writing or modifying any C# code, confirm the exact approach with the user. Planning and analysis are the default. This is the highest-priority rule for the VTCCP project.

- **Push-script viewer — ALWAYS update on every new version**: every time a new push script version is written (any vX.YY), immediately and without being asked: (1) copy the new script to `artifacts/script-viewer/src/vXYY.txt`, (2) update `App.tsx` to import `vXYY.txt?raw` and change the header label, download filename, and install-confirm `<PushScriptDiag>` string to the new version, (3) restart the `artifacts/script-viewer: web` workflow. This is not optional and must not require a reminder.
- **Always rev synthesis/architecture documents when updating them**: any time a `.md` or `.html` doc in `references/` (or any other synthesis/design document) is modified, increment its version in the header — e.g. add or bump a `v1.0 → v1.1` tag and revision date in the top section. Never deliver an updated doc without a version bump. If the document has no version header yet, add one on first edit.

- **Language**: say "waiting for you" not "waiting on you."

- **When told to use a file as a template, use it literally**: load and modify that exact file (strip junk, add token markers) — do not use it as a reference to reimplement the same content in code. "Use X as the template" means X IS the template, not "make something that looks like X."

- **Verifier data is canonical — never re-derive it**: any value the verifier (DM TC or Webscan TruCheck) reports must be taken verbatim from the verifier's own output (DM TC HTML or push XML). VTCCP must NEVER recalculate, reconstruct, or substitute a locally-computed value (grades, check digits, pass/fail outcomes, AI extraction, codeword counts, symbology names, DFC rows, etc.) in place of what the verifier reported. If the verifier data is absent, the field must be left absent or blank — not filled in algorithmically. Any exception to this rule requires an explicit design discussion before any code is written.

- **Ask more, assume less**: when implementation details are ambiguous or multiple valid approaches exist, ask a clarifying question rather than picking one and discovering the wrong choice after the fact. The cost of a question is far lower than the cost of a rework.

- **Review recent past work in detail before starting any session**: because this development is sporadic and sessions can be days apart, do not assume continuity. At the start of every session — and before writing any code — read `transcript/chat-transcript.md` (recent entries), the active memory index (`.agents/memory/MEMORY.md`), and any topic files relevant to the task. If something is not definitively in active memory or confirmed by direct reading, research it. Never proceed on an assumption; verify first.

- **Always provide complete URLs when referencing mockup previews or any hosted artifact**: never give a bare path fragment like `/preview/grading/GradingStandards`. Always give the full URL the user can open directly, e.g. `https://3e1c7688-a8f7-43a4-a93e-fbbb755e6a82-00-2uu0hix24eyfn.worf.replit.dev/__mockup/preview/grading/GradingStandards`. Mockup sandbox base: `https://3e1c7688-a8f7-43a4-a93e-fbbb755e6a82-00-2uu0hix24eyfn.worf.replit.dev/__mockup/preview/<subfolder>/<ComponentName>` where subfolder/ComponentName mirrors the file path under `artifacts/mockup-sandbox/src/components/mockups/`.
