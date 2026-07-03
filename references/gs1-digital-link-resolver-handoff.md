# GS1 Digital Link Resolver — Handoff Notes

Rev 1.0 — 2026-07-02

## Purpose of this document

This is a handoff summary from the VTCCP/Command Pilot project's Replit agent to whoever
(human or agent) is standing up the new, separate GS1 Digital Link resolver project on
Azure. It captures architectural decisions and context from planning discussions that
happened in the VTCCP workspace before the resolver had its own project. It is **not**
a design spec for the resolver itself — that lives in the design plan/scaffolding already
produced with Claude. This document exists purely to carry over context about the shared
GS1 parsing engine so the resolver doesn't end up as an independent reimplementation of
the same spec.

## What the resolver is

A GS1 Digital Link resolver service: parses incoming requests in GS1 Digital Link URI
form (e.g. `/01/<GTIN>/21/<serial>`), looks up a mapping, and redirects the consumer to
the appropriate destination page. Public-facing web service, unrelated in product,
branding, and purpose to VTCCP/Command Pilot (which is an on-prem/desktop barcode
verification platform for Cognex DataMan and Webscan TruCheck hardware). The two
projects are intentionally kept separate; the only thing they should share is the
underlying GS1 parsing engine, not code, infrastructure, or deployment.

## Why this matters: the shared engine decision

VTCCP already vendors GS1's own official reference implementation for AI element-string
and GS1 Digital Link URI parsing/validation — the **GS1 Barcode Syntax Engine**
(`gs1.github.io/gs1-syntax-engine`, GitHub `gs1/gs1-syntax-engine`), not a custom or
third-party reimplementation. Key facts:

- **License**: Apache 2.0. The project's own README explicitly permits vendoring the
  source into an application codebase (open source or proprietary) or redistributing
  the pre-built shared library — using it in two separate applications is a sanctioned,
  intended use case. No licensing obstacle to reuse.
- **What VTCCP vendors**: the native C core (`gs1encoders.dll`) plus the official C#
  .NET P/Invoke wrapper (`GS1Encoder.cs`). Located at `vtccp/lib/gs1-syntax-engine/` in
  the VTCCP repo, under `LICENSE` (Apache 2.0) and `README.md` for reference.
- **Capabilities already confirmed working in VTCCP**: parses/validates bracketed AI
  element strings (e.g. `(01)00312345678905(21)ABC123`), parses GS1 Digital Link URIs
  (auto-detects `http://`/`https://` prefix via the `DataStr` setter), and **generates**
  DL URIs from AI data via `GetDLuri(stem)`. Full bidirectional AI ↔ Digital Link
  conversion, which is exactly what a resolver needs on the decode side.
- **Important distinction**: this is a different project from GS1's separately
  maintained `GS1DigitalLinkToolkit.js` (an independently-authored JS tool). Do not
  conflate the two — reusing `GS1DigitalLinkToolkit.js` alongside VTCCP's vendored
  engine would reintroduce the exact "two independent implementations of the same spec"
  drift risk this handoff is meant to avoid.

## Recommended integration approach for the resolver

**Do not stand up a shared network service between VTCCP and the resolver.** The two
projects have very different deployment profiles (VTCCP: on-prem/desktop; resolver:
public-facing web service on Azure), and a shared running service would create an
unwanted uptime/network coupling between otherwise-independent products.

**Instead: vendor the matching official binding of the same upstream release into the
resolver project directly.**

- The upstream `gs1-syntax-engine` project ships **official bindings from the same C
  source** for multiple languages/runtimes, including a **JavaScript + WebAssembly**
  binding (browser and Node.js), a **C# .NET** wrapper (what VTCCP uses), Java, and
  Swift. The JS/Wasm binding is compiled from the identical C core — not a separate
  reimplementation — so vendoring it into the resolver gives you the same authoritative
  parser without any network dependency on VTCCP.
- If the resolver ends up being .NET, it's even simpler: reference the same C# wrapper
  project VTCCP already uses, no Wasm layer needed.
- Whichever binding the resolver uses, **pin it to the same upstream release version**
  as VTCCP's vendored copy, and track version upgrades in both projects together. This
  is the only real discipline required to keep "one authoritative engine" true over
  time — the risk isn't licensing, it's the two vendored copies silently drifting apart
  on separate upgrade schedules.
- Upstream repo also ships a documented Node.js HTTP web-service example
  (`src/js-wasm`, "Node.js HTTP web service with example client") worth reviewing as a
  reference pattern for how GS1 itself expects the JS/Wasm binding to be used in a
  server context — relevant since the resolver is exactly that kind of service.

## Domain / hosting notes (tangential, for completeness)

Discussed separately from the engine question: Replit supports both purchasing a new
domain directly (with auto DNS configuration and WHOIS privacy) and connecting an
already-owned domain to a deployment. Since the resolver is being hosted on Azure
instead, this is not directly applicable, but is noted here in case domain
management ownership (registrar vs. host) comes up again — no technical dependency
either way, purely a matter of where billing/renewal control should sit.

## What's *not* covered here

- The resolver's actual mapping/lookup data model, routing logic, or Azure service
  selection — that's in the design plan/scaffolding already produced with Claude.
- Any VTCCP-specific code, architecture, or roadmap — the resolver should not depend on
  or reference VTCCP's codebase beyond the shared engine version alignment described
  above.
