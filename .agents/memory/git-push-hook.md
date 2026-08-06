---
name: vtccp git push hook
description: How auto-push to GitHub works for the vtccp repo, what broke it, and how it was fixed.
---

## Rule
Edit `.githooks/post-commit` at the **workspace root** — NOT `vtccp/.githooks/post-commit`.

**Why:** The workspace root (`/home/runner/workspace`) IS the git repo root.
`vtccp/` is a subdirectory tracked within it, not a separate git repo.
Any file under `vtccp/.githooks/` is ignored by git's hook mechanism entirely.

**How to apply:** When modifying the post-commit hook, always edit `.githooks/post-commit`
(workspace root). Verify with `git rev-parse --show-toplevel` from any subdirectory —
it always returns `/home/runner/workspace`.

## Two fixes applied 2026-08-06
1. **Token** — `GITHUB_TOKEN` is expired/revoked. `GITHUB_PAT` is the working secret.
   Hook now tries: `GITHUB_PAT2` → `GITHUB_TOKEN2` → `GITHUB_PAT` → `GITHUB_TOKEN`.
2. **GIT_ASKPASS** — Replit sets `GIT_ASKPASS=replit-git-askpass` in the environment.
   This intercepts every `git push` and overrides URL-embedded credentials.
   Fix: `GIT_ASKPASS="" git -c credential.helper="" push --force "$GITHUB_URL" HEAD:main`

## Secret status (as of 2026-08-06)
- `GITHUB_PAT` — confirmed working
- `GITHUB_TOKEN` — present in env but expired/revoked; do not rely on it
- `GITHUB_PAT2` / `GITHUB_TOKEN2` — user created one of these (name not confirmed);
  not yet injected into environment; hook will pick it up automatically on restart
