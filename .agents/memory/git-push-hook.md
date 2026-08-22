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

## Remote-tracking refs
The automatic post-commit push can succeed while the local `origin/main` and
`github/main` tracking refs still point to the previous commit until a fetch.

**Why:** A clean push does not automatically refresh every local remote-tracking
ref, so `git status` may temporarily report the branch as ahead even though the
remote server already has the current commit.

**How to apply:** After the final automatic push, fetch the tracked `main` branch
from each configured GitHub remote before treating divergence output as final.

## Remote hygiene
Task-agent sessions leave temporary `subrepl-*` SSH remotes and matching local
branches. Do not run `git fetch --all`: it contacts those SSH remotes and can
block on a Replit SSH password prompt.

**Why:** A workspace accumulated 60 task-agent remotes; `fetch --all` stalled
after the primary GitHub remotes had refreshed. The task-agent remotes are not
needed after their work has merged.

**How to apply:** When no task merge is active, remove only the `subrepl-*`
remote definitions, preserving local branches plus `origin`, `github`, and
`gitsafe-backup`. Fetch only `origin main` and `github main` to verify a push.

## GitHub connector fallback

When a shell push is rejected because the configured credential is unavailable,
use the already-connected GitHub integration to update the required tracked files
through the Git database API.

**Why:** The workspace may not have a usable shell credential even though the
GitHub integration remains authorized. The connector can create blobs, a tree,
and one non-forced commit after confirming the remote branch has not moved.

**How to apply:** Fetch only `origin main`, compare the exact branch tip, publish
only the approved tracked paths, update the branch without force, then fetch
`origin main` again and verify the remote commit plus the app/report versions.
Never include local SDK binaries, scan-report folders, or Visual Studio solution
state in that update.
