#!/bin/bash
set -e
pnpm install --frozen-lockfile
pnpm --filter db push
bash scripts/setup-git-hooks.sh

# Replit task merges update the branch outside of Git's normal `commit` path,
# so `.githooks/post-commit` is not invoked automatically. Run the same
# authenticated, status-logging sync after every completed task merge.
bash .githooks/post-commit
