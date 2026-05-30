#!/bin/bash
# Emergency sync: push Replit's current branch to GitHub even if GitHub is ahead.
#
# This is a FORCE push — it makes GitHub match Replit exactly.
# Only use this when you are certain Replit's code is the authoritative version.
#
# Usage: bash scripts/sync-github.sh

set -euo pipefail

if [ -z "${GITHUB_TOKEN:-}" ]; then
    echo "ERROR: GITHUB_TOKEN is not set. Configure it as a Replit secret."
    exit 1
fi

GITHUB_URL="https://${GITHUB_TOKEN}@github.com/gw4-pips/vtccp.git"

echo "==> Replit HEAD: $(git --no-optional-locks rev-parse --short HEAD)"
echo "==> Force-pushing to GitHub..."

git push --force-with-lease "$GITHUB_URL" HEAD:main 2>&1 \
    | sed "s|https://[^@]*@|https://***@|g"

EXIT=${PIPESTATUS[0]}
if [ $EXIT -eq 0 ]; then
    echo "==> GitHub is now in sync with Replit."
else
    echo "==> Push failed (exit $EXIT)."
    echo "    If --force-with-lease was rejected, someone else pushed to GitHub"
    echo "    after you started. Re-run to retry, or investigate first."
    exit $EXIT
fi
