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
REDACTED_URL="https://***@github.com/gw4-pips/vtccp.git"

LOCAL_SHA=$(git --no-optional-locks rev-parse --short HEAD)
echo "==> Replit HEAD: $LOCAL_SHA"

# Fetch the current remote ref so that --force-with-lease has an up-to-date
# tracking ref to compare against.  In Replit the local repo may never have
# fetched from origin, which causes --force-with-lease to report "stale info"
# and reject the push even though Replit is the authoritative branch.
echo "==> Fetching remote ref to update tracking info..."
if git fetch "$GITHUB_URL" "main:refs/remotes/origin/main" 2>&1 \
        | sed "s|https://[^@]*@|https://***@|g"; then
    REMOTE_SHA=$(git --no-optional-locks rev-parse --short refs/remotes/origin/main 2>/dev/null || echo "unknown")
    echo "==> GitHub HEAD before push: $REMOTE_SHA"
else
    # Fetch failed (e.g. repo is empty or unreachable). Fall through and let
    # the push surface the real error.
    echo "    (Could not fetch remote ref — continuing anyway)"
fi

echo "==> Force-pushing to GitHub ($REDACTED_URL)..."
PUSH_OUTPUT=$(git push --force-with-lease "$GITHUB_URL" HEAD:main 2>&1 \
    | sed "s|https://[^@]*@|https://***@|g") || PUSH_EXIT=$?

echo "$PUSH_OUTPUT"

if [ "${PUSH_EXIT:-0}" -eq 0 ]; then
    echo "==> GitHub is now in sync with Replit."
    exit 0
fi

# --force-with-lease can still be rejected if someone pushed between our fetch
# and our push.  Detect that case and give a clear explanation.
if echo "$PUSH_OUTPUT" | grep -q "stale info\|rejected\|\[rejected\]"; then
    echo ""
    echo "ERROR: Push rejected. GitHub was updated between the fetch and the push."
    echo "       This means someone else pushed to GitHub after this script started."
    echo "       Re-run the script to retry (it will fetch again and try once more)."
    echo "       If Replit is definitely the authoritative source and you want to"
    echo "       override regardless, run:"
    echo "         git push --force \"$REDACTED_URL\" HEAD:main"
else
    echo "==> Push failed (exit ${PUSH_EXIT:-?})."
fi

exit "${PUSH_EXIT:-1}"
