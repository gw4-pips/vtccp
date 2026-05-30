#!/bin/bash
# Emergency sync: push Replit's current branch to GitHub even if GitHub is ahead.
#
# This is a FORCE push — it makes GitHub match Replit exactly.
# Only use this when you are certain Replit's code is the authoritative version.
#
# Usage: bash scripts/sync-github.sh [--dry-run] [--yes]
#   --dry-run  Fetch and show what would be overwritten, then exit without pushing.
#              Safe to use as a status check. --yes is ignored in dry-run mode.
#   --yes      Skip the confirmation prompt when GitHub has commits Replit doesn't.

set -euo pipefail

YES=0
DRY_RUN=0
for arg in "$@"; do
    case "$arg" in
        --yes|-y) YES=1 ;;
        --dry-run) DRY_RUN=1 ;;
        *) echo "Unknown argument: $arg"; exit 1 ;;
    esac
done

if [ "$DRY_RUN" -eq 1 ] && [ "$YES" -eq 1 ]; then
    echo "(note: --yes is ignored in dry-run mode)"
fi

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

    # Check for commits GitHub has that Replit doesn't (would be overwritten).
    AHEAD_COMMITS=$(git --no-optional-locks log HEAD..refs/remotes/origin/main --oneline 2>/dev/null || true)
    if [ -n "$AHEAD_COMMITS" ]; then
        COMMIT_COUNT=$(echo "$AHEAD_COMMITS" | wc -l | tr -d ' ')
        echo ""
        echo "WARNING: GitHub has $COMMIT_COUNT commit(s) that Replit does not:"
        echo "$AHEAD_COMMITS" | sed 's/^/    /'
        echo ""
        echo "A force-push will permanently overwrite these commits on GitHub."
        if [ "$DRY_RUN" -eq 1 ]; then
            echo "==> Dry run complete. No changes were pushed."
            exit 0
        fi
        if [ "$YES" -eq 0 ]; then
            read -r -p "==> Continue with force-push? [y/N] " REPLY
            case "$REPLY" in
                [yY][eE][sS]|[yY]) ;;
                *)
                    echo "Aborted. No changes were pushed to GitHub."
                    exit 1
                    ;;
            esac
        else
            echo "==> --yes flag set, skipping confirmation."
        fi
        echo ""
    else
        if [ "$DRY_RUN" -eq 1 ]; then
            echo "==> GitHub is in sync. No commits would be overwritten."
            echo "==> Dry run complete. No changes were pushed."
            exit 0
        fi
    fi
else
    # Fetch failed (e.g. repo is empty or unreachable).
    if [ "$DRY_RUN" -eq 1 ]; then
        echo "    (Could not fetch remote ref — dry run cannot show diff.)"
        echo "==> Dry run complete. No changes were pushed."
        exit 1
    fi
    # Not dry-run: fall through and let the push surface the real error.
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
