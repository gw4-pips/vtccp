#!/bin/bash
# Push current Replit main branch to GitHub via the GitHub Contents API.
# Run this from the Replit shell after any session where code was changed.
#
# Usage: bash scripts/push-github.sh
#
# Requires: GITHUB_TOKEN Replit secret (already configured).

set -euo pipefail

OWNER="gw4-pips"
REPO="vtccp"

if [ -z "${GITHUB_TOKEN:-}" ]; then
    echo "ERROR: GITHUB_TOKEN is not set. Configure it as a Replit secret."
    exit 1
fi

echo "==> Collecting changed files vs GitHub main..."

# Get the list of files changed in commits GitHub doesn't have yet.
# We compare against the GitHub remote using the API to find the common ancestor.
GITHUB_SHA=$(curl -s \
    -H "Authorization: Bearer $GITHUB_TOKEN" \
    -H "Accept: application/vnd.github.v3+json" \
    "https://api.github.com/repos/$OWNER/$REPO/git/refs/heads/main" \
    | python3 -c "import sys,json; print(json.load(sys.stdin)['object']['sha'])")

echo "    GitHub HEAD: $GITHUB_SHA"
echo "    Replit HEAD: $(git --no-optional-locks rev-parse HEAD)"

# Get all files that differ between GitHub HEAD and Replit HEAD.
# Primary: git diff between the two SHAs (requires GitHub SHA to be in local history).
# Fallback: if GitHub SHA is not in local history (Replit checkpoint divergence), collect
#           files changed in the last 30 commits on the Replit side — safe over-approximation.
CHANGED=""
if git --no-optional-locks cat-file -e "$GITHUB_SHA^{commit}" 2>/dev/null; then
    CHANGED=$(git --no-optional-locks diff --name-only "$GITHUB_SHA" HEAD 2>/dev/null)
    echo "    Diff mode: git diff $GITHUB_SHA..HEAD"
else
    echo "    WARNING: GitHub SHA not in local history (Replit checkpoint divergence)."
    echo "    Fallback: collecting files changed in last 30 Replit commits."
    CHANGED=$(git --no-optional-locks log --name-only --pretty=format: -n 30 HEAD 2>/dev/null \
              | grep -v '^$' | sort -u)
fi

if [ -z "$CHANGED" ]; then
    echo "==> Nothing to push — GitHub is already up to date."
    exit 0
fi

echo "==> Files to update:"
echo "$CHANGED" | sed 's/^/    /'

PUSHED=0
FAILED=0

while IFS= read -r FILE; do
    [ -z "$FILE" ] && continue

    if [ ! -f "$FILE" ]; then
        echo "  SKIP (deleted or not a file): $FILE"
        continue
    fi

    # Get current SHA of this file on GitHub (needed for update API).
    META=$(curl -s \
        -H "Authorization: Bearer $GITHUB_TOKEN" \
        -H "Accept: application/vnd.github.v3+json" \
        "https://api.github.com/repos/$OWNER/$REPO/contents/$(python3 -c "import urllib.parse,sys; print(urllib.parse.quote('$FILE', safe='/'))")")

    FILE_SHA=$(echo "$META" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('sha',''))" 2>/dev/null || true)

    COMMIT_MSG="Sync: $FILE"

    # Write payload to a temp file to avoid "Argument list too long" on large binaries.
    PAYLOAD_FILE=$(mktemp /tmp/gh_payload_XXXXXX.json)
    FPATH="$FILE"
    if [ -n "$FILE_SHA" ]; then
        python3 -c "
import json, base64
content = open('$FPATH', 'rb').read()
print(json.dumps({'message': '$COMMIT_MSG', 'content': base64.b64encode(content).decode(), 'sha': '$FILE_SHA'}))
" > "$PAYLOAD_FILE"
    else
        python3 -c "
import json, base64
content = open('$FPATH', 'rb').read()
print(json.dumps({'message': '$COMMIT_MSG', 'content': base64.b64encode(content).decode()}))
" > "$PAYLOAD_FILE"
    fi

    STATUS=$(curl -s -o /tmp/gh_push_resp.json -w "%{http_code}" \
        -X PUT \
        -H "Authorization: Bearer $GITHUB_TOKEN" \
        -H "Accept: application/vnd.github.v3+json" \
        -H "Content-Type: application/json" \
        --data-binary "@$PAYLOAD_FILE" \
        "https://api.github.com/repos/$OWNER/$REPO/contents/$(python3 -c "import urllib.parse,sys; print(urllib.parse.quote('$FILE', safe='/'))")")
    rm -f "$PAYLOAD_FILE"

    if [ "$STATUS" = "200" ] || [ "$STATUS" = "201" ]; then
        echo "  OK ($STATUS): $FILE"
        PUSHED=$((PUSHED + 1))
    else
        echo "  FAIL ($STATUS): $FILE"
        python3 -c "import json; d=json.load(open('/tmp/gh_push_resp.json')); print('    ', d.get('message','?'))" 2>/dev/null || true
        FAILED=$((FAILED + 1))
    fi

done <<< "$CHANGED"

echo ""
echo "==> Done. Pushed: $PUSHED  Failed: $FAILED"
[ $FAILED -eq 0 ] && exit 0 || exit 1
