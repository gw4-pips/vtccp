#!/usr/bin/env bash
# Canonical VTCCP delivery gate:
#   1. build
#   2. require a committed tree
#   3. push the exact current HEAD
#   4. fetch GitHub and confirm the remote SHA and app version
#
# Usage: bash scripts/build-push-confirm.sh

set -euo pipefail

ROOT="$(git rev-parse --show-toplevel)"
cd "$ROOT"

echo "==> Checking repository state..."
if [[ -n "$(git status --porcelain)" ]]; then
  echo "ERROR: Working tree is dirty. Commit the completed change before delivery."
  git status --short
  exit 1
fi

echo "==> Building VTCCP..."
dotnet build vtccp/VtccpApp/VtccpApp.csproj \
  -p:EnableWindowsTargeting=true \
  --no-restore

echo "==> Pushing exact HEAD to GitHub..."
bash .githooks/post-commit

echo "==> Confirming GitHub..."
git fetch origin main
LOCAL_SHA="$(git rev-parse HEAD)"
REMOTE_SHA="$(git rev-parse origin/main)"
if [[ "$LOCAL_SHA" != "$REMOTE_SHA" ]]; then
  echo "ERROR: GitHub SHA does not match local HEAD."
  echo "       local : $LOCAL_SHA"
  echo "       remote: $REMOTE_SHA"
  exit 1
fi

REMOTE_VERSION="$(git show "origin/main:vtccp/VtccpApp/VtccpApp.csproj" |
  sed -n 's/.*<Version>\([^<]*\)<\/Version>.*/\1/p' | head -n 1)"
echo "==> Confirmed GitHub commit: $REMOTE_SHA"
echo "==> Confirmed remote application version: ${REMOTE_VERSION:-unknown}"