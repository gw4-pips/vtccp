#!/bin/bash
# Run this ONCE in the Replit shell to configure auto-push to GitHub.
# After this, every commit/checkpoint will automatically push to GitHub.
#
# Usage: bash scripts/setup-git-hooks.sh

set -euo pipefail

echo "==> Configuring git to use .githooks/ directory..."
git config core.hooksPath .githooks
echo "    core.hooksPath = .githooks"

echo "==> Making hooks executable..."
chmod +x .githooks/post-commit
echo "    .githooks/post-commit — OK"

echo ""
echo "==> Done. Every future Replit commit will now auto-push to GitHub."
echo "    You can verify the setting with: git config --get core.hooksPath"
