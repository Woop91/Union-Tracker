#!/usr/bin/env bash
# rollback.sh — Roll back GAS deployment to a previous Git commit's dist/
# Usage: ./scripts/rollback.sh <git-sha>
set -euo pipefail

SHA="${1:?Usage: rollback.sh <git-sha>}"
REPO_ROOT="$(git rev-parse --show-toplevel)"

echo "=== DDS Dashboard Rollback ==="
echo "Target commit: $SHA"

# Verify the SHA exists
if ! git cat-file -t "$SHA" >/dev/null 2>&1; then
  echo "ERROR: Commit $SHA not found in this repository."
  exit 1
fi

# Verify dist/ exists in that commit
if ! git ls-tree -d "$SHA" -- dist/ >/dev/null 2>&1; then
  echo "ERROR: No dist/ directory found in commit $SHA."
  exit 1
fi

# Safety: ensure working tree is clean
if [[ -n "$(git status --porcelain)" ]]; then
  echo "ERROR: Working tree is dirty. Commit or stash changes first."
  exit 1
fi

# Show what we're rolling back to
echo ""
echo "Rolling back to:"
git log --oneline -1 "$SHA"
echo ""

read -rp "Continue with rollback? [y/N] " confirm
if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
  echo "Rollback cancelled."
  exit 0
fi

# Checkout dist/ from the target commit
git checkout "$SHA" -- dist/
echo "Restored dist/ from $SHA"

# Commit the rollback so it's recorded in history
git add dist/
git commit -m "Rollback dist/ to $SHA"
echo "Committed rollback to git history."

# Push to GAS
echo "Pushing to Google Apps Script..."
npx clasp push --force

echo ""
echo "Rollback complete. Deployed dist/ from commit $SHA."
echo "To undo this rollback, run: git revert HEAD"
