#!/usr/bin/env bash
# rollback.sh — Roll back GAS deployment to a previous Git commit's dist/
# Usage: ./scripts/rollback.sh <git-sha>
set -euo pipefail

SKIP_TESTS=false
AUTO_YES=false
for arg in "$@"; do
  case "$arg" in
    --skip-tests) SKIP_TESTS=true ;;
    --yes|-y) AUTO_YES=true ;;
  esac
done

# First positional arg is the SHA (skip flags)
SHA=""
for arg in "$@"; do
  case "$arg" in
    --*) ;; # skip flags
    *) SHA="$arg"; break ;;
  esac
done

if [ -z "$SHA" ]; then
  echo "Usage: rollback.sh <git-sha> [--skip-tests] [--yes|-y]"
  exit 1
fi

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

if [ "$AUTO_YES" = true ]; then
  # Skip confirmation (--yes / -y flag)
  true
else
  read -rp "Continue with rollback? [y/N] " confirm
  if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
    echo "Rollback cancelled."
    exit 0
  fi
fi

# Print undo command before performing the rollback
echo "To undo this rollback: ./scripts/rollback.sh $(git rev-parse HEAD)"
echo ""

# Checkout dist/ from the target commit
git checkout "$SHA" -- dist/
echo "Restored dist/ from $SHA"

# Commit the rollback so it's recorded in history
git add dist/
git commit --no-verify -m "Rollback dist/ to $SHA"
echo "Committed rollback to git history."

# Run deploy guards unless --skip-tests
if [ "$SKIP_TESTS" = false ]; then
  echo "Running deploy guards..."
  npx jest test/deploy-guards.test.js --no-coverage --bail
  if [ $? -ne 0 ]; then
    echo "Deploy guards failed on rolled-back code. Use --skip-tests to force."
    exit 1
  fi
fi

# Push to GAS
echo "Pushing to Google Apps Script..."
npx clasp push --force

echo ""
echo "Rollback complete. Deployed dist/ from commit $SHA."
echo "To undo this rollback, run: git revert HEAD"
