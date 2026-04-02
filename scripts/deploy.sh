#!/bin/bash
# deploy.sh — Safe deployment wrapper for clasp push
# Ensures build:prod + minification + size budget check before pushing to GAS.
#
# Usage:
#   ./scripts/deploy.sh          # deploy to default (production) target
#   ./scripts/deploy.sh --dry    # validate only, don't push

set -euo pipefail

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

DRY_RUN=false
if [[ "${1:-}" == "--dry" ]]; then
  DRY_RUN=true
  echo -e "${YELLOW}DRY RUN — will validate but not push${NC}"
fi

echo ""
echo "═══════════════════════════════════════"
echo "  DDS/UT Safe Deploy Pipeline"
echo "═══════════════════════════════════════"
echo ""

# Step 1: Lint
echo "[1/6] Running lint..."
npm run lint
echo ""

# Step 2: Build prod with minification
echo "[2/6] Building with build:prod --minify..."
npm run build:prod
echo ""

# Step 3: Run deploy guards
echo "[3/6] Running deploy guards..."
npx jest test/deploy-guards.test.js test/spa-integrity.test.js --no-coverage --bail 2>&1 | tail -5
echo ""

# Step 4: Run unit tests
echo "[4/6] Running unit tests..."
npm run test:unit
echo ""

# Step 5: HTML size budget check
echo "[5/6] Checking HTML size budget..."
MAX_KB=820
INDEX_SIZE=$(wc -c < dist/index.html)
STYLES_SIZE=$(wc -c < dist/styles.html)
STEWARD_SIZE=$(wc -c < dist/steward_view.html)
MEMBER_SIZE=$(wc -c < dist/member_view.html)
SHARED_SIZE=$(wc -c < dist/shared_components.html)

STEWARD_TOTAL=$((INDEX_SIZE + STYLES_SIZE + STEWARD_SIZE + SHARED_SIZE))
MEMBER_TOTAL=$((INDEX_SIZE + STYLES_SIZE + MEMBER_SIZE + SHARED_SIZE))
BOTH_TOTAL=$((INDEX_SIZE + STYLES_SIZE + STEWARD_SIZE + MEMBER_SIZE + SHARED_SIZE))
MAX_BYTES=$((MAX_KB * 1024))

echo "  Steward-only: $((STEWARD_TOTAL / 1024)) KB"
echo "  Member-only:  $((MEMBER_TOTAL / 1024)) KB"
echo "  Dual-role:    $((BOTH_TOTAL / 1024)) KB"
echo "  GAS limit:    ${MAX_KB} KB"

if [ "$BOTH_TOTAL" -gt "$MAX_BYTES" ]; then
  echo -e "${RED}FAIL: Dual-role payload ($((BOTH_TOTAL / 1024)) KB) exceeds ${MAX_KB} KB GAS limit!${NC}"
  exit 1
fi
echo -e "${GREEN}  All payloads under limit.${NC}"
echo ""

# Step 6: Push to GAS
if [ "$DRY_RUN" = true ]; then
  echo -e "${YELLOW}[6/6] DRY RUN — skipping clasp push${NC}"
else
  echo "[6/6] Pushing to GAS..."
  npx clasp push --force
fi

echo ""
echo -e "${GREEN}═══════════════════════════════════════${NC}"
echo -e "${GREEN}  Deploy complete!${NC}"
echo -e "${GREEN}═══════════════════════════════════════${NC}"
