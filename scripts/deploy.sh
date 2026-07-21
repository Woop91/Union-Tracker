#!/bin/bash
# deploy.sh — Safe, exact-SHA production deployment wrapper.
# Builds, pushes, versions, redeploys, and verifies the public health endpoint.
#
# Usage:
#   ./scripts/deploy.sh          # deploy to default (production) target
#   ./scripts/deploy.sh --dry    # validate only, don't push

set -euo pipefail

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

# npm can launch this script through WSL on Windows. WSL resolves the Windows
# runtime as node.exe (not node), so select the executable explicitly.
if command -v node >/dev/null 2>&1; then
  NODE_BIN='node'
elif command -v node.exe >/dev/null 2>&1; then
  NODE_BIN='node.exe'
else
  echo -e "${RED}FAIL: Node.js is not available in this shell.${NC}"
  exit 1
fi

DRY_RUN=false
if [ "${1:-}" == "--dry" ]; then
  DRY_RUN=true
  echo -e "${YELLOW}DRY RUN — will validate but not push${NC}"
elif [ -n "${1:-}" ]; then
  echo "⚠️  Unrecognized argument: $1 (did you mean --dry?)"
  exit 1
fi

SOURCE_SHA="$(git rev-parse HEAD)"
CURRENT_BRANCH="$(git branch --show-current)"
APP_VERSION="$($NODE_BIN -p "require('./package.json').version")"

if [ -n "$(git status --porcelain --untracked-files=no)" ]; then
  echo -e "${RED}FAIL: tracked working tree is dirty. Commit first.${NC}"
  exit 1
fi

if [ "$DRY_RUN" = false ]; then
  if [ "$CURRENT_BRANCH" != "Main" ]; then
    echo -e "${RED}FAIL: production deploy requires checked-out Main; found $CURRENT_BRANCH.${NC}"
    exit 1
  fi
  if [ -z "${SOLIDBASE_DEPLOYMENT_ID:-}" ]; then
    echo -e "${RED}FAIL: SOLIDBASE_DEPLOYMENT_ID is not configured.${NC}"
    exit 1
  fi
fi

# Release builds replace a source placeholder in dist/. Always restore the
# deterministic tracked build so deploy verification cannot leave a dirty tree.
restore_dist() {
  "$NODE_BIN" build.js --prod --minify >/dev/null 2>&1 || true
}
trap restore_dist EXIT

echo ""
echo "═══════════════════════════════════════"
echo "  SolidBase Exact-SHA Deploy Pipeline"
echo "═══════════════════════════════════════"
echo "  Version: $APP_VERSION"
echo "  SHA:     $SOURCE_SHA"
echo ""

# Step 1: Lint
echo "[1/8] Running lint..."
npm run lint
echo ""

# Step 2: Build prod with minification
echo "[2/8] Building release artifact with exact source SHA..."
"$NODE_BIN" build.js --prod --minify --source-sha "$SOURCE_SHA"
if ! grep -q "$SOURCE_SHA" dist/01_Core.gs; then
  echo -e "${RED}FAIL: release artifact does not contain $SOURCE_SHA.${NC}"
  exit 1
fi
echo ""

# Step 3: Run deploy guards
echo "[3/8] Running deploy guards..."
npx jest test/deploy-guards.test.js test/spa-integrity.test.js --no-coverage --bail 2>&1 | tail -5
echo ""

# Step 4: Run unit tests
echo "[4/8] Running unit tests..."
npm run test:unit
echo ""

# Step 5: HTML size budget check
echo "[5/8] Checking HTML size budget..."
MAX_KB=810  # GAS limit ~820KB; 10KB safety margin
INDEX_SIZE=$(wc -c < dist/index.html)
STYLES_SIZE=$(wc -c < dist/styles.html)
STEWARD_SIZE=$(wc -c < dist/steward_view.html)
MEMBER_SIZE=$(wc -c < dist/member_view.html)
SHARED_SIZE=$(wc -c < dist/shared_components.html)

STEWARD_TOTAL=$((INDEX_SIZE + STYLES_SIZE + STEWARD_SIZE + SHARED_SIZE))
MEMBER_TOTAL=$((INDEX_SIZE + STYLES_SIZE + MEMBER_SIZE + SHARED_SIZE))
MAX_BYTES=$((MAX_KB * 1024))
MAX_INITIAL_TOTAL=$STEWARD_TOTAL
if [ "$MEMBER_TOTAL" -gt "$MAX_INITIAL_TOTAL" ]; then
  MAX_INITIAL_TOTAL=$MEMBER_TOTAL
fi

echo "  Steward-only: $((STEWARD_TOTAL / 1024)) KB"
echo "  Member-only:  $((MEMBER_TOTAL / 1024)) KB"
echo "  Dual-role:    $((STEWARD_TOTAL / 1024)) KB initial (member view lazy-loaded)"
echo "  GAS limit:    ${MAX_KB} KB"

if [ "$MAX_INITIAL_TOTAL" -gt "$MAX_BYTES" ]; then
  echo -e "${RED}FAIL: Initial payload ($((MAX_INITIAL_TOTAL / 1024)) KB) exceeds ${MAX_KB} KB GAS limit!${NC}"
  exit 1
fi
echo -e "${GREEN}  All payloads under limit.${NC}"
echo ""

# Step 6: Push source to GAS
if [ "$DRY_RUN" = true ]; then
  echo -e "${YELLOW}[6/8] DRY RUN — skipping clasp push${NC}"
  echo -e "${YELLOW}[7/8] DRY RUN — skipping immutable version + redeploy${NC}"
  echo -e "${YELLOW}[8/8] DRY RUN — skipping public health verification${NC}"
else
  echo "[6/8] Pushing to GAS..."
  npx clasp push --force

  echo "[7/8] Creating immutable version and updating deployment..."
  VERSION_OUTPUT="$(npx clasp version "v${APP_VERSION} git:${SOURCE_SHA}")"
  VERSION_NUMBER="$(printf '%s\n' "$VERSION_OUTPUT" | grep -Eo '[0-9]+' | tail -1)"
  if ! [[ "$VERSION_NUMBER" =~ ^[0-9]+$ ]]; then
    echo -e "${RED}FAIL: could not parse clasp version number.${NC}"
    exit 1
  fi
  npx clasp deploy --deploymentId "$SOLIDBASE_DEPLOYMENT_ID" \
    --versionNumber "$VERSION_NUMBER" \
    --description "v${APP_VERSION} git:${SOURCE_SHA}"

  echo "[8/8] Verifying public health, version, and exact SHA..."
  WEBAPP_URL="${SOLIDBASE_WEBAPP_URL:-https://script.google.com/macros/s/${SOLIDBASE_DEPLOYMENT_ID}/exec}"
  HEALTH_JSON="$(curl --fail --silent --show-error --location --max-time 60 "${WEBAPP_URL}?healthz=1")"
  "$NODE_BIN" -e '
    const data = JSON.parse(process.argv[1]);
    const expectedSha = process.argv[2];
    const expectedVersion = process.argv[3];
    if (data.ok !== true) throw new Error("health ok=false");
    if (data.sourceSha !== expectedSha) throw new Error(`SHA mismatch: ${data.sourceSha} != ${expectedSha}`);
    if (data.version !== expectedVersion) throw new Error(`version mismatch: ${data.version} != ${expectedVersion}`);
    if (data.releaseVerified !== true) throw new Error("releaseVerified=false");
  ' "$HEALTH_JSON" "$SOURCE_SHA" "$APP_VERSION"
fi

echo ""
echo -e "${GREEN}═══════════════════════════════════════${NC}"
echo -e "${GREEN}  Deploy complete!${NC}"
echo -e "${GREEN}═══════════════════════════════════════${NC}"
