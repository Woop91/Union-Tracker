#!/bin/bash
# ═══════════════════════════════════════════════
# PUSH SCRIPT — Web Dashboard to DDS-Dashboard
# Run this in your local terminal (PowerShell or Git Bash)
# ═══════════════════════════════════════════════

REPO_DIR="$HOME/Documents/DDS-Dashboard"
BRANCH="web-dashboard"

echo "═══════════════════════════════════════"
echo "  Pushing Web Dashboard to GitHub"
echo "═══════════════════════════════════════"

# Step 1: Navigate to repo
cd "$REPO_DIR" || {
  echo "ERROR: Repo not found at $REPO_DIR"
  echo "Clone it first: git clone https://github.com/Woop91/DDS-Dashboard.git $REPO_DIR"
  exit 1
}

# Step 2: Fetch latest
echo "[1/5] Fetching latest..."
git fetch origin

# Step 3: Create or switch to branch
if git show-ref --quiet refs/heads/$BRANCH; then
  echo "[2/5] Switching to existing branch: $BRANCH"
  git checkout $BRANCH
  git pull origin $BRANCH
else
  echo "[2/5] Creating new branch: $BRANCH"
  git checkout -b $BRANCH
fi

# Step 4: Copy files (user needs to have downloaded them)
echo "[3/5] Ready for files."
echo ""
echo "  Copy the downloaded files into: $REPO_DIR"
echo "  Files needed:"
echo "    - WebApp.gs"
echo "    - Auth.gs"
echo "    - ConfigReader.gs"
echo "    - DataService.gs"
echo "    - index.html"
echo "    - styles.html"
echo "    - auth_view.html"
echo "    - steward_view.html"
echo "    - member_view.html"
echo "    - error_view.html"
echo "    - AI_REFERENCE.md"
echo "    - README.md"
echo ""
read -p "Press Enter once files are copied..."

# Step 5: Stage and commit
echo "[4/5] Staging files..."
git add -A
git status

echo ""
read -p "Looks good? Press Enter to commit and push..."

git commit -m "feat: web dashboard Phase 1 — auth, config, data service, steward + member views

- WebApp.gs: doGet() with auth routing and role detection
- Auth.gs: Google SSO + magic link (7-day expiry, 30-day session)
- ConfigReader.gs: reads Config tab, caches 6hrs
- DataService.gs: dynamic column lookup, role-based data access
- Steward view: Mono Signal theme, KPIs, case list, filters
- Member view: Glass Depth theme, grievance card, timeline, contact
- White-label: all org-specific values from Config tab
- AI_REFERENCE.md: persistent LLM reference document"

# Step 6: Push
echo "[5/5] Pushing to origin/$BRANCH..."
git push origin $BRANCH

echo ""
echo "═══════════════════════════════════════"
echo "  Done! Branch pushed: $BRANCH"
echo "═══════════════════════════════════════"
echo ""
echo "Next: Open GitHub and verify at:"
echo "  https://github.com/Woop91/DDS-Dashboard/tree/$BRANCH"
