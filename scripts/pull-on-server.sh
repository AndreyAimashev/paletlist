#!/usr/bin/env bash
# Manual deploy helper: run on the production host when GitHub Actions SSH is down.
# Usage: sudo bash /var/www/paletlist/scripts/pull-on-server.sh
set -euo pipefail

APP_DIR="${APP_DIR:-/var/www/paletlist}"
cd "$APP_DIR"

if ! git diff --quiet || ! git diff --cached --quiet; then
  git stash push -u -m "manual-deploy-stash-$(date +%s)" || true
fi

git fetch origin
git checkout main
git pull --ff-only origin main
echo "Deployed: $(git log -1 --oneline)"

if ! grep -q "notifyOrderMonitoringLoadStateChanged" admin.html; then
  echo "admin.html: missing Order Monitoring prefetch freeze fix" >&2
  exit 1
fi

if command -v nginx >/dev/null 2>&1; then
  systemctl reload nginx || true
fi

systemctl restart paletlist-api.service || systemctl try-restart paletlist-api.service || true
sleep 1
if systemctl is-active --quiet paletlist-api.service; then
  echo "API service: active"
else
  echo "API service: not active" >&2
  systemctl --no-pager --full status paletlist-api.service || true
  exit 1
fi

echo "Done. Hard-refresh the browser (Ctrl+F5)."
