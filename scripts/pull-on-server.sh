#!/usr/bin/env bash
# Deploy helper: git pull + restart API. Used by HTTPS /api/deploy webhook and manually.
# Usage: sudo bash /usr/local/sbin/paletlist-deploy
#    or: sudo bash /var/www/paletlist/scripts/pull-on-server.sh
set -euo pipefail

APP_DIR="${APP_DIR:-/var/www/paletlist}"
cd "$APP_DIR"

if ! git diff --quiet || ! git diff --cached --quiet; then
  git stash push -u -m "deploy-stash-$(date +%s)" || true
fi

git fetch origin
git checkout main
git pull --ff-only origin main
echo "Deployed: $(git log -1 --oneline)"

grep -q "build_blank_arnest_unirus_pallet_sheets_pdf" api_server.py || {
  echo "api_server.py: нет генератора PDF паллетных листов" >&2
  exit 1
}
grep -q "/api/deploy" api_server.py || {
  echo "api_server.py: нет webhook деплоя" >&2
  exit 1
}

if [ -f requirements.txt ]; then
  if ! python3 -m pip --version 2>/dev/null; then
    apt-get update -qq
    apt-get install -y python3-pip
  fi
  apt-get update -qq
  DEBIAN_FRONTEND=noninteractive apt-get install -y \
    libpango-1.0-0 libpangocairo-1.0-0 libpangoft2-1.0-0 libcairo2 \
    libharfbuzz0b libharfbuzz-subset0 libfontconfig1 libgdk-pixbuf-2.0-0 \
    libglib2.0-0 shared-mime-info || true
  python3 -m pip install --break-system-packages -r requirements.txt || true
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

echo "Done."
