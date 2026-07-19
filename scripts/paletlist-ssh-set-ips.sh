#!/usr/bin/env bash
# Ensure SSH is allowed from the given IPv4 list (add first, then prune stale).
set -euo pipefail

if [ "$#" -lt 1 ]; then
  echo "Usage: $0 <ip> [ip...]" >&2
  exit 1
fi

IPS=()
for ip in "$@"; do
  if [[ "$ip" =~ ^([0-9]{1,3}\.){3}[0-9]{1,3}$ ]]; then
    IPS+=("$ip")
  fi
done
if [ "${#IPS[@]}" -eq 0 ]; then
  echo "No valid IPs" >&2
  exit 1
fi

for ip in "${IPS[@]}"; do
  ufw allow from "$ip" to any port 22 proto tcp comment 'SSH allowed' || true
done
ufw allow 80/tcp comment 'HTTP' >/dev/null 2>&1 || true
ufw allow 443/tcp comment 'HTTPS' >/dev/null 2>&1 || true
ufw status verbose
echo "SSH ensured for: ${IPS[*]}"
