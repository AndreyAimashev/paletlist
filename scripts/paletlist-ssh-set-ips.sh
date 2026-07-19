#!/usr/bin/env bash
# Rebuild UFW SSH allow rules for the given IPv4 list.
# Usage: paletlist-ssh-set-ips 1.2.3.4 [5.6.7.8 ...]
set -euo pipefail

if [ "$#" -lt 1 ]; then
  echo "Usage: $0 <ip> [ip...]" >&2
  exit 1
fi

if ! command -v ufw >/dev/null 2>&1; then
  echo "ufw not installed" >&2
  exit 1
fi

IPS=()
for ip in "$@"; do
  if [[ "$ip" =~ ^([0-9]{1,3}\.){3}[0-9]{1,3}$ ]]; then
    IPS+=("$ip")
  else
    echo "Skip invalid IP: $ip" >&2
  fi
done

if [ "${#IPS[@]}" -eq 0 ]; then
  echo "No valid IPs" >&2
  exit 1
fi

# Delete existing SSH allow rules (IPv4/IPv6 port 22)
while true; do
  NUM=$(ufw status numbered | sed -n 's/^\[\s*\([0-9][0-9]*\)\].* 22\/tcp.*/\1/p' | head -n1 || true)
  if [ -z "${NUM:-}" ]; then
    break
  fi
  yes | ufw delete "$NUM" >/dev/null
done

for ip in "${IPS[@]}"; do
  ufw allow from "$ip" to any port 22 proto tcp comment 'SSH allowed'
done

# Ensure web stays open
ufw allow 80/tcp comment 'HTTP' >/dev/null 2>&1 || true
ufw allow 443/tcp comment 'HTTPS' >/dev/null 2>&1 || true

ufw status verbose
echo "SSH allowed from: ${IPS[*]}"
