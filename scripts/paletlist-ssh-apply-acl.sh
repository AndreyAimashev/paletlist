#!/usr/bin/env bash
# Apply UFW + sshd Match Address ACL from /etc/paletlist/{ssh-owner,ssh-audit}-ips.txt
set -euo pipefail
OWNER_FILE=/etc/paletlist/ssh-owner-ips.txt
AUDIT_FILE=/etc/paletlist/ssh-audit-ips.txt
PIN_FILE=/etc/paletlist/ssh-pinned-ips.txt
mkdir -p /etc/paletlist
touch "$OWNER_FILE" "$AUDIT_FILE"
chmod 600 "$OWNER_FILE" "$AUDIT_FILE" 2>/dev/null || true
cat "$OWNER_FILE" "$AUDIT_FILE" | awk 'NF && !seen[$0]++' > "$PIN_FILE"
chmod 600 "$PIN_FILE"
while read -r ip; do
  [ -n "$ip" ] || continue
  ufw allow from "$ip" to any port 22 proto tcp comment 'SSH allowed' || true
done < "$PIN_FILE"
ufw allow 80/tcp >/dev/null 2>&1 || true
ufw allow 443/tcp >/dev/null 2>&1 || true
owner_csv=$(awk 'NF' "$OWNER_FILE" | paste -sd, -)
audit_csv=$(awk 'NF' "$AUDIT_FILE" | paste -sd, -)
SSHD=/etc/ssh/sshd_config
sed -i '/# BEGIN ssh address ACL/,/# END ssh address ACL/d' "$SSHD"
{
  echo
  echo '# BEGIN ssh address ACL'
  if [ -n "$audit_csv" ]; then
    echo "# Audit IPs: only security-audit (SFTP). No root."
    echo "Match Address ${audit_csv}"
    echo "    AllowUsers security-audit"
  fi
  if [ -n "$owner_csv" ]; then
    echo "# Owner IPs: root (key) + security-audit"
    echo "Match Address ${owner_csv}"
    echo "    AllowUsers root security-audit"
  fi
  echo '# END ssh address ACL'
} >> "$SSHD"
mkdir -p /run/sshd
sshd -t
# Ubuntu uses unit "ssh"; "sshd" often does not exist.
if systemctl reload ssh 2>/dev/null; then
  :
elif systemctl reload sshd 2>/dev/null; then
  :
elif systemctl restart ssh 2>/dev/null; then
  :
elif systemctl restart sshd 2>/dev/null; then
  :
else
  echo "Warning: could not reload SSH daemon" >&2
fi
echo "ACL applied. owner=${owner_csv:-none} audit=${audit_csv:-none}"
