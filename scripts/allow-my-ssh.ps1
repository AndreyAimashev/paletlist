# Unlock SSH from this PC's current public IP (works from any network).
# Usage: right-click -> Run with PowerShell, or:
#   powershell -ExecutionPolicy Bypass -File scripts\allow-my-ssh.ps1

$ErrorActionPreference = "Stop"
$ServerHost = if ($env:PALETLIST_SERVER_HOST) { $env:PALETLIST_SERVER_HOST } else { "185.244.172.114" }
$SecretFile = Join-Path $env:USERPROFILE ".paletlist\ssh-allow.secret"
$Secret = $env:PALETLIST_SSH_ALLOW_SECRET

if (-not $Secret -and (Test-Path $SecretFile)) {
  $Secret = (Get-Content -Path $SecretFile -Raw).Trim()
}

if (-not $Secret) {
  Write-Host "Secret not found."
  Write-Host "Create file: $SecretFile"
  Write-Host "Put the SSH-allow secret on one line (same as on the server)."
  exit 1
}

Write-Host "Detecting public IP..."
$MyIp = (Invoke-RestMethod -Uri "https://api.ipify.org" -TimeoutSec 15).Trim()
Write-Host "Public IP: $MyIp"

$uri = "https://$ServerHost/api/ssh-allow"
Write-Host "Requesting SSH allow on $uri ..."

try {
  $resp = Invoke-RestMethod -Method Post -Uri $uri -TimeoutSec 30 `
    -Headers @{ Authorization = "Bearer $Secret" } `
    -ContentType "application/json; charset=utf-8" `
    -Body (@{ ip = $MyIp } | ConvertTo-Json) `
    -SkipCertificateCheck
} catch {
  # Windows PowerShell 5 may not support -SkipCertificateCheck
  if ($_.Exception.Message -match "SkipCertificateCheck|parameter") {
    add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAll : ICertificatePolicy {
  public bool CheckValidationResult(ServicePoint s, X509Certificate c, WebRequest r, int p) { return true; }
}
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAll
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $resp = Invoke-RestMethod -Method Post -Uri $uri -TimeoutSec 30 `
      -Headers @{ Authorization = "Bearer $Secret" } `
      -ContentType "application/json; charset=utf-8" `
      -Body (@{ ip = $MyIp } | ConvertTo-Json)
  } else {
    throw
  }
}

Write-Host ($resp | ConvertTo-Json -Compress)
Write-Host ""
Write-Host "OK. Now connect (try alt port if 22 is filtered):"
Write-Host "  ssh -p 2222 -i `$env:USERPROFILE\.ssh\paletlist_ed25519 root@$ServerHost"
Write-Host "  ssh -i `$env:USERPROFILE\.ssh\paletlist_ed25519 root@$ServerHost"
