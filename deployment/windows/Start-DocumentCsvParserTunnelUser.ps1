[CmdletBinding()]
param(
    [string]$AppPath = "C:\Apps\document-csv-parser",
    [string]$CloudflaredPath = "C:\Apps\cloudflared.exe",
    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"

$logs = Join-Path $AppPath "logs"
New-Item -ItemType Directory -Force -Path $logs | Out-Null

$existing = Get-Process -Name cloudflared -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -eq $CloudflaredPath } |
    Select-Object -First 1
if ($existing) {
    "cloudflared is already running on process $($existing.Id)." |
        Out-File -FilePath (Join-Path $logs "scheduled-cloudflared.log") -Append -Encoding utf8
    exit 0
}

Set-Location -LiteralPath $AppPath

$ErrorActionPreference = "Continue"
& $CloudflaredPath tunnel `
    --url "http://127.0.0.1:$Port" `
    --no-autoupdate `
    --proxy-connect-timeout 2m `
    --proxy-keepalive-timeout 5m `
    --proxy-tcp-keepalive 1m `
    --retries 10 `
    --loglevel info `
    *> (Join-Path $logs "scheduled-cloudflared.log")
