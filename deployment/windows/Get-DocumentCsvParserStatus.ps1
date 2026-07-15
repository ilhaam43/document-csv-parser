[CmdletBinding()]
param(
    [string]$AppPath = "C:\Apps\document-csv-parser",
    [string]$ServiceName = "document-csv-parser",
    [int]$Port = 8000,
    [string]$BaseUrl = "http://localhost:8000",
    [int]$LogTail = 20
)

$ErrorActionPreference = "Continue"

$rootUrl = $BaseUrl.TrimEnd("/")
$healthUrl = "$rootUrl/health"
$firewallRuleName = "Document CSV Parser API $Port"

$serviceInfo = Get-CimInstance Win32_Service -Filter "Name='$ServiceName'" -ErrorAction SilentlyContinue
if ($serviceInfo) {
    [PSCustomObject]@{
        ServiceName = $serviceInfo.Name
        State = $serviceInfo.State
        StartMode = $serviceInfo.StartMode
        Account = $serviceInfo.StartName
        PathName = $serviceInfo.PathName
    } | Format-List
}
else {
    Write-Warning "Service not found: $ServiceName"
}

$listeners = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($listeners) {
    Write-Host "Listening endpoints on port ${Port}:"
    $listeners | Select-Object LocalAddress, LocalPort, OwningProcess | Format-Table -AutoSize
}
else {
    Write-Warning "No local listener found on port $Port."
}

$firewallRules = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
if ($firewallRules) {
    Write-Host "Firewall rule:"
    $firewallRules | Select-Object DisplayName, Enabled, Direction, Action | Format-Table -AutoSize
}
else {
    Write-Warning "Firewall rule not found: $firewallRuleName"
}

try {
    $response = Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 10
    Write-Host "Health check: HTTP $($response.StatusCode) $($response.Content)"
}
catch {
    Write-Warning "Health check failed for ${healthUrl}: $($_.Exception.Message)"
}

$stdoutLog = Join-Path $AppPath "logs\service-out.log"
$stderrLog = Join-Path $AppPath "logs\service-error.log"
foreach ($logPath in @($stdoutLog, $stderrLog)) {
    if (Test-Path -LiteralPath $logPath) {
        Write-Host ""
        Write-Host "Last $LogTail lines from $logPath"
        Get-Content -LiteralPath $logPath -Tail $LogTail
    }
    else {
        Write-Warning "Log file not found: $logPath"
    }
}
