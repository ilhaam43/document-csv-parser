#Requires -RunAsAdministrator
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ServiceName = "document-csv-parser",
    [int]$Port = 8000,
    [string]$NssmPath = "nssm.exe",
    [switch]$RemoveFirewallRule
)

$ErrorActionPreference = "Stop"

function Resolve-Executable {
    param([Parameter(Mandatory = $true)][string]$Command)

    if (Test-Path -LiteralPath $Command) {
        return (Resolve-Path -LiteralPath $Command).Path
    }

    $resolved = Get-Command $Command -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "Could not find '$Command'. Install it, pass a full path, or remove the service manually."
    }

    return $resolved.Source
}

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
    Write-Host "Service is already absent: $ServiceName"
}
else {
    if ($PSCmdlet.ShouldProcess($ServiceName, "stop and remove Windows service")) {
        if ($service.Status -ne "Stopped") {
            Write-Host "Stopping service '$ServiceName'"
            Stop-Service -Name $ServiceName -Force
            $service.WaitForStatus("Stopped", "00:00:30")
        }

        $nssm = Resolve-Executable -Command $NssmPath
        Write-Host "Removing service '$ServiceName'"
        Invoke-Native -FilePath $nssm -Arguments @("remove", $ServiceName, "confirm")
    }
}

if ($RemoveFirewallRule) {
    $firewallRuleName = "Document CSV Parser API $Port"
    $firewallRules = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
    if ($firewallRules) {
        if ($PSCmdlet.ShouldProcess($firewallRuleName, "remove Windows Firewall rule")) {
            $firewallRules | Remove-NetFirewallRule
            Write-Host "Removed firewall rule: $firewallRuleName"
        }
    }
    else {
        Write-Host "Firewall rule is already absent: $firewallRuleName"
    }
}

Write-Host "Application files, virtual environment, uploads, outputs, and logs were left in place."
