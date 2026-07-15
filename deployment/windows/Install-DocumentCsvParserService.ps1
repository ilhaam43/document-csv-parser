#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$AppPath = "C:\Apps\document-csv-parser",
    [string]$ServiceName = "document-csv-parser",
    [int]$Port = 8000,
    [string]$NssmPath = "nssm.exe",
    [string]$PythonLauncher = "python",
    [string]$ServiceAccount,
    [securestring]$ServicePassword,
    [switch]$SkipPackageInstall,
    [switch]$SkipFirewall,
    [switch]$NoStart
)

$ErrorActionPreference = "Stop"

function Resolve-Executable {
    param([Parameter(Mandatory = $true)][string]$Command)

    if (Test-Path -LiteralPath $Command) {
        return (Resolve-Path -LiteralPath $Command).Path
    }

    $resolved = Get-Command $Command -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "Could not find '$Command'. Install it or pass a full path."
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

function ConvertFrom-SecureStringToPlainText {
    param([Parameter(Mandatory = $true)][securestring]$Value)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Value)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

if (-not (Test-Path -LiteralPath $AppPath -PathType Container)) {
    throw "Application path does not exist: $AppPath"
}

$AppPath = (Resolve-Path -LiteralPath $AppPath).Path
$requiredItems = @("app.py", "requirements.txt", "templates", "static")
foreach ($relative in $requiredItems) {
    $fullPath = Join-Path $AppPath $relative
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Required application item is missing: $fullPath"
    }
}

$runtimeDirs = @(
    "input-today",
    "output-today",
    "output-today\api",
    "output-outgoing",
    "output-iphone",
    "vlookup-yesterday",
    "logs"
)

foreach ($relative in $runtimeDirs) {
    New-Item -ItemType Directory -Force -Path (Join-Path $AppPath $relative) | Out-Null
}

$venvPython = Join-Path $AppPath ".venv\Scripts\python.exe"

Push-Location $AppPath
try {
    if (-not (Test-Path -LiteralPath $venvPython)) {
        Write-Host "Creating Python virtual environment at $AppPath\.venv"
        Invoke-Native -FilePath $PythonLauncher -Arguments @("-m", "venv", ".venv")
    }

    if (-not $SkipPackageInstall) {
        Write-Host "Installing Python packages from requirements.txt"
        Invoke-Native -FilePath $venvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip")
        Invoke-Native -FilePath $venvPython -Arguments @("-m", "pip", "install", "-r", "requirements.txt")
    }

    Write-Host "Validating application import"
    Invoke-Native -FilePath $venvPython -Arguments @("-c", "import app; assert app.health() == {'status': 'ok'}")
}
finally {
    Pop-Location
}

if (-not $SkipFirewall) {
    $firewallRuleName = "Document CSV Parser API $Port"
    $firewallRule = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
    if (-not $firewallRule) {
        Write-Host "Opening Windows Firewall inbound TCP port $Port"
        New-NetFirewallRule `
            -DisplayName $firewallRuleName `
            -Direction Inbound `
            -Protocol TCP `
            -LocalPort $Port `
            -Action Allow | Out-Null
    }
    else {
        Set-NetFirewallRule -DisplayName $firewallRuleName -Enabled True -Action Allow | Out-Null
        Write-Host "Firewall rule already exists: $firewallRuleName"
    }
}

$nssm = Resolve-Executable -Command $NssmPath
$pythonArgs = "-m uvicorn app:app --host 0.0.0.0 --port $Port --proxy-headers --forwarded-allow-ips=*"
$stdoutLog = Join-Path $AppPath "logs\service-out.log"
$stderrLog = Join-Path $AppPath "logs\service-error.log"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

if (-not $service) {
    Write-Host "Installing Windows service '$ServiceName'"
    Invoke-Native -FilePath $nssm -Arguments @("install", $ServiceName, $venvPython, $pythonArgs)
}
else {
    Write-Host "Updating existing Windows service '$ServiceName'"
}

Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "Application", $venvPython)
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppParameters", $pythonArgs)
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppDirectory", $AppPath)
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppStdout", $stdoutLog)
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppStderr", $stderrLog)
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppEnvironmentExtra", "PYTHONUNBUFFERED=1")
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppRotateFiles", "1")
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppRotateOnline", "1")
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppRotateBytes", "10485760")
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppRestartDelay", "5000")
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "AppExit", "Default", "Restart")
Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "Start", "SERVICE_AUTO_START")

if ($ServiceAccount) {
    $builtInAccounts = @("LocalSystem", "NT AUTHORITY\LocalService", "NT AUTHORITY\NetworkService")
    $isBuiltInAccount = $builtInAccounts -contains $ServiceAccount
    if (-not $ServicePassword -and -not $isBuiltInAccount) {
        throw "Pass -ServicePassword for service account '$ServiceAccount'."
    }

    if ($ServicePassword) {
        $plainPassword = ConvertFrom-SecureStringToPlainText -Value $ServicePassword
        try {
            Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "ObjectName", $ServiceAccount, $plainPassword)
        }
        finally {
            $plainPassword = $null
        }
    }
    else {
        Invoke-Native -FilePath $nssm -Arguments @("set", $ServiceName, "ObjectName", $ServiceAccount)
    }
}

if (-not $NoStart) {
    $service = Get-Service -Name $ServiceName -ErrorAction Stop
    if ($service.Status -eq "Running") {
        Write-Host "Restarting service '$ServiceName'"
        Restart-Service -Name $ServiceName -Force
    }
    else {
        Write-Host "Starting service '$ServiceName'"
        Start-Service -Name $ServiceName
    }

    $healthUrl = "http://localhost:$Port/health"
    $healthy = $false
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline -and -not $healthy) {
        try {
            $response = Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 5
            $healthy = ($response.StatusCode -eq 200 -and $response.Content -match '"status"\s*:\s*"ok"')
        }
        catch {
            Start-Sleep -Seconds 1
        }
    }

    if ($healthy) {
        Write-Host "Health check passed: $healthUrl"
    }
    else {
        Write-Warning "Service was configured, but the local health check did not pass within 30 seconds. Check $stderrLog."
    }
}

$service = Get-Service -Name $ServiceName -ErrorAction Stop
$listenEndpoint = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($listenEndpoint) {
    Write-Host "Port $Port is listening locally."
}
else {
    Write-Warning "Port $Port is not listening locally. Check service status and logs."
}

Write-Host "Service status: $($service.Status)"
Write-Host "Windows service '$ServiceName' is configured for $AppPath on port $Port."
