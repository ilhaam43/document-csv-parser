#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$SourcePath,
    [string]$AppPath = "C:\Apps\document-csv-parser",
    [string]$ServiceName = "document-csv-parser",
    [int]$Port = 8000,
    [string]$LogPath,
    [switch]$SkipChocolateyInstall
)

$ErrorActionPreference = "Stop"

if (-not $SourcePath) {
    $SourcePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..")).Path
}

if (-not $LogPath) {
    $safeTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $LogPath = Join-Path $SourcePath "deployment\windows\deploy-$safeTimestamp.log"
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LogPath) | Out-Null
Start-Transcript -Path $LogPath -Append | Out-Null

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [int[]]$SuccessExitCodes = @(0)
    )

    Write-Host "> $FilePath $($Arguments -join ' ')"
    & $FilePath @Arguments
    if ($SuccessExitCodes -notcontains $LASTEXITCODE) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

function Test-Python312 {
    $pyLauncher = Join-Path $env:WINDIR "py.exe"
    if (-not (Test-Path -LiteralPath $pyLauncher)) {
        return $false
    }

    & $pyLauncher -3.12 --version *> $null
    return $LASTEXITCODE -eq 0
}

function Resolve-Python312 {
    $pyLauncher = Join-Path $env:WINDIR "py.exe"
    if (Test-Path -LiteralPath $pyLauncher) {
        $pythonExe = (& $pyLauncher -3.12 -c "import sys; print(sys.executable)") | Select-Object -Last 1
        if ($LASTEXITCODE -eq 0 -and $pythonExe) {
            return $pythonExe.Trim()
        }
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        $version = (& $python.Source -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')") | Select-Object -Last 1
        if ($LASTEXITCODE -eq 0 -and $version -eq "3.12") {
            return $python.Source
        }
    }

    throw "Python 3.12 was not found after prerequisite installation."
}

function Resolve-Executable {
    param([Parameter(Mandatory = $true)][string]$Command)

    $resolved = Get-Command $Command -ErrorAction SilentlyContinue
    if ($resolved) {
        return $resolved.Source
    }

    $chocoShim = Join-Path $env:ProgramData "chocolatey\bin\$Command.exe"
    if (Test-Path -LiteralPath $chocoShim) {
        return $chocoShim
    }

    throw "Could not find executable: $Command"
}

try {
    $SourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
    Write-Host "Source path: $SourcePath"
    Write-Host "App path: $AppPath"
    Write-Host "Log path: $LogPath"

    if (-not $SkipChocolateyInstall) {
        $choco = Resolve-Executable -Command "choco"
        $packages = @()
        if (-not (Test-Python312)) {
            $packages += "python312"
        }
        if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
            $packages += "git"
        }
        if (-not (Get-Command nssm -ErrorAction SilentlyContinue)) {
            $packages += "nssm"
        }
        $packages += "vcredist140"

        if ($packages.Count -gt 0) {
            Invoke-Native -FilePath $choco -Arguments @("install") + $packages + @("-y", "--no-progress")
        }
        else {
            Write-Host "Chocolatey prerequisites already appear to be installed."
        }
    }

    $appPathExists = Test-Path -LiteralPath $AppPath -PathType Container
    if (-not $appPathExists) {
        New-Item -ItemType Directory -Force -Path $AppPath | Out-Null
    }
    $resolvedAppPath = (Resolve-Path -LiteralPath $AppPath).Path

    if ($SourcePath -ne $resolvedAppPath) {
        Write-Host "Copying application files to $resolvedAppPath"
        $robocopyArgs = @(
            $SourcePath,
            $resolvedAppPath,
            "/E",
            "/XD",
            ".git",
            ".venv",
            "__pycache__",
            ".pytest_cache",
            "input-today",
            "output-today",
            "output-outgoing",
            "output-iphone",
            "vlookup-yesterday",
            "input-pipeline",
            "output-pipeline",
            "reference-pipeline",
            "/XF",
            "*.pyc",
            "~$*.xlsx"
        )
        Invoke-Native -FilePath "robocopy.exe" -Arguments $robocopyArgs -SuccessExitCodes @(0, 1, 2, 3, 4, 5, 6, 7)
    }
    else {
        Write-Host "Source and app path are the same; skipping copy."
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
        New-Item -ItemType Directory -Force -Path (Join-Path $resolvedAppPath $relative) | Out-Null
    }

    $python312 = Resolve-Python312
    $venvPython = Join-Path $resolvedAppPath ".venv\Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $venvPython)) {
        Invoke-Native -FilePath $python312 -Arguments @("-m", "venv", (Join-Path $resolvedAppPath ".venv"))
    }

    Push-Location $resolvedAppPath
    try {
        Invoke-Native -FilePath $venvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip")
        Invoke-Native -FilePath $venvPython -Arguments @("-m", "pip", "install", "-r", "requirements.txt")

        $nssm = Resolve-Executable -Command "nssm"
        $installScript = Join-Path $resolvedAppPath "deployment\windows\Install-DocumentCsvParserService.ps1"
        & $installScript `
            -AppPath $resolvedAppPath `
            -ServiceName $ServiceName `
            -Port $Port `
            -NssmPath $nssm `
            -SkipPackageInstall
        if ($LASTEXITCODE -ne 0) {
            throw "Service installer failed with exit code $LASTEXITCODE."
        }

        $testScript = Join-Path $resolvedAppPath "deployment\windows\Test-DocumentCsvParserDeployment.ps1"
        & $testScript -BaseUrl "http://localhost:$Port"
        if ($LASTEXITCODE -ne 0) {
            throw "Deployment test failed with exit code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }

    Write-Host "Deployment completed successfully."
}
finally {
    Stop-Transcript | Out-Null
}
