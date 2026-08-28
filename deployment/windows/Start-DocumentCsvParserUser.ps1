[CmdletBinding()]
param(
    [string]$AppPath = "C:\Apps\document-csv-parser",
    [string]$PythonPath = "C:\Apps\Python312\python.exe",
    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"

$logs = Join-Path $AppPath "logs"
New-Item -ItemType Directory -Force -Path $logs | Out-Null

$existingListener = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    Select-Object -First 1
if ($existingListener) {
    "Port $Port is already listening on process $($existingListener.OwningProcess)." |
        Out-File -FilePath (Join-Path $logs "scheduled-uvicorn.log") -Append -Encoding utf8
    exit 0
}

$env:PYTHONUNBUFFERED = "1"
Set-Location -LiteralPath $AppPath

$ErrorActionPreference = "Continue"
& $PythonPath -m uvicorn app:app --host 0.0.0.0 --port $Port --workers 1 --proxy-headers --forwarded-allow-ips=* *> (Join-Path $logs "scheduled-uvicorn.log")
