[CmdletBinding()]
param(
    [string]$BaseUrl = "http://localhost:8000",
    [int]$TimeoutSeconds = 10
)

$ErrorActionPreference = "Stop"

$rootUrl = $BaseUrl.TrimEnd("/")
$paths = @("/health", "/report-1", "/report-2", "/report-3")
$failed = $false
$results = foreach ($path in $paths) {
    $url = "$rootUrl$path"
    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec $TimeoutSeconds
        if ($response.StatusCode -lt 200 -or $response.StatusCode -ge 300) {
            throw "Unexpected status code $($response.StatusCode)"
        }

        if ($path -eq "/health") {
            $payload = $response.Content | ConvertFrom-Json
            if ($payload.status -ne "ok") {
                throw "Unexpected health payload: $($response.Content)"
            }
        }

        [PSCustomObject]@{
            Url = $url
            StatusCode = $response.StatusCode
            Result = "OK"
        }
    }
    catch {
        $failed = $true
        [PSCustomObject]@{
            Url = $url
            StatusCode = ""
            Result = "FAILED: $($_.Exception.Message)"
        }
    }
}

$results | Format-Table -AutoSize

if ($failed) {
    exit 1
}
