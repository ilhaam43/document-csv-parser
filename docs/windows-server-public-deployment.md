# Windows Server Public Deployment

This deploys the FastAPI app natively on Windows Server, runs it as a background Windows service with NSSM, and exposes it at:

```text
http://<PUBLIC_SERVER_IP>:8000
```

The app is intentionally open with no login for this first deployment. Anyone who can reach the URL can use the upload pages.

## Prerequisites

Install these on the Windows Server first:

- Python 3.12 x64, available as `python` in PowerShell
- Git
- Microsoft Excel, if the report flow needs Excel COM automation
- NSSM, available as `nssm.exe` in `PATH` or from a known full path
- Microsoft Visual C++ runtime, if any Python package requires it

Also confirm the hosting provider firewall, router, or security group allows inbound TCP `8000`. The Windows Firewall rule alone is not enough if the provider blocks the port.

## Prepare The App Directory

Use a stable path for the service:

```powershell
New-Item -ItemType Directory -Force C:\Apps | Out-Null
cd C:\Apps
git clone <REPO_URL> document-csv-parser
cd C:\Apps\document-csv-parser
```

If the code is already copied to the server, place it at:

```text
C:\Apps\document-csv-parser
```

Keep `templates` and `static` inside this directory. The web pages depend on those folders.

## Optional Local Preflight

Before installing the service, you can run the app directly from the server console:

```powershell
cd C:\Apps\document-csv-parser
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\.venv\Scripts\python.exe -m uvicorn app:app --host 0.0.0.0 --port 8000
```

In another PowerShell window on the server:

```powershell
Invoke-RestMethod http://localhost:8000/health
```

Stop the preflight server with `Ctrl+C` before installing or starting the Windows service.

## Install Or Update The Windows Service

Open PowerShell as Administrator, then run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

.\deployment\windows\Install-DocumentCsvParserService.ps1 `
  -AppPath C:\Apps\document-csv-parser `
  -ServiceName document-csv-parser `
  -Port 8000 `
  -NssmPath nssm.exe
```

If NSSM is not in `PATH`, pass the full path:

```powershell
.\deployment\windows\Install-DocumentCsvParserService.ps1 `
  -AppPath C:\Apps\document-csv-parser `
  -NssmPath C:\Tools\nssm\nssm.exe
```

Recommended for public use: run the service under a dedicated local user. Create the user, grant it modify access to the app folder, then pass the account to the installer:

```powershell
$servicePassword = Read-Host "Service account password" -AsSecureString
New-LocalUser `
  -Name svc-doc-csv-parser `
  -Password $servicePassword `
  -PasswordNeverExpires

icacls C:\Apps\document-csv-parser /grant "svc-doc-csv-parser:(OI)(CI)M"

.\deployment\windows\Install-DocumentCsvParserService.ps1 `
  -AppPath C:\Apps\document-csv-parser `
  -ServiceAccount ".\svc-doc-csv-parser" `
  -ServicePassword $servicePassword
```

The install script:

- creates `.venv`
- installs `requirements.txt`
- creates runtime folders: `input-today`, `output-today`, `output-outgoing`, `output-iphone`, `vlookup-yesterday`, and `logs`
- opens Windows Firewall inbound TCP `8000`
- installs or updates the NSSM service named `document-csv-parser`
- configures service stdout/stderr logs in `logs`
- sets the service to restart on failure and start automatically
- starts the service and checks `http://localhost:8000/health`

To update service settings without reinstalling Python packages:

```powershell
.\deployment\windows\Install-DocumentCsvParserService.ps1 `
  -AppPath C:\Apps\document-csv-parser `
  -SkipPackageInstall
```

## Verify Locally

From the server:

```powershell
Invoke-RestMethod http://localhost:8000/health

.\deployment\windows\Test-DocumentCsvParserDeployment.ps1 `
  -BaseUrl http://localhost:8000
```

Expected health response:

```json
{"status":"ok"}
```

## Verify Public Access

From another machine:

```powershell
Invoke-RestMethod http://<PUBLIC_SERVER_IP>:8000/health
```

Then open these pages in a browser:

```text
http://<PUBLIC_SERVER_IP>:8000/report-1
http://<PUBLIC_SERVER_IP>:8000/report-2
http://<PUBLIC_SERVER_IP>:8000/report-3
```

You can also run:

```powershell
.\deployment\windows\Test-DocumentCsvParserDeployment.ps1 `
  -BaseUrl http://<PUBLIC_SERVER_IP>:8000
```

## Service Operations

```powershell
Get-Service document-csv-parser
Restart-Service document-csv-parser
Stop-Service document-csv-parser
Start-Service document-csv-parser
```

For a fuller status check:

```powershell
.\deployment\windows\Get-DocumentCsvParserStatus.ps1 `
  -AppPath C:\Apps\document-csv-parser `
  -BaseUrl http://localhost:8000
```

Logs are written to:

```text
C:\Apps\document-csv-parser\logs\service-out.log
C:\Apps\document-csv-parser\logs\service-error.log
```

If you did not pass `-ServiceAccount`, NSSM defaults to the LocalSystem account. For public use, change the service to a dedicated local user; do not run the service as Administrator.

To remove the Windows service while leaving the application files and generated outputs in place:

```powershell
.\deployment\windows\Remove-DocumentCsvParserService.ps1 `
  -ServiceName document-csv-parser `
  -NssmPath nssm.exe
```

To remove the service and the Windows Firewall rule:

```powershell
.\deployment\windows\Remove-DocumentCsvParserService.ps1 `
  -ServiceName document-csv-parser `
  -NssmPath nssm.exe `
  -RemoveFirewallRule
```

## Upload Workflow Test

After the pages load, run an end-to-end upload test in the browser:

1. Open `/report-1`, upload the raw CSV and yesterday cleaned workbook, then confirm a download link is returned.
2. Open `/report-2`, upload the tracking workbook, log update CSV, and previous ongoing workbook, then confirm a download link is returned.
3. Open `/report-3`, upload the tracking workbook and previous iPhone workbook, then confirm a download link is returned.
4. Open a returned download link from a different public machine.
5. Restart Windows Server and confirm the service starts automatically.

## Security Notes

This deployment exposes the app by raw public IP with no HTTPS and no authentication.

Minimum hardening before public use:

- Keep Windows Server patched.
- Run the service under a dedicated Windows user instead of Administrator.
- Monitor disk usage in the output folders.
- Keep uploaded reference workbooks and generated outputs out of source control.
- Limit upload size at a proxy or network layer if possible.

Recommended later upgrade:

- Put IIS or another reverse proxy in front of Uvicorn.
- Add HTTPS with a real domain name.
- Add authentication before sharing the URL broadly.
