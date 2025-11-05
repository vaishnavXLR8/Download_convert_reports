# Power BI Workspace Reports Downloader (Python)

CLI tool that downloads all PBIX reports from a given Power BI workspace (group) ID.

- Auth: Azure AD client credentials (service principal)
- APIs used:
  - List reports: `GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports`
  - Export PBIX: `GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports/{reportId}/Export`
- Timeout: 3 minutes per report. If a report fails early or exceeds 3 minutes, it will be reported.

## Prereqs
- Python 3.10+
- AAD App Registration (service principal) with Power BI API access
- Power BI tenant settings must allow service principals and PBIX export where applicable

## Setup
1. Create and fill env vars (you can use a `.env` in this folder):
```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

2. Install dependencies:
```
python -m pip install -r requirements.txt
```

## Usage
```
# Minimal (downloads to ./downloads by default)
python download_reports.py --group-id <workspaceId>

# Specify output directory explicitly (no square brackets)
python download_reports.py --group-id <workspaceId> --output ./downloads
```

Examples (PowerShell):
```
# Default output folder
python download_reports.py --group-id 00000000-0000-0000-0000-000000000000

# Custom output folder
python download_reports.py --group-id 00000000-0000-0000-0000-000000000000 --output .\downloads

# Using short flag
python download_reports.py -g 00000000-0000-0000-0000-000000000000 -o .\out
```

Behavior:
- Downloads each report to the output directory
- Enforces a 3-minute per-report timeout; partial files are removed on timeout or error
- Prints a summary of successes and failures

Notes:
- Some reports cannot be exported to PBIX due to tenant policy or dataset type (e.g., live connection).
- If you see 403/404 errors on export while listing works, check admin settings and workspace permissions.

## PowerShell: Export all reports from a workspace (export_reports.ps1)

Script that downloads all PBIX reports from a workspace using the official Power BI PowerShell modules, optionally converts them to PBIP via Power BI Desktop CLI, and prints file sizes and timings. The built-in "Report Usage Metrics Report" is automatically skipped.

### Prerequisites
- Windows PowerShell 5.1 (recommended)
- Power BI PowerShell modules:
  - Install with: `Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser`
- Permission to access the target workspace and to export PBIX (tenant policy and licensing apply)
- Optional PBIP conversion: Power BI Desktop installed (path to `PBIDesktop.exe` required)

### Parameters
- `-WorkspaceName [string]` Workspace display name (optional)
- `-WorkspaceId [string]` Workspace GUID (optional)
- `-OutputFolder [string]` Destination folder for PBIX files (default: `./downloads`)
- `-ConvertToPBIP` Switch to also convert downloaded PBIX to PBIP
- `-PBIDesktopPath [string]` Full path to PBIDesktop.exe (required when using `-ConvertToPBIP`)

Provide either `-WorkspaceName` or `-WorkspaceId`.

### Usage
```powershell
# Sign in interactively when prompted
.\u005cexport_reports.ps1 -WorkspaceId 00000000-0000-0000-0000-000000000000 -OutputFolder .\downloads

# By name
.\u005cexport_reports.ps1 -WorkspaceName "My Workspace" -OutputFolder .\downloads

# With PBIP conversion
.\u005cexport_reports.ps1 -WorkspaceId 00000000-0000-0000-0000-000000000000 -OutputFolder .\downloads -ConvertToPBIP -PBIDesktopPath "C:\\Program Files\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe"
```

### Behavior
- Exports each report to `<OutputFolder>\\<ReportName>.pbix`
- Skips the built-in report named exactly `Report Usage Metrics Report`
- If `-ConvertToPBIP` is specified, converts to PBIP under `<OutputFolder>\\pbip_files\\<ReportName>`
- Prints per-report status with:
  - Size (MB) of the PBIX file
  - Export time (ms)
  - Optional conversion time (ms)
- Prints a summary table with columns: `Name`, `SizeMB`, `ExportMs`, `ConvertMs`, `Status`

### Errors and troubleshooting
- If an export fails, the script prints detailed error information including exception type, message, and (when available) HTTP response content.
- Common causes:
  - Tenant policy disallows PBIX export
  - User lacks permission on the dataset/workspace
  - Report type doesn’t support export
- For PBIP conversion issues:
  - Ensure `PBIDesktop.exe` path is correct
  - Different Desktop builds may vary CLI flags; the script tries multiple variants

---

## PowerShell: GUI PBIX → PBIP/PBIT converter (test.psm1)

Module providing `PBIXtoPBIP_PBITConversion` that opens a PBIX in Power BI Desktop and uses keyboard automation to Save As PBIP or PBIT.

### Prerequisites
- Windows PowerShell 5.1
- Windows Power BI Desktop installed and associated with `.pbix` files
- Keep your mouse/keyboard idle during automation; the script sends keystrokes to the active window

### Import and usage
```powershell
Import-Module ".\test.psm1" -Force

# Convert to PBIP
PBIXtoPBIP_PBITConversion -PBIXFilePath ".\downloads\MyReport.pbix" -ConversionFileType "PBIP"

# Or convert to PBIT
PBIXtoPBIP_PBITConversion -PBIXFilePath ".\downloads\MyReport.pbix" -ConversionFileType "PBIT"
```

### Important notes
- Output path is currently hardcoded in the module to:
  `C:\Users\VaishnavKamartiMAQSo\Desktop\VS code explorations\DeDeuplication\Download_convert_reports\downloads\pbip_files`
  - Edit this string in `test.psm1` to change the destination folder.
- The script waits for the PBIDesktop window to appear, brings it to foreground, then performs File → Save As and selects PBIP/PBIT via keyboard.
- If your UI language or Power BI Desktop layout differs significantly, the keystroke sequence might need adjustments.

### Troubleshooting
- Stuck on “report is still loading…”: the module now has a bounded wait. If your reports load slowly, increase the timeout inside the script.
- Sign-in prompts: the script attempts to close the AAD broker process. If sign-in is required, sign in manually, then retry.
- SendKeys not acting on Desktop: ensure Power BI Desktop is in the foreground; avoid interacting with other windows during the run.
- Run in Windows PowerShell (5.1) rather than PowerShell 7 if SendKeys behaves inconsistently.

### Roadmap (optional improvements)
- Parameterize the destination path instead of a hardcoded string
- Add option to batch-convert a folder of PBIX files
- Use UI automation libraries for more robust element targeting
