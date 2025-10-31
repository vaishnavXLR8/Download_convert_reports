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
