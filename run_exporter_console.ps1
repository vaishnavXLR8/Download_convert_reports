<#!
Console wrapper for exporting PBIX reports and optional PBIP conversion.
Prompts user for:
  1) Workspace Id
Uses repo folders for outputs:
  - PBIX:   .\download_pbix_files
  - PBIP:   .\converted_pbip_files
Runs export_reports.ps1 and optionally runs UI-automation conversion via PBIXtoPBIP_PBITConversion.psm1.

Build to .exe (optional):
  Install-Module PS2EXE -Scope CurrentUser
  Invoke-PS2EXE -InputFile .\run_exporter_console.ps1 -OutputFile .\ReportExportTool.exe -Title "Power BI Report Exporter" -Company "" -Product "" -Copyright "" -IconFile "" -NoConsole:$false -STA
!#>

[CmdletBinding()]
param()

function Read-NonEmpty([string]$prompt) {
  while ($true) {
    $val = Read-Host $prompt
    if ($val) { return $val.Trim('"').Trim() }
    Write-Host "Please enter a value." -ForegroundColor Yellow
  }
}

try {
  # Establish script root (PS2EXE-safe)
  $ScriptRoot = $PSScriptRoot; if (-not $ScriptRoot) { $ScriptRoot = (Get-Location).Path }

  # Hardcoded repo-relative folders
  $pbixFolder = Join-Path -Path $ScriptRoot -ChildPath 'download_pbix_files'
  $pbipFolder = Join-Path -Path $ScriptRoot -ChildPath 'converted_pbip_files'
  New-Item -ItemType Directory -Path $pbixFolder -Force | Out-Null
  New-Item -ItemType Directory -Path $pbipFolder -Force | Out-Null
  Write-Host "Where to find Workspace Id: In Power BI Service, open the workspace and look at the URL. It's the GUID after /groups/ e.g. https://app.powerbi.com/groups/<WorkspaceId>/..." -ForegroundColor DarkCyan
  $workspaceId = Read-NonEmpty "Enter Workspace Id (GUID)"

  # Normalize and ensure directories
  $pbixFolder = (Resolve-Path -LiteralPath (New-Item -ItemType Directory -Path $pbixFolder -Force)).Path
  $pbipFolder = (Resolve-Path -LiteralPath (New-Item -ItemType Directory -Path $pbipFolder -Force)).Path

  Write-Host "\nPBIX folder : $pbixFolder" -ForegroundColor Cyan
  Write-Host "PBIP folder : $pbipFolder" -ForegroundColor Cyan
  Write-Host "WorkspaceId : $workspaceId" -ForegroundColor Cyan

  # Run export
  Write-Host "\nStarting export..." -ForegroundColor Green
  $exportScript = Join-Path -Path $ScriptRoot -ChildPath "export_reports.ps1"
  if (-not (Test-Path $exportScript)) { throw "Cannot find export_reports.ps1 next to this script." }

  try {
    & $exportScript -WorkspaceId $workspaceId
  }
  catch {
    Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press Enter to close..." -ForegroundColor Yellow
    [void] (Read-Host)
    return
  }
  $exportExit = $LASTEXITCODE
  if ($null -ne $exportExit -and $exportExit -ne 0) {
    Write-Host "Export script signaled exit code $exportExit." -ForegroundColor Yellow
  }

  # Ask about conversion
  Write-Host "\nDo you want to convert the downloaded PBIX files to PBIP using UI automation?" -ForegroundColor White
  Write-Host "WARNING: During conversion, do not click or type. Keep the machine idle until it finishes." -ForegroundColor Yellow
  $ans = Read-Host "Proceed? (Y/N)"
  if ($ans -match '^(y|yes)$') {
    # Validate and import conversion module
  $modulePath = Join-Path -Path $ScriptRoot -ChildPath "PBIXtoPBIP_PBITConversion.psm1"
    if (-not (Test-Path $modulePath)) { throw "Cannot find PBIXtoPBIP_PBITConversion.psm1 next to this script." }
    Import-Module $modulePath -Force
    Get-Module PBIXtoPBIP_PBITConversion | Out-Host

  $files = Get-ChildItem -LiteralPath $pbixFolder -Filter *.pbix -File | Sort-Object Name
    if (-not $files) {
      Write-Host "No PBIX files found to convert in $pbixFolder" -ForegroundColor Yellow
    } else {
      $i = 0
      foreach ($f in $files) {
        $i++
        Write-Host "[$i/$($files.Count)] Converting '$($f.Name)'..." -ForegroundColor White
        try {
          # Synchronous call; next file starts only after the previous conversion closes PBIDesktop and returns
          PBIXtoPBIP_PBITConversion -PBIXFilePath $f.FullName -ConversionFileType "PBIP"
          Write-Host "   Done." -ForegroundColor Green
        }
        catch {
          Write-Host "   FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
      }
      Write-Host "\nConversion complete." -ForegroundColor Green
    }
  } else {
    Write-Host "Skipping PBIP conversion." -ForegroundColor Yellow
  }

  Write-Host "\nAll done. Press Enter to close..." -ForegroundColor Cyan
  [void] (Read-Host)
}
catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Press Enter to close..." -ForegroundColor Yellow
  [void] (Read-Host)
}
