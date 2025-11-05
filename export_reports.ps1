<#!
Downloads all reports in a Power BI workspace to PBIX using Power BI PowerShell modules,
identifies connection types, handles failures with clear errors, and optionally converts to PBIP
using Power BI Desktop CLI. Tracks per-report and total durations.

Prereqs:
- Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser
- Ensure you can sign in with Connect-PowerBIServiceAccount and have access to the workspace
- For optional PBIP conversion: Power BI Desktop installed and PBIDesktop.exe path available

Usage examples (PowerShell):
  # By workspace name
  .\export_reports.ps1 -WorkspaceName "My Workspace" -OutputFolder .\downloads

  # By workspace id
  .\export_reports.ps1 -WorkspaceId 00000000-0000-0000-0000-000000000000 -OutputFolder .\downloads

  # With PBIP conversion
  .\export_reports.ps1 -WorkspaceName "My Workspace" -OutputFolder .\downloads -ConvertToPBIP -PBIDesktopPath "C:\\Program Files\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe"
!#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [string]$WorkspaceName,

  [Parameter(Mandatory=$false)]
  [string]$WorkspaceId,

  [Parameter(Mandatory=$false)]
  [string]$OutputFolder = ".\downloads",

  [switch]$ConvertToPBIP,

  [Parameter(Mandatory=$false)]
  [string]$PBIDesktopPath
)

function Ensure-Module {
  param([string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Host "Installing module: $Name" -ForegroundColor Yellow
    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module $Name -ErrorAction Stop
}

function Sanitize-FileName {
  param([string]$Name)
  $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
  $regex = "[" + [Regex]::Escape($invalid) + "]"
  return [Regex]::Replace($Name, $regex, '_').Trim()
}

function Format-Error {
  param([System.Management.Automation.ErrorRecord]$ErrorRecord)
  $lines = @()
  if ($ErrorRecord.Exception) {
    $lines += ("Exception: " + $ErrorRecord.Exception.GetType().FullName)
    $lines += ("Message  : " + $ErrorRecord.Exception.Message)
    # Include innermost exception message if different
    $inner = $ErrorRecord.Exception
    while ($inner.InnerException) { $inner = $inner.InnerException }
    if ($inner -and $inner -ne $ErrorRecord.Exception -and $inner.Message) {
      $lines += ("InnerMsg : " + $inner.Message)
    }
  }
  if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
    $lines += ("Details  : " + $ErrorRecord.ErrorDetails.Message)
  }
  if ($ErrorRecord.CategoryInfo) {
    $lines += ("Category : " + $ErrorRecord.CategoryInfo.Category)
    $lines += ("Reason   : " + $ErrorRecord.CategoryInfo.Reason)
    if ($ErrorRecord.CategoryInfo.TargetName) { $lines += ("Target   : " + $ErrorRecord.CategoryInfo.TargetName) }
  }
  if ($ErrorRecord.FullyQualifiedErrorId) {
    $lines += ("FQID     : " + $ErrorRecord.FullyQualifiedErrorId)
  }
  # Try to include web response body if available
  $webEx = $ErrorRecord.Exception
  while ($webEx -and $webEx.InnerException) { $webEx = $webEx.InnerException }
  if ($webEx -is [System.Net.WebException] -and $webEx.Response) {
    try {
      $sr = New-Object System.IO.StreamReader($webEx.Response.GetResponseStream())
      $body = $sr.ReadToEnd()
      $sr.Close()
      if ($body) { $lines += ("Response : " + $body) }
    } catch { }
  }
  return ($lines -join [Environment]::NewLine)
}

function Convert-ToPBIP {
  param(
    [Parameter(Mandatory=$true)][string]$PBIXPath,
    [Parameter(Mandatory=$true)][string]$DesktopPath,
    [Parameter(Mandatory=$true)][string]$PBIPRoot
  )
  if (-not (Test-Path $DesktopPath)) {
    throw "PBIDesktop executable not found at: $DesktopPath"
  }
  $name = [System.IO.Path]::GetFileNameWithoutExtension($PBIXPath)
  $target = Join-Path -Path $PBIPRoot -ChildPath $name
  New-Item -ItemType Directory -Path $PBIPRoot -Force | Out-Null
  if (Test-Path $target) { Remove-Item -Recurse -Force $target }

  $pbixAbs = (Resolve-Path $PBIXPath).Path
  $pbipAbs = (Resolve-Path $PBIPRoot).Path + "\" + $name

  $variants = @(
    @("--convert", $pbixAbs, "--to-pbip", $pbipAbs, "--quit"),
    @("-convert", $pbixAbs, "-to-pbip", $pbipAbs, "-quit"),
    @("/convert", $pbixAbs, "/to-pbip", $pbipAbs, "/quit")
  )
  foreach ($argsVariant in $variants) {
    try {
      & $DesktopPath @argsVariant | Out-Null
      if (Test-Path $target) { return $true }
      Start-Sleep -Seconds 5
      if (Test-Path $target) { return $true }
    }
    catch { }
  }
  return $false
}

function Export-AllReports {
  param(
    [Parameter(Mandatory=$true)][string]$GroupId,
    [Parameter(Mandatory=$true)][string]$OutFolder
  )

  New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null
  $pbipRoot = Join-Path -Path $OutFolder -ChildPath 'pbip_files'
  if ($ConvertToPBIP) { New-Item -ItemType Directory -Path $pbipRoot -Force | Out-Null }

  $reports = Get-PowerBIReport -WorkspaceId $GroupId
  # Exclude built-in Usage Metrics report by exact name
  if ($reports) { $reports = $reports | Where-Object { $_.Name -ne 'Report Usage Metrics Report' } }
  if (-not $reports) {
    Write-Host "No reports found in workspace." -ForegroundColor Yellow
    return
  }

  $totalSw = [System.Diagnostics.Stopwatch]::StartNew()
  $results = @()

  Write-Host "Found $($reports.Count) report(s). Output: $((Resolve-Path $OutFolder).Path)" -ForegroundColor Cyan

  $i = 0
  foreach ($rep in $reports) {
    $i++
    $name = Sanitize-FileName $rep.Name
    $pbixPath = Join-Path -Path $OutFolder -ChildPath ($name + '.pbix')
    

    Write-Host "[$i/$($reports.Count)] $($rep.Name) (Connection: $connType) - exporting..." -ForegroundColor White
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ok = $true
    $errMsg = $null
    try {
      Export-PowerBIReport -Id $rep.Id -WorkspaceId $GroupId -OutFile $pbixPath -ErrorAction Stop | Out-Null
    }
    catch {
      $ok = $false
      $errMsg = $_.Exception.Message
      Write-Host "   Export failed: $errMsg" -ForegroundColor Red
    }
    $sw.Stop()
    $exportMs = $sw.ElapsedMilliseconds

    $convMs = 0
    $convOk = $false
    if ($ok -and $ConvertToPBIP) {
      Write-Host "   Converting to PBIP…" -ForegroundColor White
      $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
      try {
        $convOk = Convert-ToPBIP -PBIXPath $pbixPath -DesktopPath $PBIDesktopPath -PBIPRoot $pbipRoot
        if (-not $convOk) { $errMsg = "PBIP conversion failed" }
      }
      catch {
        $convOk = $false
        $errMsg = $_.Exception.Message
      }
      $sw2.Stop()
      $convMs = $sw2.ElapsedMilliseconds
    }

    if ($ok -and ($ConvertToPBIP -and $convOk -or -not $ConvertToPBIP)) {
      if (Test-Path $pbixPath) {
        $fi = Get-Item $pbixPath
        $sizeBytes = [int64]$fi.Length
        $sizeMB = [Math]::Round($sizeBytes / 1MB, 2)
      }
      Write-Host ("   Done. Size: {0} MB. Export: {1}ms{2}" -f $sizeMB, $exportMs, ($(if ($ConvertToPBIP) { ", Convert: $convMs`ms" } else { "" })))
      if ($ConvertToPBIP) { Write-Host ", Convert: ${convMs}ms" } else { Write-Host "" }
      $results += [PSCustomObject]@{
        Name = $rep.Name
  SizeMB = $sizeMB
  SizeBytes = $sizeBytes
        ExportMs = $exportMs
        ConvertMs = $(if ($ConvertToPBIP) { $convMs } else { 0 })
        Status = "Success"
      }
    }
    else {
      # Print detailed error information
      if ($errMsg) {
        Write-Host "   FAILED. Export: ${exportMs}ms" -ForegroundColor Red
        Write-Host ("   Error details:`n" + (Format-Error -ErrorRecord $_)) -ForegroundColor DarkRed
      } else {
        Write-Host "   FAILED. Export: ${exportMs}ms" -ForegroundColor Red
      }
      $results += [PSCustomObject]@{
        Name = $rep.Name
  SizeMB = 0
  SizeBytes = 0
        ExportMs = $exportMs
        ConvertMs = $(if ($ConvertToPBIP) { $convMs } else { 0 })
        Status = "Failed"
        Error = $errMsg
      }
    }
  }

  $totalSw.Stop()
  Write-Host ""; Write-Host "========================================" -ForegroundColor Cyan
  Write-Host "SUMMARY" -ForegroundColor Cyan
  Write-Host "========================================" -ForegroundColor Cyan
  $succ = ($results | Where-Object { $_.Status -eq 'Success' }).Count
  $fail = ($results | Where-Object { $_.Status -ne 'Success' }).Count
  Write-Host "Reports: $($results.Count)" -ForegroundColor White
  Write-Host "Success: $succ" -ForegroundColor Green
  Write-Host "Failed : $fail" -ForegroundColor Red
  Write-Host "Total time: $($totalSw.Elapsed.ToString())" -ForegroundColor White
  Write-Host ""; $results | Format-Table Name, SizeMB, ExportMs, ConvertMs, Status -AutoSize
}

try {
  Ensure-Module -Name MicrosoftPowerBIMgmt
  Ensure-Module -Name MicrosoftPowerBIMgmt.Reports
  Ensure-Module -Name MicrosoftPowerBIMgmt.Workspaces
}
catch {
  Write-Host "Failed to install/import Power BI modules: $_" -ForegroundColor Red
  exit 1
}

try {
  Connect-PowerBIServiceAccount -ErrorAction Stop | Out-Null
}
catch {
  Write-Host "Failed to connect to Power BI. $_" -ForegroundColor Red
  exit 1
}

if (-not $WorkspaceId -and -not $WorkspaceName) {
  Write-Host "Please provide -WorkspaceName or -WorkspaceId" -ForegroundColor Yellow
  exit 1
}

$ws = $null
if ($WorkspaceId) {
  $ws = Get-PowerBIWorkspace -Id $WorkspaceId -Scope Organization -ErrorAction SilentlyContinue
  if (-not $ws) { $ws = Get-PowerBIWorkspace -Id $WorkspaceId -ErrorAction SilentlyContinue }
}
elseif ($WorkspaceName) {
  $ws = Get-PowerBIWorkspace -Name $WorkspaceName -ErrorAction SilentlyContinue | Select-Object -First 1
}

if (-not $ws) {
  Write-Host "Workspace not found." -ForegroundColor Red
  exit 1
}

Write-Host "Workspace: $($ws.Name) [$($ws.Id)]" -ForegroundColor Cyan

if ($ConvertToPBIP -and -not $PBIDesktopPath) {
  Write-Host "-ConvertToPBIP specified but -PBIDesktopPath is missing." -ForegroundColor Yellow
}

Export-AllReports -GroupId $ws.Id -OutFolder $OutputFolder
