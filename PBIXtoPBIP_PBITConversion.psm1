<#

     .EXAMPLE
     PBIXtoPBIP_PBITConversion -PBIXFilePath "<<PBIXFilePath>>" -ConversionFileType "<<ConversionFileType>>"
      

     Import-Module "C:\Users\VaishnavKamartiMAQSo\Desktop\VS code explorations\DeDeuplication\Download_convert_reports\test.psm1" -Force
    Get-Module test
    PBIXtoPBIP_PBITConversion

#>
Function PBIXtoPBIP_PBITConversion
{

 

    Param 
    (

    [Parameter(Mandatory=$true)]
    [string] $PBIXFilePath,
    [Parameter(Mandatory=$true)]
    [string] $ConversionFileType,
    [Parameter(Mandatory=$false)]
    [string] $OutputFolder = "C:\\Users\\VaishnavKamartiMAQSo\\Desktop\\VS code explorations\\DeDeuplication\\Download_convert_reports\\downloads\\pbip_files"

    )

$ConversionFileTypeUpper =$ConversionFileType.ToUpper()

# Hardcode destination folder to repo .\pbip_files regardless of input
$moduleRoot = $PSScriptRoot; if (-not $moduleRoot) { $moduleRoot = (Get-Location).Path }
$OutputFolder = Join-Path -Path $moduleRoot -ChildPath 'pbip_files'

if (($ConversionFileTypeUpper -eq "PBIP") -or ($ConversionFileTypeUpper -eq "PBIT"))
{
}
else
{
Write-Host "Incorrect parameter value passed. ConversionFileType should only have value as either PBIP or PBIT"
}

# Ensure required assemblies are available
Add-Type -AssemblyName System.Windows.Forms | Out-Null

# Open the PBIX file / Power BI report and capture the process
$proc = Start-Process -FilePath $PBIXFilePath -PassThru
Start-Sleep -Seconds 2

# Wait for the main window to appear (up to 5 minutes)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$timeoutSec = 300
$windowReady = $false
while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
    $p = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
    if ($p -and $p.MainWindowTitle) { $windowReady = $true; break }
    Write-Host "Waiting for Power BI window..."
    Start-Sleep -Seconds 2
}
if (-not $windowReady) { Write-Host "Window not ready after timeout; continuing anyway." -ForegroundColor Yellow }

# Try to bring PBIDesktop to the foreground
try {
    $wshell = New-Object -ComObject WScript.Shell
    $null = $wshell.AppActivate($proc.Id)
    Start-Sleep -Milliseconds 500
} catch { }

# Wait for CPU activity to settle to avoid acting too early
$stableCount = 0
for ($i=0; $i -lt 30 -and $stableCount -lt 2; $i++) {
    $p = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
    if (-not $p) { break }
    $cpu1 = [double]$p.CPU
    Start-Sleep -Seconds 3
    try { $p = Get-Process -Id $proc.Id -ErrorAction Stop } catch { break }
    $cpu2 = [double]$p.CPU
    $delta = [math]::Abs($cpu2 - $cpu1)
    if ($delta -lt 0.05) { $stableCount++ } else { $stableCount = 0 }
    Write-Host ("Report is still loading... CPU delta: {0:N2}" -f $delta)
}

#Close the window in case if sign in/authentication window pops up
$process = Get-Process | Where-Object{$_.ProcessName -eq "Microsoft.AAD.BrokerPlugin"}
if ($process -ne $null)
{
Write-Host "Terminating Authentication Window"
Stop-Process -Name "Microsoft.AAD.BrokerPlugin"
}
else
{
}
write-host "Starting with conversion.."
Start-Sleep -Seconds 5
[System.Windows.Forms.SendKeys]::SendWait('%')
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait('F')
 
Start-Sleep -Seconds 1
#Click File on Power BI desktop
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
#Click Save AS
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 2

if ($ConversionFileTypeUpper -eq "PBIP")
{
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')

}
elseif ($ConversionFileTypeUpper -eq "PBIT")
{
    [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
}

Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 1

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 1

# Ensure destination exists
try { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null } catch { }

#Add path
[System.Windows.Forms.SendKeys]::SendWait($OutputFolder)
#click enter to confirm path
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 1

#Redirect to save button
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 1

#finish saving the file
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 2


Start-Sleep -Seconds 30
Stop-Process -Name "PBIDesktop" # Close the Power BI Report
Start-Sleep -Seconds 2

} 

Export-ModuleMember -Function PBIXtoPBIP_PBITConversion
