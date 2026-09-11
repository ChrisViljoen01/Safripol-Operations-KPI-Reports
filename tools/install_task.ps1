<#
    Registers the Safripol report refresh as a Windows scheduled task.

    Delegated Graph access is tied to a signed-in user, so the refresh runs on
    this machine rather than on a CI runner. Run tools/../ingest --login once
    first, then run this script.

    Usage:
        powershell -ExecutionPolicy Bypass -File tools\install_task.ps1
        powershell -ExecutionPolicy Bypass -File tools\install_task.ps1 -Minutes 15
        powershell -ExecutionPolicy Bypass -File tools\install_task.ps1 -Remove
#>
[CmdletBinding()]
param(
    [int]$Minutes = 15,
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$TaskName = "Safripol Ops Report Refresh"
$Repo     = Split-Path -Parent $PSScriptRoot
$Script   = Join-Path $Repo "tools\refresh_and_publish.py"

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Yellow
    return
}

$python = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $python) { $python = (Get-Command python.exe).Source }
if (-not (Test-Path $Script)) { throw "Cannot find $Script" }

# The token cache lives in the user profile, so the task must run as this user.
# Interactive-only (no stored password) keeps it running whenever Chris is
# logged in, which matches the operating hours the report is used in.
$action = New-ScheduledTaskAction -Execute $python `
    -Argument "`"$Script`"" -WorkingDirectory $Repo

$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) `
    -RepetitionInterval (New-TimeSpan -Minutes $Minutes)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME `
    -LogonType Interactive -RunLevel Limited

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal `
    -Description "Rebuilds the Safripol TAC IMOLA ops report from SharePoint and publishes it to GitHub Pages." | Out-Null

Write-Host "Registered '$TaskName' - every $Minutes minutes." -ForegroundColor Green
Write-Host "  Log:  $Repo\refresh.log"
Write-Host "  Run now:  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "  Remove :  powershell -File tools\install_task.ps1 -Remove"
