[CmdletBinding()]
param(
    [switch]$RunNow
)

$ErrorActionPreference = "Continue"
$TaskName = "Safripol Ops Report Refresh"
$Repo = Split-Path -Parent $PSScriptRoot
$Log = Join-Path $Repo "refresh.log"
$Url = "https://chrisviljoen01.github.io/Safripol-Operations-KPI-Reports/"
$VersionUrl = $Url + "data/version.json"
$SnapshotUrl = $Url + "data/snapshot.json"

function Section($Title) {
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
}

Write-Host "Safripol Operations KPI Report - Spot Check" -ForegroundColor Green
Write-Host "Report: $Url"

Section "Scheduled task"
try {
    if ($RunNow) {
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "Manual refresh triggered. Wait 30-60 seconds, then run this check again."
        Write-Host ""
    }
    $task = Get-ScheduledTask -TaskName $TaskName
    $info = $task | Get-ScheduledTaskInfo
    Write-Host ("State          : {0}" -f $task.State)
    Write-Host ("Last run       : {0}" -f $info.LastRunTime)
    Write-Host ("Last result    : {0}" -f $info.LastTaskResult)
    Write-Host ("Next run       : {0}" -f $info.NextRunTime)
    if ($info.LastTaskResult -eq 0) {
        Write-Host "Status         : OK" -ForegroundColor Green
    } else {
        Write-Host "Status         : CHECK refresh.log" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Scheduled task not found or unreadable: $($_.Exception.Message)" -ForegroundColor Red
}

Section "Published report data"
try {
    $version = (Invoke-WebRequest -Uri $VersionUrl -UseBasicParsing -TimeoutSec 20).Content | ConvertFrom-Json
    $snapshot = (Invoke-WebRequest -Uri $SnapshotUrl -UseBasicParsing -TimeoutSec 30).Content | ConvertFrom-Json
    Write-Host ("Last published : {0}" -f $version.generated_display)
    Write-Host ("Degraded       : {0}" -f $version.degraded)
    Write-Host ("Received       : {0:N2} MT" -f [double]$snapshot.headline.received_admin_mt)
    Write-Host ("Delivered      : {0:N2} MT" -f [double]$snapshot.headline.delivered_mt)
    Write-Host ("Outstanding    : {0:N2} MT" -f [double]$snapshot.headline.outstanding_mt)
    Write-Host ""
    Write-Host "Sources:"
    foreach ($source in $snapshot.meta.sources) {
        $state = if ($source.ok) { "OK" } else { "FAIL" }
        Write-Host ("  {0,-9} {1,-4} via {2,-6} {3}" -f $source.key, $state, $source.origin, $source.last_modified)
    }
} catch {
    Write-Host "Could not read published JSON: $($_.Exception.Message)" -ForegroundColor Red
}

Section "Local refresh log"
if (Test-Path $Log) {
    Get-Content $Log -Tail 14
} else {
    Write-Host "No refresh.log found yet."
}

Section "Local repository"
try {
    Push-Location $Repo
    $status = git status --short
    if ($status) {
        Write-Host "Uncommitted files:" -ForegroundColor Yellow
        $status
    } else {
        Write-Host "Clean working tree." -ForegroundColor Green
    }
} catch {
    Write-Host "Could not read git status: $($_.Exception.Message)" -ForegroundColor Yellow
} finally {
    Pop-Location
}

Write-Host ""
