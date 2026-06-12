#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Proactive Remediation - DETECTION script.
    Triggers remediation when:
      - No backup has ever been completed, OR
      - The OneNote cache has been modified since the last backup
        AND the minimum backup interval has elapsed.

EXIT CODES (Intune convention):
    0 = Compliant   - backup is up to date, no action needed.
    1 = Non-compliant - no backup, or cache changed since last backup.

RECOMMENDED INTUNE SCHEDULE: Every 1 hour.
The $MinBackupIntervalHours guard prevents redundant re-backups.
#>

# ===========================================================================
# CONFIGURATION
# ===========================================================================
$OutputRoot            = "C:\ProgramData\GFD-MIG\OneNoteMigration"
$MinBackupIntervalHours = 4    # Never re-backup faster than this, even if cache changed
$OneNoteCachePath       = Join-Path $env:LOCALAPPDATA "Microsoft\OneNote\16.0"
# ===========================================================================

$UserName     = $env:USERNAME
$ComputerName = $env:COMPUTERNAME
$InventoryDir = Join-Path $OutputRoot "Inventory"

# ---------------------------------------------------------------------------
# Find the most recent valid backup for this user on this machine
# ---------------------------------------------------------------------------
function Get-LastBackup {
    if (-not (Test-Path $InventoryDir)) { return $null }

    $markers = Get-ChildItem -Path $InventoryDir -Filter "PreCompleted.json" -Recurse -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -match [regex]::Escape($ComputerName) } |
               Sort-Object LastWriteTime -Descending

    foreach ($file in $markers) {
        try {
            $data = Get-Content -Path $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($data.Completed -eq $true -and $data.UserName -eq $UserName) {
                return [PSCustomObject]@{
                    CompletedAt = [datetime]$data.CompletedAt
                    OutputPath  = $data.OutputPath
                    MarkerFile  = $file.FullName
                }
            }
        } catch { }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Get the newest LastWriteTime of any file in the OneNote cache folder
# ---------------------------------------------------------------------------
function Get-CacheLastModified {
    if (-not (Test-Path $OneNoteCachePath)) { return $null }
    try {
        $newest = Get-ChildItem -Path $OneNoteCachePath -Recurse -File -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending |
                  Select-Object -First 1
        return if ($newest) { $newest.LastWriteTime } else { $null }
    } catch { return $null }
}

# ---------------------------------------------------------------------------
# Decision logic
# ---------------------------------------------------------------------------
$lastBackup = Get-LastBackup

if (-not $lastBackup) {
    Write-Host "Non-compliant: no backup found for $UserName on $ComputerName."
    exit 1
}

$now             = Get-Date
$hoursSinceBackup = ($now - $lastBackup.CompletedAt).TotalHours

# Enforce minimum interval - avoid hammering if Intune runs frequently
if ($hoursSinceBackup -lt $MinBackupIntervalHours) {
    Write-Host ("Compliant: last backup {0:F1}h ago (min interval {1}h). Too soon to re-check." -f $hoursSinceBackup, $MinBackupIntervalHours)
    exit 0
}

$cacheModified = Get-CacheLastModified

if (-not $cacheModified) {
    Write-Host "Compliant: OneNote cache folder not found or empty - no changes to back up."
    exit 0
}

if ($cacheModified -gt $lastBackup.CompletedAt) {
    Write-Host ("Non-compliant: cache last modified {0} - newer than last backup {1}. Re-backup needed." -f $cacheModified, $lastBackup.CompletedAt)
    exit 1
}

Write-Host ("Compliant: cache last modified {0} - no changes since last backup {1}." -f $cacheModified, $lastBackup.CompletedAt)
exit 0
