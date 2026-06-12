#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Proactive Remediation — REMEDIATION script.
    Invokes Pre-OneNote-InventoryBackup.ps1 with parameters suitable
    for a migration preparation run.

DEPLOYMENT REQUIREMENT:
    Pre-OneNote-InventoryBackup.ps1 must be deployed to the path defined
    in $MainScriptPath before this remediation script runs.
    Recommended: deploy it as an Intune Win32 app (or script) to
    C:\ProgramData\GFD-MIG\Scripts\Pre-OneNote-InventoryBackup.ps1

EXIT CODES (Intune convention):
    0  = Remediation succeeded.
    1  = Remediation failed.
#>

# ---------------------------------------------------------------------------
# Configuration — adjust to match your environment
# ---------------------------------------------------------------------------
$MainScriptPath   = "C:\ProgramData\GFD-MIG\Scripts\Pre-OneNote-InventoryBackup.ps1"
$SourceTenantHost = "sesvanderhave.sharepoint.com"
$OutputRoot       = "C:\ProgramData\GFD-MIG\OneNoteMigration"

# Stop OneNote during remediation so the cache is consistent for backup.
# Set to $false if OneNote must stay open during the inventory window.
$ShouldStopOneNote = $true

# Back up the Store app cache in addition to the desktop app cache.
$IncludeStoreAppCache = $false
# ---------------------------------------------------------------------------

function Write-Status {
    param([string]$Message, [string]$Severity = "INFO")
    $entry = "{0}  [{1}]  {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Severity.PadRight(7), $Message
    Write-Host $entry
}

# Validate main script is present
if (-not (Test-Path $MainScriptPath)) {
    Write-Status "Main script not found: $MainScriptPath" "ERROR"
    Write-Status "Deploy Pre-OneNote-InventoryBackup.ps1 to $MainScriptPath before running remediation." "ERROR"
    exit 1
}

Write-Status "Starting OneNote pre-migration inventory via: $MainScriptPath"

# Build argument list
$scriptArgs = @(
    "-SourceTenantHost", $SourceTenantHost,
    "-OutputRoot",       $OutputRoot
)
if ($ShouldStopOneNote)    { $scriptArgs += "-StopOneNote" }
if ($IncludeStoreAppCache) { $scriptArgs += "-IncludeStoreAppCache" }

try {
    # Run in the current (user) PowerShell session so COM access works
    $result = & $MainScriptPath @scriptArgs
    $exitCode = $LASTEXITCODE

    # LASTEXITCODE is set by the exit statement in the main script
    if ($exitCode -eq 0) {
        Write-Status "Remediation completed successfully (exit $exitCode)."
        exit 0
    } else {
        Write-Status "Remediation script returned non-zero exit code: $exitCode" "ERROR"
        exit 1
    }
} catch {
    Write-Status "Unhandled exception during remediation: $_" "ERROR"
    exit 1
}
