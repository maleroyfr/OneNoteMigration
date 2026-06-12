#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Proactive Remediation — DETECTION script.
    Checks whether Pre-OneNote-InventoryBackup.ps1 has already completed
    for the current user on this machine.

EXIT CODES (Intune convention):
    0  = Compliant   — inventory already completed, no remediation needed.
    1  = Non-compliant — inventory not yet completed, trigger remediation.
#>

$OutputRoot   = "C:\ProgramData\GFD-MIG\OneNoteMigration"
$InventoryDir = Join-Path $OutputRoot "Inventory"

function Test-InventoryCompleted {
    if (-not (Test-Path $InventoryDir)) { return $false }

    # Look for any PreCompleted.json written by the current user on this machine
    $userName     = $env:USERNAME
    $computerName = $env:COMPUTERNAME

    $markers = Get-ChildItem -Path $InventoryDir -Filter "PreCompleted.json" -Recurse -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -match [regex]::Escape($computerName) }

    foreach ($file in $markers) {
        try {
            $data = Get-Content -Path $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($data.Completed -eq $true -and $data.UserName -eq $userName) {
                Write-Host "Compliant: found marker $($file.FullName) (completed at $($data.CompletedAt))"
                return $true
            }
        } catch { }
    }
    return $false
}

if (Test-InventoryCompleted) {
    exit 0   # Compliant
} else {
    Write-Host "Non-compliant: no valid PreCompleted.json found for $env:USERNAME on $env:COMPUTERNAME"
    exit 1   # Non-compliant → triggers remediation
}
