#Requires -Version 5.1
<#
.SYNOPSIS
    Post-migration OneNote cutover script.
#>
[CmdletBinding()]
param(
    [string]$OutputRoot = "C:\ProgramData\GFD-MIG\OneNoteMigration",
    [string]$SourceTenantHost = "sesvanderhave.sharepoint.com",
    [string]$MappingCsvPath,
    [switch]$IncludeStoreAppCache,
    [switch]$OpenTargetNotebooks,
    [switch]$UseOneNoteComOpenHierarchy,
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptName = "Post-OneNote-Cutover.ps1"
$ScriptVersion = "1.0.0"

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$UserName = $env:USERNAME
$UserDomain = $env:USERDOMAIN
$ComputerName = $env:COMPUTERNAME
$UserPrincipalName = ""
try {
    $upnClaim = $CurrentUser.Claims |
        Where-Object { $_.Type -eq "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn" } |
        Select-Object -First 1
    if ($upnClaim) { $UserPrincipalName = $upnClaim.Value.Trim() }
} catch { }
if (-not $UserPrincipalName) {
    try {
        $adsiUser = [ADSI]"LDAP://<SID=$($CurrentUser.User.Value)>"
        if ($adsiUser.userPrincipalName) { $UserPrincipalName = [string]$adsiUser.userPrincipalName.Trim() }
    } catch { }
}

$RunId = if ($UserPrincipalName) {
    ($UserPrincipalName.Trim() -replace '[\\/:*?"<>|@]', '_') + "_${ComputerName}_${Timestamp}"
} else {
    "${UserDomain}_${UserName}_${ComputerName}_${Timestamp}"
}

$PostMigrationPath = Join-Path $OutputRoot "PostMigration\$RunId"
$QuarantinePath = Join-Path $OutputRoot "Quarantine\$RunId"
$LogsPath = Join-Path $OutputRoot "Logs"
$LogFile = Join-Path $LogsPath "${RunId}.log"
$SummaryFile = Join-Path $PostMigrationPath "PostSummary.json"
$CompletedFile = Join-Path $PostMigrationPath "PostCompleted.json"
$OpenedTargetsFile = Join-Path $PostMigrationPath "OpenedTargets.csv"
$SkippedTargetsFile = Join-Path $PostMigrationPath "SkippedTargets.csv"
$UserActionRequiredFile = Join-Path $PostMigrationPath "UserActionRequired.txt"
$SourceNotebooksFile = Join-Path $PostMigrationPath "SourceNotebooks.csv"
$EmergencyHierarchyFile = Join-Path $PostMigrationPath "EmergencyHierarchy.xml"
$OpenedTargets = New-Object System.Collections.Generic.List[object]
$SkippedTargets = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARNING","ERROR")][string]$Severity = "INFO"
    )
    $entry = "{0}  [{1}]  {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Severity.PadRight(7), $Message
    try { Add-Content -Path $LogFile -Value $entry -Encoding UTF8 } catch { }
    switch ($Severity) {
        "WARNING" { Write-Warning $Message }
        "ERROR" { Write-Host "[ERROR] $Message" -ForegroundColor Red }
        default { Write-Host $entry }
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-UniquePath {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $Path }
    $base = $Path
    $suffix = 1
    while (Test-Path $Path) {
        $Path = "{0}_{1}" -f $base, $suffix
        $suffix++
    }
    return $Path
}

function Invoke-RobocopyBackup {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Test-Path $Source)) {
        Write-Log ("Source not found for {0}: {1}" -f $Label, $Source) -Severity WARNING
        return $false
    }
    if ($WhatIf) {
        Write-Log "[WhatIf] Would robocopy $Label from '$Source' to '$Destination'" -Severity WARNING
        return $true
    }
    Ensure-Directory -Path $Destination
    $args = @(
        ('"{0}"' -f $Source),
        ('"{0}"' -f $Destination),
        "/E",
        "/R:1",
        "/W:1",
        "/XJ",
        "/NP",
        ('/LOG+:"{0}"' -f $LogFile)
    )
    try {
        $proc = Start-Process -FilePath "robocopy.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow
        Write-Log "Robocopy [$Label] exit code: $($proc.ExitCode)"
        return ($proc.ExitCode -le 7)
    } catch {
        Write-Log "Robocopy [$Label] failed: $_" -Severity ERROR
        return $false
    }
}

function Stop-OneNoteProcesses {
    if ($WhatIf) {
        Write-Log "[WhatIf] Would stop OneNote processes." -Severity WARNING
        return
    }
    foreach ($name in @("ONENOTE","ONENOTEM")) {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
        foreach ($proc in @($procs)) {
            try {
                Write-Log "Attempting graceful close for $name PID $($proc.Id)"
                $null = $proc.CloseMainWindow()
                $null = $proc.WaitForExit(10000)
            } catch {
                Write-Log "Graceful close failed for PID $($proc.Id): $_" -Severity WARNING
            }
            if (-not $proc.HasExited) {
                try {
                    Write-Log "Force stopping $name PID $($proc.Id)" -Severity WARNING
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                } catch {
                    Write-Log "Failed to stop PID $($proc.Id): $_" -Severity ERROR
                }
            }
        }
    }
}

function Get-LatestPreCompleted {
    $inventoryRoot = Join-Path $OutputRoot "Inventory"
    if (-not (Test-Path $inventoryRoot)) { return $null }
    Get-ChildItem -Path $inventoryRoot -Filter "PreCompleted.json" -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Get-LatestPreInventoryCsv {
    $inventoryRoot = Join-Path $OutputRoot "Inventory"
    if (-not (Test-Path $inventoryRoot)) { return $null }
    Get-ChildItem -Path $inventoryRoot -Filter "OneNote-Inventory.csv" -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Capture-EmergencyHierarchy {
    try {
        $app = New-Object -ComObject OneNote.Application
        [string]$hierarchyXml = ""
        $app.GetHierarchy("", 3, [ref]$hierarchyXml)
        if ($hierarchyXml) {
            $hierarchyXml | Set-Content -Path $EmergencyHierarchyFile -Encoding UTF8
            return $true
        }
    } catch {
        Write-Log "Emergency hierarchy capture failed: $_" -Severity WARNING
    }
    return $false
}

function Quarantine-Path {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Test-Path $Source)) {
        Write-Log "$Label source not found: $Source" -Severity WARNING
        return $false
    }

    $Destination = Get-UniquePath -Path $Destination
    if ($WhatIf) {
        Write-Log "[WhatIf] Would quarantine $Label from '$Source' to '$Destination'" -Severity WARNING
        return $true
    }

    try {
        Move-Item -Path $Source -Destination $Destination -ErrorAction Stop
        Write-Log "Moved $Label to quarantine: $Destination"
        return $true
    } catch {
        Write-Log "Move failed for $Label, falling back to robocopy copy: $_" -Severity WARNING
        $copyOk = Invoke-RobocopyBackup -Source $Source -Destination $Destination -Label $Label
        if ($copyOk) {
            Write-Log "$Label copied to quarantine because move failed." -Severity WARNING
        }
        return $copyOk
    }
}

function Open-TargetUrl {
    param([Parameter(Mandatory)][string]$TargetUrl)

    if ($UseOneNoteComOpenHierarchy) {
        try {
            $app = New-Object -ComObject OneNote.Application
            [string]$objectId = ""
            $app.OpenHierarchy($TargetUrl, "", [ref]$objectId)
            return $true
        } catch {
            Write-Log "COM open failed for '$TargetUrl', falling back to Start-Process: $_" -Severity WARNING
        }
    }

    try {
        Start-Process -FilePath $TargetUrl | Out-Null
        return $true
    } catch {
        Write-Log "Failed to open '$TargetUrl': $_" -Severity ERROR
        return $false
    }
}

try {
    Ensure-Directory -Path $PostMigrationPath
    Ensure-Directory -Path $QuarantinePath
    Ensure-Directory -Path $LogsPath
} catch {
    Write-Host "[ERROR] Unable to create required output folders: $_" -ForegroundColor Red
    exit 1
}

Write-Log "===== $ScriptName v$ScriptVersion - Start ====="
Write-Log "User: $UserDomain\$UserName $(if ($UserPrincipalName) { "($UserPrincipalName)" })"
Write-Log "Computer: $ComputerName"
Write-Log "OutputRoot: $OutputRoot"
Write-Log "PostMigrationPath: $PostMigrationPath"
Write-Log "QuarantinePath: $QuarantinePath"
Write-Log "WhatIf: $WhatIf"

$preCompletedFound = $false
$preCompletedPath = ""
$desktopCacheFound = $false
$desktopBackupOk = $false
$desktopQuarantineOk = $false
$storeCacheFound = $false
$storeBackupOk = $false
$storeQuarantineOk = $false
$hierarchyExported = $false
$targetsOpened = 0
$targetsSkipped = 0
$exitCode = 0

$latestPre = Get-LatestPreCompleted
if ($latestPre) {
    $preCompletedFound = $true
    $preCompletedPath = $latestPre.FullName
    Write-Log "Latest pre-migration completion found: $preCompletedPath"
} else {
    Write-Log "No pre-migration completion marker found." -Severity WARNING
}

$latestInventory = Get-LatestPreInventoryCsv
if ($latestInventory) {
    Write-Log "Latest pre-migration inventory found: $($latestInventory.FullName)"
}

Stop-OneNoteProcesses

if (Capture-EmergencyHierarchy) {
    $hierarchyExported = $true
    Write-Log "Emergency hierarchy exported to $EmergencyHierarchyFile"
}

$desktopSource = Join-Path $env:LOCALAPPDATA "Microsoft\OneNote\16.0"
$desktopBackupDest = Join-Path $PostMigrationPath "EmergencyBackup\DesktopOneNote-16.0"
$desktopQuarantineDest = Join-Path $QuarantinePath "DesktopOneNote-16.0"
if (Test-Path $desktopSource) {
    $desktopCacheFound = $true
    $desktopBackupOk = Invoke-RobocopyBackup -Source $desktopSource -Destination $desktopBackupDest -Label "DesktopOneNote16EmergencyBackup"
    $desktopQuarantineOk = Quarantine-Path -Source $desktopSource -Destination $desktopQuarantineDest -Label "DesktopOneNote-16.0"
}

if ($IncludeStoreAppCache) {
    $storeSource = Join-Path $env:LOCALAPPDATA "Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalCache"
    $storeBackupDest = Join-Path $PostMigrationPath "EmergencyBackup\StoreApp-LocalCache"
    $storeQuarantineDest = Join-Path $QuarantinePath "StoreApp-LocalCache"
    if (Test-Path $storeSource) {
        $storeCacheFound = $true
        $storeBackupOk = Invoke-RobocopyBackup -Source $storeSource -Destination $storeBackupDest -Label "StoreAppEmergencyBackup"
        $storeQuarantineOk = Quarantine-Path -Source $storeSource -Destination $storeQuarantineDest -Label "StoreApp-LocalCache"
    } else {
        Write-Log "Store app cache not found." -Severity WARNING
    }
}

if ($OpenTargetNotebooks) {
    if (-not $MappingCsvPath) {
        Write-Log "-OpenTargetNotebooks requested without -MappingCsvPath. Skipping notebook reopen." -Severity WARNING
        if (-not $WhatIf) {
            @(
                "No mapping CSV was provided.",
                "The old OneNote cache/config was quarantined successfully.",
                "Support/user action required: reopen the needed notebooks from the target Teams or SharePoint location.",
                ""
                "If the target notebooks are already known, rerun this script with -MappingCsvPath to reopen them automatically."
            ) | Set-Content -Path $UserActionRequiredFile -Encoding UTF8
        }
    } elseif (-not (Test-Path $MappingCsvPath)) {
        Write-Log "Mapping CSV not found: $MappingCsvPath. Skipping notebook reopen." -Severity WARNING
        if (-not $WhatIf) {
            @(
                "The provided mapping CSV could not be found.",
                "The old OneNote cache/config was quarantined successfully.",
                "Support/user action required: reopen the needed notebooks from the target Teams or SharePoint location.",
                ""
                "Provide a validated mapping CSV and rerun if automatic reopening is required."
            ) | Set-Content -Path $UserActionRequiredFile -Encoding UTF8
        }
    } else {
        try {
            $mappingRows = Import-Csv -Path $MappingCsvPath
            $openRows = $mappingRows | Where-Object {
                $_.Status -eq "Validated" -and $_.Action -eq "OpenTarget" -and $_.TargetOpenUrl
            } | Sort-Object TargetOpenUrl -Unique

            foreach ($row in $mappingRows | Where-Object {
                -not ($_.Status -eq "Validated" -and $_.Action -eq "OpenTarget" -and $_.TargetOpenUrl)
            }) {
                $SkippedTargets.Add([PSCustomObject]@{
                    TargetOpenUrl = $row.TargetOpenUrl
                    Status = $row.Status
                    Action = $row.Action
                    Comment = $row.Comment
                }) | Out-Null
            }

            foreach ($row in $openRows) {
                $ok = Open-TargetUrl -TargetUrl $row.TargetOpenUrl
                if ($ok) {
                    $targetsOpened++
                    $OpenedTargets.Add([PSCustomObject]@{
                        TargetOpenUrl = $row.TargetOpenUrl
                        Status = $row.Status
                        Action = $row.Action
                        Comment = $row.Comment
                    }) | Out-Null
                } else {
                    $targetsSkipped++
                    $SkippedTargets.Add([PSCustomObject]@{
                        TargetOpenUrl = $row.TargetOpenUrl
                        Status = $row.Status
                        Action = $row.Action
                        Comment = $row.Comment
                    }) | Out-Null
                }
            }
        } catch {
            Write-Log "Failed to process mapping CSV: $_" -Severity ERROR
            $exitCode = 1
        }
    }
} elseif (-not $MappingCsvPath) {
    Write-Log "No mapping CSV provided. Skipping notebook reopen by design." -Severity WARNING
    if (-not $WhatIf) {
        @(
            "No mapping CSV was provided.",
            "The old OneNote cache/config was quarantined successfully.",
            "Support/user action required: reopen the needed notebooks from the target Teams or SharePoint location.",
            ""
            "If automatic reopening is needed later, rerun this script with a validated -MappingCsvPath."
        ) | Set-Content -Path $UserActionRequiredFile -Encoding UTF8
    }
}

try {
    if ($OpenedTargets.Count -gt 0) { $OpenedTargets | Export-Csv -Path $OpenedTargetsFile -NoTypeInformation -Encoding UTF8 } else { "" | Set-Content -Path $OpenedTargetsFile -Encoding UTF8 }
    if ($SkippedTargets.Count -gt 0) { $SkippedTargets | Export-Csv -Path $SkippedTargetsFile -NoTypeInformation -Encoding UTF8 } else { "" | Set-Content -Path $SkippedTargetsFile -Encoding UTF8 }
} catch {
    Write-Log "Failed to export target CSV files: $_" -Severity WARNING
}

if ($latestInventory -and (Test-Path $latestInventory.FullName)) {
    try {
        $inventoryRows = Import-Csv -Path $latestInventory.FullName
        $sourceRows = $inventoryRows | Where-Object {
            $_.IsSourceTenant -eq "True" -or $_.IsSourceTenant -eq $true
        }
        if ($sourceRows) {
            $sourceRows | Export-Csv -Path $SourceNotebooksFile -NoTypeInformation -Encoding UTF8
            Write-Log "Exported source notebook copy from latest inventory to $SourceNotebooksFile"
        } else {
            "" | Set-Content -Path $SourceNotebooksFile -Encoding UTF8
            Write-Log "Latest inventory had no source-tenant rows to export." -Severity WARNING
        }
    } catch {
        Write-Log "Failed to export source notebook copy from inventory: $_" -Severity WARNING
    }
} else {
    Write-Log "No pre-migration inventory available to export source notebooks from." -Severity WARNING
}

$summary = [ordered]@{
    UserName = $UserName
    UserDomain = $UserDomain
    UserPrincipalName = $UserPrincipalName
    ComputerName = $ComputerName
    Timestamp = $Timestamp
    SourceTenantHost = $SourceTenantHost
    PreCompletedFound = $preCompletedFound
    PreCompletedPath = $preCompletedPath
    EmergencyHierarchyExported = $hierarchyExported
    DesktopCacheFound = $desktopCacheFound
    DesktopEmergencyBackupSucceeded = $desktopBackupOk
    DesktopQuarantineSucceeded = $desktopQuarantineOk
    StoreAppCacheFound = $storeCacheFound
    StoreAppEmergencyBackupSucceeded = $storeBackupOk
    StoreAppQuarantineSucceeded = $storeQuarantineOk
    MappingCsvPath = $MappingCsvPath
    OpenTargetNotebooks = $OpenTargetNotebooks.IsPresent
    TargetsOpened = $targetsOpened
    TargetsSkipped = $targetsSkipped
    PostMigrationPath = $PostMigrationPath
    QuarantinePath = $QuarantinePath
    ExitCode = $exitCode
}

try {
    $summary | ConvertTo-Json -Depth 5 | Set-Content -Path $SummaryFile -Encoding UTF8
} catch {
    Write-Log "Failed to write summary JSON: $_" -Severity ERROR
    $exitCode = 1
}

try {
    [ordered]@{
        Completed = $true
        CompletedAt = (Get-Date -Format "o")
        ScriptName = $ScriptName
        ScriptVersion = $ScriptVersion
        UserName = $UserName
        ComputerName = $ComputerName
        PostMigrationPath = $PostMigrationPath
        QuarantinePath = $QuarantinePath
    } | ConvertTo-Json -Depth 3 | Set-Content -Path $CompletedFile -Encoding UTF8
} catch {
    Write-Log "Failed to write completion JSON: $_" -Severity ERROR
    $exitCode = 1
}

if (-not $desktopCacheFound -and -not $storeCacheFound -and -not $preCompletedFound -and -not $hierarchyExported) {
    Write-Log "No post-cutover action was needed or possible, but script completed safely." -Severity WARNING
}

Write-Log "===== $ScriptName - End (exit $exitCode) ====="
exit $exitCode
