#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Proactive Remediation — REMEDIATION script.
    Self-contained: no external script dependency required.
    Inventories opened OneNote notebooks and backs up local cache.

EXIT CODES (Intune convention):
    0  = Remediation succeeded (inventory or cache backup completed).
    1  = Remediation failed    (neither inventory nor cache backup completed).
#>

# ===========================================================================
# CONFIGURATION — adjust before uploading to Intune
# ===========================================================================
$SourceTenantHost     = "sesvanderhave.sharepoint.com"
$OutputRoot           = "C:\ProgramData\GFD-MIG\OneNoteMigration"
$ShouldStopOneNote    = $true   # Recommended: $true for migration prep
$IncludeStoreAppCache = $true  # Set $true if Store version of OneNote is used
# ===========================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptName    = "Remediate-OneNotePreMigration.ps1"
$ScriptVersion = "1.0.0"

# ---------------------------------------------------------------------------
# Identity & timestamp
# ---------------------------------------------------------------------------
$Timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$CurrentUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$UserName     = $env:USERNAME
$UserDomain   = $env:USERDOMAIN
$ComputerName = $env:COMPUTERNAME

$UserPrincipalName = ""
try {
    $upnClaim = $CurrentUser.Claims |
                Where-Object { $_.Type -eq "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn" } |
                Select-Object -First 1
    if ($upnClaim) { $UserPrincipalName = $upnClaim.Value }
} catch { }
if (-not $UserPrincipalName) {
    try {
        $adsiUser = [ADSI]"LDAP://<SID=$($CurrentUser.User.Value)>"
        if ($adsiUser.userPrincipalName) { $UserPrincipalName = $adsiUser.userPrincipalName }
    } catch { }
}

$RunId = if ($UserPrincipalName) {
    ($UserPrincipalName -replace '[\\/:*?"<>|@]', '_') + "_${ComputerName}_${Timestamp}"
} else {
    "${UserDomain}_${UserName}_${ComputerName}_${Timestamp}"
}

# ---------------------------------------------------------------------------
# Output paths
# ---------------------------------------------------------------------------
$InventoryPath    = Join-Path $OutputRoot "Inventory\$RunId"
$CacheBackupPath  = Join-Path $OutputRoot "CacheBackup\$RunId"
$LogDir           = Join-Path $OutputRoot "Logs"
$LogFile          = Join-Path $LogDir "${RunId}.log"
$XmlFile          = Join-Path $InventoryPath "OneNote-Hierarchy.xml"
$CsvFile          = Join-Path $InventoryPath "OneNote-Inventory.csv"
$SummaryFile      = Join-Path $InventoryPath "Summary.json"
$PreCompletedFile = Join-Path $InventoryPath "PreCompleted.json"
$InventoryLogFile = Join-Path $InventoryPath "OneNote-PreMigration.log"

foreach ($dir in @($InventoryPath, $CacheBackupPath, $LogDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARNING","ERROR")][string]$Severity = "INFO"
    )
    $entry = "{0}  [{1}]  {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Severity.PadRight(7), $Message
    try { Add-Content -Path $LogFile          -Value $entry -Encoding UTF8 } catch { }
    try { Add-Content -Path $InventoryLogFile -Value $entry -Encoding UTF8 } catch { }
    switch ($Severity) {
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Host "[ERROR] $Message" -ForegroundColor Red }
        default   { Write-Host $entry }
    }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Decode-Url {
    param([string]$Url)
    try { return [System.Uri]::UnescapeDataString($Url) } catch { return $Url }
}

function Parse-SharePointUrl {
    param([string]$Url)
    $result = [PSCustomObject]@{ Host = ""; SiteUrl = ""; RelativePath = "" }
    if (-not $Url) { return $result }
    try {
        $decoded = Decode-Url $Url
        $uri = [System.Uri]$decoded
        $result.Host = $uri.Host
        if ($decoded -match "(?i)(https?://[^/]+(?:/(?:sites|teams)/[^/]+))(/.*)?") {
            $result.SiteUrl      = $Matches[1]
            $result.RelativePath = if ($Matches[2]) { $Matches[2] } else { "" }
        } else {
            $result.SiteUrl      = "$($uri.Scheme)://$($uri.Host)"
            $result.RelativePath = $uri.AbsolutePath
        }
    } catch { }
    return $result
}

# ---------------------------------------------------------------------------
# Stop-OneNoteProcesses
# ---------------------------------------------------------------------------
function Stop-OneNoteProcesses {
    foreach ($name in @("ONENOTE", "ONENOTEM")) {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
        if (-not $procs) { Write-Log "No running process found for '$name'."; continue }
        foreach ($proc in $procs) {
            Write-Log "Found process '$name' (PID $($proc.Id)). Attempting graceful close."
            try { $proc.CloseMainWindow() | Out-Null; $proc.WaitForExit(10000) | Out-Null }
            catch { Write-Log "CloseMainWindow failed for PID $($proc.Id): $_" -Severity WARNING }
            if (-not $proc.HasExited) {
                Write-Log "Process '$name' (PID $($proc.Id)) did not exit gracefully. Forcing stop." -Severity WARNING
                try { Stop-Process -Id $proc.Id -Force -ErrorAction Stop; Write-Log "Forced stop of '$name' (PID $($proc.Id)) succeeded." }
                catch { Write-Log "Failed to force-stop '$name' (PID $($proc.Id)): $_" -Severity ERROR }
            } else {
                Write-Log "Process '$name' (PID $($proc.Id)) exited gracefully."
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Get-OneNoteInventoryRows
# ---------------------------------------------------------------------------
function Get-OneNoteInventoryRows {
    param([string]$HierarchyXml, [string]$SourceHost)

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
    try { [xml]$xml = $HierarchyXml }
    catch { Write-Log "Failed to parse hierarchy XML: $_" -Severity ERROR; return $rows }

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $rootNs = $xml.DocumentElement.NamespaceURI
    if (-not $rootNs) { $rootNs = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
    $ns.AddNamespace("one", $rootNs)

    $notebooks = $xml.SelectNodes("//one:Notebook", $ns)
    if (-not $notebooks -or $notebooks.Count -eq 0) {
        $notebooks = $xml.DocumentElement.SelectNodes("//Notebook")
    }

    foreach ($nb in $notebooks) {
        $nbName = $nb.GetAttribute("name")
        $nbId   = $nb.GetAttribute("ID")
        $nbPath = $nb.GetAttribute("path")

        $sections = $nb.SelectNodes(".//one:Section", $ns)
        if (-not $sections -or $sections.Count -eq 0) { $sections = $nb.SelectNodes(".//Section") }

        if ($sections.Count -eq 0) {
            $urlInfo  = Parse-SharePointUrl $nbPath
            $isSource = ($nbPath -like "*$SourceHost*")
            $rows.Add([PSCustomObject]@{
                UserName = $UserName; UserDomain = $UserDomain; UserPrincipalName = $UserPrincipalName
                ComputerName = $ComputerName; NotebookName = $nbName; NotebookId = $nbId; NotebookPath = $nbPath
                SectionName = ""; SectionId = ""; SectionPath = ""
                LastModifiedTime = $nb.GetAttribute("lastModifiedTime")
                IsSourceTenant = $isSource; SourceTenantHost = $SourceHost
                DetectedUrlHost = $urlInfo.Host; DetectedSiteUrl = $urlInfo.SiteUrl; DetectedRelativePath = $urlInfo.RelativePath
            })
        } else {
            foreach ($sec in $sections) {
                $secName   = $sec.GetAttribute("name")
                $secId     = $sec.GetAttribute("ID")
                $secPath   = $sec.GetAttribute("path")
                $checkPath = if ($secPath) { $secPath } else { $nbPath }
                $isSource  = ($checkPath -like "*$SourceHost*") -or ($nbPath -like "*$SourceHost*")
                $urlInfo   = Parse-SharePointUrl $checkPath
                $rows.Add([PSCustomObject]@{
                    UserName = $UserName; UserDomain = $UserDomain; UserPrincipalName = $UserPrincipalName
                    ComputerName = $ComputerName; NotebookName = $nbName; NotebookId = $nbId; NotebookPath = $nbPath
                    SectionName = $secName; SectionId = $secId; SectionPath = $secPath
                    LastModifiedTime = $sec.GetAttribute("lastModifiedTime")
                    IsSourceTenant = $isSource; SourceTenantHost = $SourceHost
                    DetectedUrlHost = $urlInfo.Host; DetectedSiteUrl = $urlInfo.SiteUrl; DetectedRelativePath = $urlInfo.RelativePath
                })
            }
        }
    }
    return $rows
}

# ---------------------------------------------------------------------------
# Invoke-RobocopyBackup
# ---------------------------------------------------------------------------
function Invoke-RobocopyBackup {
    param([string]$Source, [string]$Destination, [string]$Label)
    if (-not (Test-Path $Source)) {
        Write-Log "Source path not found, skipping backup: $Source" -Severity WARNING
        return $false
    }
    Write-Log "Starting robocopy backup [$Label]: '$Source' -> '$Destination'"
    if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
    try {
        $proc = Start-Process -FilePath "robocopy.exe" `
            -ArgumentList @($Source, $Destination, "/E", "/R:1", "/W:1", "/XJ", "/NP", "/LOG+:$LogFile") `
            -Wait -PassThru -NoNewWindow
        $rc = $proc.ExitCode
        Write-Log "Robocopy [$Label] exit code: $rc"
        if ($rc -le 7) { Write-Log "Robocopy [$Label] succeeded."; return $true }
        Write-Log "Robocopy [$Label] reported an error (exit $rc)." -Severity ERROR
        return $false
    } catch {
        Write-Log "Robocopy [$Label] threw an exception: $_" -Severity ERROR
        return $false
    }
}

# ===========================================================================
# MAIN
# ===========================================================================

Write-Log "===== $ScriptName v$ScriptVersion — Start ====="
Write-Log "User       : $UserDomain\$UserName $(if ($UserPrincipalName) { "($UserPrincipalName)" })"
Write-Log "Computer   : $ComputerName"
Write-Log "Timestamp  : $Timestamp"
Write-Log "OutputRoot : $OutputRoot"
Write-Log "SourceHost : $SourceTenantHost"
Write-Log "StopOneNote: $ShouldStopOneNote"
Write-Log "StoreCache : $IncludeStoreAppCache"

$inventorySucceeded  = $false
$cacheSucceeded      = $false
$storeAppSucceeded   = $false
$comAvailable        = $false
$hierarchyExported   = $false
$inventoryCsvCreated = $false
$totalNotebooks      = 0
$totalSections       = 0
$sourceMatches       = 0

# --- 1. Stop OneNote ---
if ($ShouldStopOneNote) {
    Write-Log "--- Stopping OneNote processes ---"
    try { Stop-OneNoteProcesses } catch { Write-Log "Error during process stop: $_" -Severity ERROR }
} else {
    Write-Log "StopOneNote not requested; skipping process stop."
}

# --- 2. OneNote COM inventory ---
Write-Log "--- OneNote COM inventory ---"
$oneNoteApp   = $null
$hierarchyXml = ""
$rows         = @()

try {
    $oneNoteApp   = New-Object -ComObject OneNote.Application
    $comAvailable = $true
    Write-Log "OneNote COM object created successfully."
} catch {
    Write-Log "OneNote COM unavailable: $_" -Severity WARNING
    Write-Log "Skipping inventory; proceeding to cache backup."
}

if ($comAvailable) {
    try {
        [string]$hierarchyXml = ""
        $oneNoteApp.GetHierarchy("", 3, [ref]$hierarchyXml)
        Write-Log "GetHierarchy returned XML (length: $($hierarchyXml.Length) chars)."
        [System.IO.File]::WriteAllText($XmlFile, $hierarchyXml, [System.Text.Encoding]::UTF8)
        $hierarchyExported = $true
        Write-Log "Hierarchy XML saved: $XmlFile"
    } catch {
        Write-Log "Failed to retrieve/save hierarchy XML: $_" -Severity ERROR
    }

    if ($hierarchyExported -and $hierarchyXml) {
        try {
            $rows           = Get-OneNoteInventoryRows -HierarchyXml $hierarchyXml -SourceHost $SourceTenantHost
            $totalNotebooks = ($rows | Select-Object -ExpandProperty NotebookId -Unique | Measure-Object).Count
            $totalSections  = ($rows | Where-Object { $_.SectionId -ne "" } | Select-Object -ExpandProperty SectionId -Unique | Measure-Object).Count
            $sourceMatches  = ($rows | Where-Object { $_.IsSourceTenant -eq $true } | Measure-Object).Count
            Write-Log "Parsed: $totalNotebooks notebook(s), $totalSections section(s), $sourceMatches source-tenant match(es)."
            $rows | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
            $inventoryCsvCreated = $true
            $inventorySucceeded  = $true
            Write-Log "Inventory CSV saved: $CsvFile"
        } catch {
            Write-Log "Failed to parse hierarchy or export CSV: $_" -Severity ERROR
        }
    }

    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($oneNoteApp) | Out-Null
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    } catch { }
}

# --- 3. Cache backup ---
Write-Log "--- OneNote cache backup ---"
$cacheBackupAttempted = $true
$cacheSucceeded = Invoke-RobocopyBackup `
    -Source      (Join-Path $env:LOCALAPPDATA "Microsoft\OneNote\16.0") `
    -Destination (Join-Path $CacheBackupPath "OneNote_16.0") `
    -Label       "OneNote16Cache"

if (-not $cacheSucceeded) {
    Write-Log "OneNote 16.0 cache backup failed or source not found." -Severity WARNING
}

$storeAppCacheAttempted = $false
$storeAppSucceeded      = $false
if ($IncludeStoreAppCache) {
    $storeAppCacheAttempted = $true
    Write-Log "--- Store app OneNote cache backup ---"
    $storeAppSucceeded = Invoke-RobocopyBackup `
        -Source      (Join-Path $env:LOCALAPPDATA "Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalCache") `
        -Destination (Join-Path $CacheBackupPath "StoreApp_LocalCache") `
        -Label       "StoreAppOneNote"
    if (-not $storeAppSucceeded) {
        Write-Log "Store app cache backup failed or source not found." -Severity WARNING
    }
}

# --- 4. Summary JSON ---
$exitCode = if (-not $inventorySucceeded -and -not $cacheSucceeded) { 1 } else { 0 }
if ($exitCode -ne 0) { Write-Log "Neither inventory nor cache backup completed." -Severity WARNING }

$summary = [ordered]@{
    UserName                     = $UserName
    UserDomain                   = $UserDomain
    UserPrincipalName            = $UserPrincipalName
    ComputerName                 = $ComputerName
    Timestamp                    = $Timestamp
    SourceTenantHost             = $SourceTenantHost
    OneNoteComAvailable          = $comAvailable
    HierarchyExported            = $hierarchyExported
    InventoryCsvCreated          = $inventoryCsvCreated
    TotalNotebooks               = $totalNotebooks
    TotalSections                = $totalSections
    SourceTenantMatches          = $sourceMatches
    CacheBackupAttempted         = $cacheBackupAttempted
    CacheBackupSucceeded         = $cacheSucceeded
    StoreAppCacheBackupAttempted = $storeAppCacheAttempted
    StoreAppCacheBackupSucceeded = $storeAppSucceeded
    OutputPath                   = $InventoryPath
    ExitCode                     = $exitCode
}
try {
    $summary | ConvertTo-Json -Depth 5 | Set-Content -Path $SummaryFile -Encoding UTF8
    Write-Log "Summary JSON saved: $SummaryFile"
} catch { Write-Log "Failed to write Summary JSON: $_" -Severity ERROR }

# --- 5. PreCompleted JSON ---
try {
    [ordered]@{
        Completed     = $true
        CompletedAt   = (Get-Date -Format "o")
        ScriptName    = $ScriptName
        ScriptVersion = $ScriptVersion
        UserName      = $UserName
        ComputerName  = $ComputerName
        OutputPath    = $InventoryPath
    } | ConvertTo-Json -Depth 3 | Set-Content -Path $PreCompletedFile -Encoding UTF8
    Write-Log "PreCompleted marker saved: $PreCompletedFile"
} catch { Write-Log "Failed to write PreCompleted JSON: $_" -Severity ERROR }

Write-Log "===== $ScriptName — End (exit $exitCode) ====="
exit $exitCode
