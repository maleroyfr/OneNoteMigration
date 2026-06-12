# Intune Proactive Remediation — OneNote Pre-Migration

## Folder contents

| File | Role |
|---|---|
| `Detect-OneNotePreMigration.ps1` | Detection script |
| `Remediate-OneNotePreMigration.ps1` | Remediation script |

## How it works

```
Intune Proactive Remediation  (runs on schedule, e.g. every hour)
  │
  ├─ Detection runs in USER context
  │
  │   No backup ever done?
  │     └─ Exit 1 → Non-compliant → trigger Remediation
  │
  │   Last backup < MinBackupIntervalHours ago?
  │     └─ Exit 0 → Compliant (too soon, skip)
  │
  │   OneNote cache modified AFTER last backup?
  │     ├─ Yes → Exit 1 → Non-compliant → trigger Remediation
  │     └─ No  → Exit 0 → Compliant (nothing changed)
  │
  └─ Remediation runs in USER context (only when Detection = 1)
        Runs full inventory + cache backup
        Exit 0 → device marked Compliant
        Exit 1 → Intune reports remediation failed
```

## Pre-requisite: deploy the main script

**None.** `Remediate-OneNotePreMigration.ps1` is fully self-contained.
Upload it directly to Intune alongside `Detect-OneNotePreMigration.ps1`.

## Intune Proactive Remediation settings

| Setting | Value |
|---|---|
| **Name** | `GFD-MIG - OneNote Pre-Migration Inventory` |
| **Detection script** | `Detect-OneNotePreMigration.ps1` |
| **Remediation script** | `Remediate-OneNotePreMigration.ps1` |
| **Run script in 64-bit PowerShell** | Yes |
| **Run this script using the logged-on credentials** | Yes (user context) |
| **Enforce script signature check** | Optional (sign if required by policy) |
| **Schedule** | Every **1 hour** (detection is cheap; remediation only fires on change) |

> ⚠️ **User context is mandatory.** The COM API (`OneNote.Application`) and
> the user's OneNote cache (`%LOCALAPPDATA%`) are only accessible in the
> logged-on user session. Running in SYSTEM context will not work.

## Re-running / resetting

To force a re-run for a user, delete (or rename) any `PreCompleted.json` files
under `C:\ProgramData\GFD-MIG\OneNoteMigration\Inventory\`. Detection will
then return exit 1 again on the next evaluation cycle.
