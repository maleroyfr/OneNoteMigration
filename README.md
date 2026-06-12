# OneNote Migration

This repo contains the OneNote handling scripts for a Quest-based Microsoft 365 tenant-to-tenant migration.

## What the scripts do

- `Intune\Detect-OneNotePreMigration.ps1`
  - Runs in the logged-on user context.
  - Checks whether the OneNote local cache has changed since the last backup.
  - Triggers remediation only when a new backup is needed.

- `Intune\Remediate-OneNotePreMigration.ps1`
  - Runs in the logged-on user context.
  - Performs the pre-migration inventory and local backup.
  - Writes the completion marker used by detection.

- `Post-OneNote-Cutover.ps1`
  - Runs after Quest profile migration, ReACL, and DUA app reconfiguration.
  - Stops OneNote, captures an emergency backup, quarantines the local OneNote state, and optionally reopens validated target notebooks.

## End-to-end process for IT

### 1. Before migration
1. Deploy the Intune proactive remediation pair.
2. Allow detection to run on schedule.
3. When remediation fires, it inventories OneNote, captures the local cache state, and writes the backup artifacts under `C:\ProgramData\GFD-MIG\OneNoteMigration`.

### 2. During migration
1. Quest migrates the user profile.
2. ReACL is applied.
3. DUA application reconfiguration is completed.

### 3. After migration
1. Run `Post-OneNote-Cutover.ps1` in the logged-on user context.
2. The script stops OneNote, captures an emergency backup, and quarantines the old local OneNote state.
3. If a validated mapping CSV is provided, the script can optionally reopen target notebooks.

## Folder layout

```
C:\ProgramData\GFD-MIG\OneNoteMigration\
  Intune\
    Detect-OneNotePreMigration.ps1
    Remediate-OneNotePreMigration.ps1
  Post-Migration outputs are created at runtime:
    PostMigration\<User>_<Computer>_<Timestamp>\
    Quarantine\<User>_<Computer>_<Timestamp>\
    Logs\
```

## Output and safety

- No `.onepkg` export
- No cloud upload
- No deletion of OneNote data
- No direct editing of OneNote cache or database files
- Scripts are designed to be rerun safely

## Recommended operational flow

1. Run the Intune remediation before cutover so the user has a clean, backed-up local state.
2. Complete migration and app reconfiguration.
3. Run the post-cutover script to quarantine old OneNote state and open validated target notebooks if required.

## Notes

- All scripts are intended for the logged-on user context.
- If OneNote COM is unavailable, the scripts continue with best-effort file operations and logging.
- Review the generated JSON and CSV files under `C:\ProgramData\GFD-MIG\OneNoteMigration` for audit and troubleshooting.
