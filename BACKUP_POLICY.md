# Latest backup policy

The project keeps exactly one managed backup reference: `backup-latest`.

## Current backup

Version `2.5.0(49)` is the initial backup:

`0fed557bcfa09ecdc45d996f2c8cb50c855188e3`

## When finalizing a backup

1. Commit the release and finish the required tests.
2. Run `bash tools/backup_latest.sh set <release-commit>`.
3. Do not create a new versioned backup tag. The command moves `backup-latest`, replacing the previous managed backup.

Use `bash tools/backup_latest.sh show` to verify the active backup.

## Reverting

After checking that the worktree is clean, run:

`bash tools/backup_latest.sh revert --confirm`

Historical `checkpoint-*` tags are kept as history; they are not additional managed backups.

## Preset protection

Preset data is independent from app code and is never part of an app-code rollback. The protected files are:

- `sdk_check_presets.json`
- `package_log_presets.json`
- `services_checker/apk_check_presets.json`
- `services_checker/gradle_check_presets.json`
- `services_checker/gradle_lib_mapping.json`
- `services_checker/podfile_check_presets.json`
- `services_checker/podfile_lib_mapping.json`
- `services_checker/manifest_check_presets.json`

When a preset list changes, update the files directly on Git and keep that update. `backup_latest.sh revert --confirm` snapshots and restores these files after reverting the app code, so an app revert must not roll back, overwrite, or delete them.
