# Claude Code instructions for EventInspector

These rules are mandatory for work in this repository.

## Scope and safety

- Work only in this repository and preserve unrelated user changes. Inspect `git status` before making changes.
- Code, build, and Git work use repository files and CLI tools. Do not use desktop UI automation.
- A browser may be used when the user asks for browser work (for example Google Drive/Sheets, Apps Script, or checking a local server page):
  - Prefer Claude's built-in browser; it is isolated from the user's Chrome.
  - Use the user's Chrome (Claude in Chrome) only when the task needs the user's signed-in session or the user asks for it. Work in a new tab, and never close, reload, or navigate the user's existing tabs or windows.
  - Never enter passwords or approve Google OAuth/permission prompts; the user does that step.
- Use USB-connected devices only. Do not enable or use network-paired devices.
- Do not stop, replace, or modify the official running app/service while testing local changes. Keep local processes and ports separate from the official app.
- When running shell commands, prefix them with `rtk` as required by the repository setup. If `rtk` is unavailable, report that before falling back to a direct command.

## Build, test, and release gates

- Do not build, package, publish, push, or trigger an official update unless the user explicitly requests that action in the current request.
- For a requested code change, test locally first and use the relevant harness/checks. Do not claim success without command evidence.
- The normal flow is: inspect -> change -> local test/harness -> show result -> wait for confirmation -> official release/update.
- For normal local work, run `python3 tools/eventchecker_harness.py --profile local`; this does not pretend an uncommitted worktree is the published release.
- For a release gate, first commit the app-only source, then run `python3 tools/sync_update_manifest.py --payload-ref <commit-sha> --version <release-version>`, followed by `python3 tools/eventchecker_harness.py --profile release`.
- Official releases go out through the in-app remote update only: commit the app payload, regenerate `Updates_2_5/remote_manifest.json`, and publish it on `main`. Do not rebuild or redistribute the `.app`/`.dmg`/`.exe` for a release. A new packaged build is allowed only when the user explicitly asks for one in the current request.
- Build numbers go up by exactly 1 per published update: the next release is the build live on `main` + 1 (e.g. 79 -> 80). Never skip a number. A rollback candidate prepared in advance also uses live + 1, but it is only a candidate and does not reserve that number: if the next release is published instead, rebuild the rollback candidate as the new live + 1 with the previous release's content.
- A request to fix code does not authorize a build or release. A request to start local authorizes only the local server/process.
- Before any official release/update, verify the update path on every requested platform and preserve a tested rollback path.
- Never use a destructive reset or overwrite to recover a release. Preserve user changes and diagnose first.

## Presets and rollback

- Never modify, revert, overwrite, delete, or regenerate SDK, Gradle, Podfile, manifest, mapping, Bright/SKAds, or Package Log preset data unless the user explicitly requests a preset change.
- Reverting app code must not revert or overwrite presets or mappings. Presets are independent data and must remain at their current version during an app rollback.
- Presets are live data, not release payload. The app release contains only the code that reads, validates, refreshes, and displays preset data; it must not bundle preset JSON or include preset JSON in the remote-update manifest.
- After a release is approved, bring the app logic/release infrastructure back to the single official `main` line. Do not merge a release branch wholesale when it carries preset-file changes; keep preset files on their independent Git data path.
- A preset update is a direct data update on its own Git path. It must be picked up by the app's remote-read/reload flow without rebuilding or republishing the app.
- Never restore or recreate anything related to releases `230` or `240`; those releases are removed from the project.
- Treat these files as protected preset data unless the user explicitly asks to edit them:
  - `sdk_check_presets.json`
  - `package_log_presets.json`
  - `services_checker/apk_check_presets.json`
  - `services_checker/gradle_check_presets.json`
  - `services_checker/gradle_lib_mapping.json`
  - `services_checker/podfile_check_presets.json`
  - `services_checker/podfile_lib_mapping.json`
  - `services_checker/manifest_check_presets.json`

## Git and worktrees

- Do not run `git reset --hard`, force-push, overwrite a user file, or delete data without explicit approval.
- Keep unrelated changes intact. Before commit or push, show the relevant status/diff and run the relevant validation.
- Do not commit generated or local-only data such as `__pycache__/`, `.venv/`, build caches, `dist/` artifacts, or user state unless explicitly requested.
- Do not commit or push merely because a change is complete; commit/push only when the user asks.

## Project workflow

- Read `CODEX_SETUP_CHECKLIST.md` for setup and test workflow and `BACKUP_POLICY.md` for rollback rules when those topics are relevant.
- Use Python 3.11 or 3.12 and the repository requirements. For local development, use the documented virtual environment and run `python desktop_app.py` from the repository root.
- Run macOS or Windows packaging scripts only after an explicit build request for that platform.
- Keep device/platform state isolated: switching Android/iOS must reset only runtime UI state, must not reuse a device from the other platform, and must not change presets.
- When diagnosing missing logs, first preserve raw input and verify parser/recording behavior separately from display filters. Do not silently drop split or multiline log records.

## Required handoff

After every change, report:

1. the exact files changed;
2. whether the action was local-only or official;
3. the validation actually run and its result;
4. any remaining blocker or user confirmation required.

