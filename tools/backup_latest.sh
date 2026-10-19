#!/usr/bin/env bash
set -euo pipefail

BACKUP_TAG="${BACKUP_TAG:-backup-latest}"

# These files are independent data releases. An app-code rollback must never
# replace them with the versions from the older backup commit.
PROTECTED_PRESET_PATHS=(
    "sdk_check_presets.json"
    "package_log_presets.json"
    "services_checker/apk_check_presets.json"
    "services_checker/gradle_check_presets.json"
    "services_checker/gradle_lib_mapping.json"
    "services_checker/podfile_check_presets.json"
    "services_checker/podfile_lib_mapping.json"
    "services_checker/manifest_check_presets.json"
)

usage() {
    cat <<'USAGE'
Usage:
  bash tools/backup_latest.sh show
  bash tools/backup_latest.sh set [<commit-or-ref>]
  bash tools/backup_latest.sh revert --confirm

The repository keeps one managed backup ref: backup-latest.
`set` moves that ref and therefore replaces the previous backup.
`revert` is deliberately guarded; it resets app code while preserving the
protected preset files listed in BACKUP_POLICY.md.
USAGE
}

repo_root="$(git rev-parse --show-toplevel)"
cd "$repo_root"

resolve_commit() {
    git rev-parse --verify "$1^{commit}" 2>/dev/null
}

require_clean_worktree() {
    if [[ -n "$(git status --porcelain --untracked-files=all)" ]]; then
        echo "Refusing to continue: the worktree is not clean." >&2
        echo "Commit or stash changes before this operation." >&2
        exit 1
    fi
}

snapshot_protected_presets() {
    local snapshot_root="$1"
    local path
    for path in "${PROTECTED_PRESET_PATHS[@]}"; do
        if [[ -e "$path" ]]; then
            mkdir -p "$snapshot_root/$(dirname "$path")"
            cp -p "$path" "$snapshot_root/$path"
        else
            mkdir -p "$snapshot_root/$(dirname "$path")"
            : > "$snapshot_root/$path.__missing__"
        fi
    done
}

restore_protected_presets() {
    local snapshot_root="$1"
    local path
    for path in "${PROTECTED_PRESET_PATHS[@]}"; do
        if [[ -f "$snapshot_root/$path.__missing__" ]]; then
            rm -f -- "$path"
            continue
        fi
        mkdir -p "$(dirname "$path")"
        cp -p "$snapshot_root/$path" "$path"
    done
}

show_backup() {
    local commit
    if ! commit="$(resolve_commit "$BACKUP_TAG")"; then
        echo "No managed backup exists: $BACKUP_TAG" >&2
        exit 1
    fi
    git show -s --format="%H%n%ad%n%s" --date=iso "$commit"
}

set_backup() {
    local target="${1:-HEAD}"
    local commit

    if [[ "$target" == "HEAD" ]]; then
        require_clean_worktree
    fi
    if ! commit="$(resolve_commit "$target")"; then
        echo "Invalid commit or ref: $target" >&2
        exit 1
    fi

    git tag -f "$BACKUP_TAG" "$commit" >/dev/null
    echo "Managed backup updated: $BACKUP_TAG -> $(git rev-parse --short "$commit")"
    git show -s --format="%H%n%ad%n%s" --date=iso "$commit"
}

revert_to_backup() {
    if [[ "${1:-}" != "--confirm" ]]; then
        echo "Revert changes the current worktree. Run: bash tools/backup_latest.sh revert --confirm" >&2
        exit 2
    fi
    require_clean_worktree

    local commit
    if ! commit="$(resolve_commit "$BACKUP_TAG")"; then
        echo "No managed backup exists: $BACKUP_TAG" >&2
        exit 1
    fi
    local preset_snapshot
    preset_snapshot="$(mktemp -d "${TMPDIR:-/tmp}/eventchecker-presets.XXXXXX")"
    trap 'rm -rf -- "$preset_snapshot"' EXIT
    snapshot_protected_presets "$preset_snapshot"
    git reset --hard "$commit"
    restore_protected_presets "$preset_snapshot"
    trap - EXIT
    rm -rf -- "$preset_snapshot"
    echo "Preserved independent preset files:"
    printf '  %s\n' "${PROTECTED_PRESET_PATHS[@]}"
    echo "Reverted to managed backup: $BACKUP_TAG -> $(git rev-parse --short "$commit")"
}

case "${1:-}" in
    show)
        show_backup
        ;;
    set)
        set_backup "${2:-HEAD}"
        ;;
    revert)
        revert_to_backup "${2:-}"
        ;;
    *)
        usage
        exit 2
        ;;
esac
