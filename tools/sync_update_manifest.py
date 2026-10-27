#!/usr/bin/env python3
"""Generate the v2.5 remote-update manifest from one immutable Git commit.

Preset JSON and mapping files are deliberately excluded.  They are live data
read from Git at runtime and must not be copied into an app release payload.
Run this only after the release source has been committed; it does not build
or publish anything.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_MANIFEST = ROOT / "Updates_2_5" / "remote_manifest.json"
IMMUTABLE_REF_RE = re.compile(r"^[0-9a-f]{40}$")
PAYLOAD_URL_BASES = (
    "https://github.com/indiezvn/Eventchecker/raw/{ref}",
    "https://raw.githubusercontent.com/indiezvn/Eventchecker/{ref}",
    "https://raw.githubusercontent.com/trucbm/Eventchecker/{ref}",
    "https://cdn.jsdelivr.net/gh/indiezvn/Eventchecker@{ref}",
)
PROTECTED_PRESET_PATHS = {
    "sdk_check_presets.json",
    "package_log_presets.json",
    "services_checker/apk_check_presets.json",
    "services_checker/gradle_check_presets.json",
    "services_checker/gradle_lib_mapping.json",
    "services_checker/podfile_check_presets.json",
    "services_checker/podfile_lib_mapping.json",
    "services_checker/manifest_check_presets.json",
}


def _git_blob(ref: str, relative: str) -> bytes:
    try:
        return subprocess.check_output(
            ["git", "show", f"{ref}:{relative}"],
            cwd=str(ROOT),
        )
    except (OSError, subprocess.CalledProcessError) as exc:
        raise SystemExit(f"Git payload is unavailable at {ref}: {relative}: {exc}") from exc


def _sha256(payload: bytes) -> str:
    import hashlib

    return hashlib.sha256(payload).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--payload-ref", required=True, help="40-character commit SHA containing the release payload")
    parser.add_argument("--version", help="Release version string; defaults to the current manifest version")
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)
    args = parser.parse_args()

    payload_ref = str(args.payload_ref or "").strip().lower()
    if not IMMUTABLE_REF_RE.fullmatch(payload_ref):
        raise SystemExit("--payload-ref must be a 40-character lowercase Git commit SHA")

    manifest_path = args.manifest if args.manifest.is_absolute() else ROOT / args.manifest
    try:
        current = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise SystemExit(f"Cannot read manifest {manifest_path}: {exc}") from exc

    paths = []
    for item in current.get("files") or []:
        relative = str(item.get("path") or "").strip().replace("\\", "/")
        if relative and relative not in PROTECTED_PRESET_PATHS and relative not in paths:
            paths.append(relative)
    if not paths:
        raise SystemExit("Manifest has no non-preset payload files to synchronize")

    files = []
    for relative in paths:
        payload = _git_blob(payload_ref, relative)
        urls = [f"{base.format(ref=payload_ref)}/{relative}" for base in PAYLOAD_URL_BASES]
        files.append(
            {
                "path": relative,
                "url": urls[1],
                "urls": urls,
                "sha256": _sha256(payload),
            }
        )

    output = {
        "version": str(args.version or current.get("version") or "").strip(),
        "payload_ref": payload_ref,
        "update_enabled": bool(current.get("update_enabled", True)),
        "update_message": str(current.get("update_message") or ""),
        "files": files,
    }
    if not output["version"]:
        raise SystemExit("Manifest version is empty; pass --version explicitly")

    manifest_path.write_text(json.dumps(output, indent=2) + "\n", encoding="utf-8")
    print(f"manifest synchronized: {manifest_path} ({len(files)} payloads, ref {payload_ref})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
