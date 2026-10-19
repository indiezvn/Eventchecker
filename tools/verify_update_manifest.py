#!/usr/bin/env python3
"""Verify that every remote-update manifest hash matches its Git payload."""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = ROOT / "Updates_2_5" / "remote_manifest.json"
IMMUTABLE_REF_RE = re.compile(r"^[0-9a-f]{40}$")
TEXT_PAYLOAD_SUFFIXES = {".py", ".json"}


def _sha256(path: Path) -> str:
    # GitHub serves repository text blobs with LF.  Normalize a Windows
    # checkout before hashing so core.autocrlf cannot break a release build.
    if path.suffix.lower() in TEXT_PAYLOAD_SUFFIXES:
        digest = hashlib.sha256(path.read_bytes().replace(b"\r\n", b"\n"))
        return digest.hexdigest()

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    try:
        manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"update manifest: cannot read {MANIFEST_PATH}: {exc}", file=sys.stderr)
        return 1

    errors: list[str] = []
    payload_ref = str(manifest.get("payload_ref") or "").strip().lower()
    if not IMMUTABLE_REF_RE.fullmatch(payload_ref):
        errors.append("manifest payload_ref must be a 40-character immutable Git commit SHA")

    files = manifest.get("files") or []
    if not files:
        errors.append("manifest has no files")

    for item in files:
        relative = str(item.get("path") or "").strip()
        expected = str(item.get("sha256") or "").strip().lower()
        if not relative:
            errors.append("manifest contains an entry without path")
            continue
        if not expected:
            errors.append(f"{relative}: missing sha256")
            continue

        urls = [item.get("url"), *(item.get("urls") or [])]
        for url in urls:
            url = str(url or "").strip()
            if not url:
                errors.append(f"{relative}: missing payload URL")
            elif payload_ref and payload_ref not in url:
                errors.append(f"{relative}: payload URL is not pinned to payload_ref: {url}")
            elif "/main/" in url or "@main/" in url:
                errors.append(f"{relative}: mutable main payload URL is not allowed: {url}")

        payload = (ROOT / relative).resolve()
        try:
            payload.relative_to(ROOT)
        except ValueError:
            errors.append(f"{relative}: path escapes repository root")
            continue
        if not payload.is_file():
            errors.append(f"{relative}: payload is missing")
            continue

        actual = _sha256(payload)
        if actual != expected:
            errors.append(f"{relative}: expected {expected}, actual {actual}")

    if errors:
        print("update manifest: FAIL", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    print(f"update manifest: PASS ({len(files)} payloads)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
