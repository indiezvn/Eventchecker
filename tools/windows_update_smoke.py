#!/usr/bin/env python3
"""Run the portable updater against a real local HTTP server on Windows.

This deliberately keeps a file open in the active payload directory.  The
test proves that a Windows portable update downloads the complete payload and
hands off a new directory without trying to replace the directory in use.
"""

from __future__ import annotations

import copy
import http.server
import json
import os
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import remote_update as updater  # noqa: E402


class _PayloadHandler(http.server.BaseHTTPRequestHandler):
    manifest: bytes = b""
    payloads: dict[str, bytes] = {}

    def do_GET(self):  # noqa: N802 - BaseHTTPRequestHandler API
        relative = urlparse(self.path).path.lstrip("/")
        if relative == "remote_manifest.json":
            body = self.manifest
        elif relative in self.payloads:
            body = self.payloads[relative]
        else:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_args):
        return


def _release_payload_bytes(payload_ref: str, relative_path: str) -> bytes:
    """Read the exact immutable Git blob referenced by the manifest."""
    try:
        result = subprocess.run(
            ["git", "show", f"{payload_ref}:{relative_path}"],
            cwd=ROOT,
            check=True,
            capture_output=True,
        )
        return result.stdout
    except (OSError, subprocess.CalledProcessError) as exc:
        raise AssertionError(
            f"manifest payload blob is unavailable for {relative_path} at {payload_ref}: {exc}"
        ) from exc


def main() -> int:
    simulate_windows = os.name != "nt" and os.getenv("EVENTINSPECTOR_WINDOWS_SMOKE_SIMULATE") == "1"
    if os.name != "nt" and not simulate_windows:
        raise SystemExit("windows_update_smoke must run on Windows (or set EVENTINSPECTOR_WINDOWS_SMOKE_SIMULATE=1 for local code-path testing)")

    manifest_guard = subprocess.run(
        [sys.executable, str(ROOT / "tools" / "verify_update_manifest.py")],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
    )
    if manifest_guard.returncode != 0:
        raise AssertionError(
            "Windows updater smoke refused a stale manifest:\n"
            + (manifest_guard.stdout or "")
            + (manifest_guard.stderr or "")
        )

    manifest = json.loads((ROOT / "Updates_2_5" / "remote_manifest.json").read_text(encoding="utf-8"))
    expected_build = updater._extract_build_number(manifest.get("version"))
    if expected_build is None:
        raise AssertionError(f"Canonical manifest has no release build: {manifest.get('version')!r}")
    payload_ref = str(manifest.get("payload_ref") or "").strip()
    payloads = {
        str(item["path"]): _release_payload_bytes(payload_ref, str(item["path"]))
        for item in manifest.get("files") or []
        if item.get("path")
    }

    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), _PayloadHandler)
    base_url = f"http://127.0.0.1:{server.server_address[1]}"
    local_manifest = copy.deepcopy(manifest)
    for item in local_manifest.get("files") or []:
        item["url"] = f"{base_url}/{item['path']}"
        item["urls"] = []
    _PayloadHandler.manifest = json.dumps(local_manifest).encode("utf-8")
    _PayloadHandler.payloads = payloads
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    original_local_app_data = os.environ.get("LOCALAPPDATA")
    original_manifest_candidates = updater._candidate_manifest_urls
    original_download_first = updater._download_first
    original_download_verified = updater._download_verified
    original_updater_os_name = updater.os.name
    active_handle = None
    try:
        if simulate_windows:
            # All standard-library modules used by this script are imported
            # before this switch.  Only the updater's platform branch is
            # changed; the HTTP requests and filesystem operations stay real.
            updater.os.name = "nt"
        temporary_support = tempfile.TemporaryDirectory(prefix="eventinspector_windows_smoke_")
        support_root = temporary_support.name
        try:
            os.environ["LOCALAPPDATA"] = support_root
            updater._candidate_manifest_urls = lambda _cfg: [f"{base_url}/remote_manifest.json"]

            active_dir = os.path.join(support_root, "EventInspector", "updates_v250")
            os.makedirs(active_dir, exist_ok=True)
            active_path = os.path.join(active_dir, "active_payload.locked")
            active_handle = open(active_path, "w", encoding="utf-8")
            active_handle.write("active")
            active_handle.flush()
            updater._save_state({
                "version": "2026-08-28-1-2.5.0-53",
                "update_dir": active_dir,
                "files": ["active_payload.locked"],
            })

            result = updater.check_for_updates()
            if result.get("status") != "updated":
                raise AssertionError(f"Windows portable update failed: {result}")

            prepared = updater.get_prepared_update_info()
            prepared_dir = prepared.get("update_dir")
            if not prepared_dir or prepared_dir == active_dir:
                raise AssertionError(f"Windows updater did not hand off a new payload directory: {prepared}")
            if not os.path.basename(prepared_dir).startswith("updates_v250_"):
                raise AssertionError(f"Unexpected Windows staged payload directory: {prepared_dir}")
            if prepared.get("build") != expected_build:
                raise AssertionError(f"Windows staged build mismatch: expected {expected_build}, got {prepared}")
            if not os.path.exists(active_path):
                raise AssertionError("Windows updater modified the active payload")
            for relative in payloads:
                if not os.path.exists(os.path.join(prepared_dir, relative)):
                    raise AssertionError(f"Windows staged payload is missing: {relative}")

            second = updater.check_for_updates()
            if second.get("status") != "up_to_date":
                raise AssertionError(f"Windows updater redownloaded an unchanged payload: {second}")

            print("windows_update_smoke: PASS")
            print(json.dumps({"version": prepared.get("version"), "update_dir": prepared_dir}, sort_keys=True))
        finally:
            # Windows cannot remove an open file while TemporaryDirectory cleans
            # up.  Close the deliberately held active-payload handle first.
            if active_handle is not None:
                active_handle.close()
                active_handle = None
            temporary_support.cleanup()
    finally:
        if active_handle is not None:
            active_handle.close()
        updater.os.name = original_updater_os_name
        updater._candidate_manifest_urls = original_manifest_candidates
        updater._download_first = original_download_first
        updater._download_verified = original_download_verified
        if original_local_app_data is None:
            os.environ.pop("LOCALAPPDATA", None)
        else:
            os.environ["LOCALAPPDATA"] = original_local_app_data
        server.shutdown()
        server.server_close()
        thread.join(timeout=2)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
