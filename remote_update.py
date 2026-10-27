import json
import os
import re
import hashlib
import time
import sys
import shutil
import tempfile

import requests

APP_NAME = "EventInspector"
CHANNEL_ID = "v250"
CONFIG_FILENAME = "remote_update_config_v250.json"
STATE_FILENAME = "update_state_v250.json"
UPDATES_DIRNAME = "updates_v250"
DOWNLOAD_ATTEMPTS = 3
DOWNLOAD_RETRY_DELAYS = (0.35, 1.0)
# The repository moved to the company organization (indiezvn).  The old
# trucbm URLs stay as fallbacks; GitHub redirects them to the new repo.
DEFAULT_MANIFEST_URLS = [
    "https://github.com/indiezvn/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
    "https://raw.githubusercontent.com/indiezvn/Eventchecker/main/Updates_2_5/remote_manifest.json",
    "https://github.com/trucbm/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
    "https://raw.githubusercontent.com/trucbm/Eventchecker/main/Updates_2_5/remote_manifest.json",
    "https://cdn.jsdelivr.net/gh/indiezvn/Eventchecker@main/Updates_2_5/remote_manifest.json",
    "https://cdn.jsdelivr.net/gh/trucbm/Eventchecker@main/Updates_2_5/remote_manifest.json",
]
DEFAULT_MANIFEST_URL = DEFAULT_MANIFEST_URLS[0]
DEFAULT_FILE_URL_BASES = [
    "https://github.com/indiezvn/Eventchecker/raw/main",
    "https://raw.githubusercontent.com/indiezvn/Eventchecker/main",
    "https://github.com/trucbm/Eventchecker/raw/main",
    "https://raw.githubusercontent.com/trucbm/Eventchecker/main",
    "https://cdn.jsdelivr.net/gh/indiezvn/Eventchecker@main",
    "https://cdn.jsdelivr.net/gh/trucbm/Eventchecker@main",
]
KNOWN_CHANNELS = ("v250",)


def _runtime_profile():
    return {
        "channel_id": CHANNEL_ID,
        "config_filename": CONFIG_FILENAME,
        "state_filename": STATE_FILENAME,
        "updates_dirname": UPDATES_DIRNAME,
        "manifest_urls": DEFAULT_MANIFEST_URLS,
        "default_manifest_url": DEFAULT_MANIFEST_URL,
        "file_url_bases": DEFAULT_FILE_URL_BASES,
    }


def _user_data_dir():
    configured = os.getenv("EVENTINSPECTOR_DATA_DIR", "").strip()
    if configured:
        return os.path.abspath(os.path.expandvars(os.path.expanduser(configured)))
    if os.name == "nt":
        base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, APP_NAME)
    if sys.platform == "darwin":
        return os.path.join(os.path.expanduser("~/Library/Application Support"), APP_NAME)
    return os.path.join(os.path.expanduser("~"), f".{APP_NAME.lower()}")


def _config_paths():
    profile = _runtime_profile()
    user_dir = _user_data_dir()
    env_name = f"EVENTINSPECTOR_UPDATE_CONFIG_{profile['channel_id'].upper()}"
    return [
        os.getenv(env_name),
        os.path.join(user_dir, profile["config_filename"]),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), profile["config_filename"]),
    ]


def _load_config():
    profile = _runtime_profile()
    for p in _config_paths():
        if p and os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                cfg.setdefault("enabled", True)
                cfg.setdefault("manifest_url", profile["default_manifest_url"])
                cfg.setdefault("manifest_urls", profile["manifest_urls"])
                cfg.setdefault("timeout_sec", 120)
                # Always prefer checking remote on launch. Existing user configs
                # may still contain stale throttling values from older builds.
                cfg["min_interval_sec"] = 0
                return cfg
    return {
        "enabled": True,
        "manifest_url": profile["default_manifest_url"],
        "manifest_urls": profile["manifest_urls"],
        "timeout_sec": 120,
        "min_interval_sec": 0,
    }


def _ensure_user_config_template():
    profile = _runtime_profile()
    user_dir = _user_data_dir()
    os.makedirs(user_dir, exist_ok=True)
    cfg_path = os.path.join(user_dir, profile["config_filename"])
    desired = {
        "enabled": True,
        "manifest_url": profile["default_manifest_url"],
        "manifest_urls": profile["manifest_urls"],
        "timeout_sec": 120,
        "min_interval_sec": 0,
    }
    current = {}
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                current = json.load(f)
        except Exception:
            current = {}
    current.update(desired)
    _write_staged_file(cfg_path, json.dumps(current, indent=2).encode("utf-8"))
    return cfg_path


def _state_path():
    return os.path.join(_user_data_dir(), _runtime_profile()["state_filename"])


def _channel_paths(channel_id):
    user_dir = _user_data_dir()
    return {
        "config": os.path.join(user_dir, f"remote_update_config_{channel_id}.json"),
        "state": os.path.join(user_dir, f"update_state_{channel_id}.json"),
        "updates": os.path.join(user_dir, f"updates_{channel_id}"),
        "updates_tmp": os.path.join(user_dir, f"updates_{channel_id}_tmp"),
    }


def _remove_path(path):
    if not path or not os.path.exists(path):
        return
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
        return
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    except Exception:
        try:
            os.unlink(path)
        except Exception:
            pass


def _safe_update_version_token(version):
    """Make a manifest version safe to use in a Windows directory name."""
    token = re.sub(r"[^A-Za-z0-9._-]+", "_", str(version or "unknown")).strip("._")
    return token[:80] or "unknown"


def _create_update_workspace(user_dir, manifest_version):
    """Return a writable staging directory without touching the active payload.

    Windows cannot reliably rename/replace a directory whose files are still
    open by the running portable process (or by antivirus scanning).  A unique
    directory per release lets the updater finish and hand off the new path in
    the channel state; the next process loads it after restart.
    """
    updates_dirname = _runtime_profile()["updates_dirname"]
    if os.name == "nt":
        prefix = f"{updates_dirname}_{_safe_update_version_token(manifest_version)}_"
        return tempfile.mkdtemp(prefix=prefix, dir=user_dir), True

    tmp_update_dir = os.path.join(user_dir, f"{updates_dirname}_tmp")
    _remove_path(tmp_update_dir)
    os.makedirs(tmp_update_dir, exist_ok=True)
    return tmp_update_dir, False


def clear_update_cache(include_all_channels=False):
    channel_ids = KNOWN_CHANNELS if include_all_channels else (_runtime_profile()["channel_id"],)
    for channel_id in channel_ids:
        for path in _channel_paths(channel_id).values():
            _remove_path(path)


def _load_state():
    path = _state_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_state(state):
    user_dir = _user_data_dir()
    os.makedirs(user_dir, exist_ok=True)
    payload = json.dumps(state, indent=2).encode("utf-8")
    _write_staged_file(_state_path(), payload)


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _extract_build_number(text):
    match = re.search(r"\d+\.\d+\.\d+-(\d+)$", str(text or "").strip())
    if match:
        try:
            return int(match.group(1))
        except Exception:
            return None
    return None


def _requested_from_bundle_build():
    try:
        return int(str(os.getenv("EVENTINSPECTOR_BUNDLED_BUILD") or "").strip())
    except Exception:
        return None


def _bundle_build_is_detected():
    return str(os.getenv("EVENTINSPECTOR_BUNDLED_BUILD_SOURCE") or "").strip().lower() == "detected"


def _cache_busted_url(url):
    """Avoid stale CDN/proxy responses while keeping the configured URL intact."""
    value = str(url or "")
    if not value.startswith(("http://", "https://")):
        return value
    separator = "&" if "?" in value else "?"
    return f"{value}{separator}eventinspector_refresh={int(time.time() * 1000)}"


def _download(url, timeout):
    # Handle Google Drive confirm page for large files
    session = requests.Session()
    request_url = _cache_busted_url(url)
    headers = {
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "User-Agent": "EventInspector-Updater/2.5",
    }
    try:
        r = session.get(
            request_url,
            allow_redirects=True,
            headers=headers,
            timeout=timeout,
        )
        r.raise_for_status()
        if r.headers.get("content-type", "").lower().startswith("text/html"):
            m = re.search(r"confirm=([0-9A-Za-z_]+)", r.text)
            if not m:
                raise ValueError("unexpected_html_response")
            confirm = m.group(1)
            sep = "&" if "?" in request_url else "?"
            url2 = f"{request_url}{sep}confirm={confirm}"
            r = session.get(
                url2,
                allow_redirects=True,
                headers=headers,
                timeout=timeout,
            )
            r.raise_for_status()
            if r.headers.get("content-type", "").lower().startswith("text/html"):
                raise ValueError("unexpected_html_response")
        return r.content
    finally:
        session.close()


def _unique_urls(urls):
    seen = set()
    out = []
    for url in urls or []:
        if not url or url in seen:
            continue
        seen.add(url)
        out.append(url)
    return out


def _candidate_manifest_urls(cfg):
    profile = _runtime_profile()
    urls = []
    urls.extend(cfg.get("manifest_urls") or [])
    single = (cfg.get("manifest_url") or "").strip()
    if single:
        urls.append(single)
    urls.extend(profile["manifest_urls"])
    return _unique_urls(urls)


def _default_repo_file_urls(rel_path):
    rel = (rel_path or "").lstrip("/")
    return [f"{base}/{rel}" for base in _runtime_profile()["file_url_bases"]]


def _download_first(urls, timeout):
    last_error = None
    for url in _unique_urls(urls):
        for attempt in range(DOWNLOAD_ATTEMPTS):
            try:
                return _download(url, timeout), url
            except Exception as exc:
                last_error = exc
                if attempt + 1 < DOWNLOAD_ATTEMPTS:
                    time.sleep(DOWNLOAD_RETRY_DELAYS[min(attempt, len(DOWNLOAD_RETRY_DELAYS) - 1)])
    if last_error:
        raise last_error
    raise ValueError("no_download_urls")


def _download_verified(urls, timeout, expected_sha256=""):
    last_error = None
    normalized_sha = str(expected_sha256 or "").strip().lower()
    for url in _unique_urls(urls):
        for attempt in range(DOWNLOAD_ATTEMPTS):
            tmp_path = None
            try:
                data = _download(url, timeout)
                if normalized_sha:
                    fd, tmp_path = tempfile.mkstemp(prefix="eventinspector_update_", suffix=".tmp")
                    os.close(fd)
                    with open(tmp_path, "wb") as f:
                        f.write(data)
                    actual_sha = _sha256_file(tmp_path).lower()
                    if actual_sha != normalized_sha:
                        raise ValueError(f"sha256_mismatch:{actual_sha}")
                return data, url
            except Exception as exc:
                last_error = exc
                if attempt + 1 < DOWNLOAD_ATTEMPTS:
                    time.sleep(DOWNLOAD_RETRY_DELAYS[min(attempt, len(DOWNLOAD_RETRY_DELAYS) - 1)])
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
    if last_error:
        raise last_error
    raise ValueError("no_verified_download_urls")


def _write_staged_file(target, data):
    """Write one staged payload with short retries for antivirus/file scanners."""
    tmp = f"{target}.{os.getpid()}.{time.time_ns()}.tmp"
    last_error = None
    try:
        for attempt in range(DOWNLOAD_ATTEMPTS):
            try:
                with open(tmp, "wb") as f:
                    f.write(data)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp, target)
                return
            except Exception as exc:
                last_error = exc
                if attempt + 1 < DOWNLOAD_ATTEMPTS:
                    time.sleep(DOWNLOAD_RETRY_DELAYS[min(attempt, len(DOWNLOAD_RETRY_DELAYS) - 1)])
        if last_error:
            raise last_error
        raise OSError("staged_file_write_failed")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def load_prepared_update_dir():
    _ensure_user_config_template()
    cfg = _load_config()

    if not cfg.get("enabled"):
        return None

    state = _load_state()
    update_dir = state.get("update_dir")
    expected_files = state.get("files") or []
    if not update_dir or not os.path.isdir(update_dir):
        return None
    for rel_path in expected_files:
        if rel_path and not os.path.exists(os.path.join(update_dir, rel_path)):
            return None
    return update_dir


def _existing_update_matches_manifest(update_dir, manifest_files):
    if not update_dir or not os.path.isdir(update_dir):
        return False
    for item in manifest_files or []:
        rel_path = item.get("path")
        if not rel_path:
            continue
        target = os.path.join(update_dir, rel_path)
        if not os.path.exists(target):
            return False
        expected_sha = str(item.get("sha256") or "").strip().lower()
        if expected_sha:
            try:
                actual_sha = _sha256_file(target).lower()
            except Exception:
                return False
            if actual_sha != expected_sha:
                return False
    return True


def get_prepared_update_info():
    update_dir = load_prepared_update_dir()
    if not update_dir:
        return {}
    state = _load_state()
    return {
        "update_dir": update_dir,
        "version": state.get("version"),
        "files": state.get("files") or [],
        "requested_from_bundle_build": state.get("requested_from_bundle_build"),
        "build": _extract_build_number(state.get("version")),
    }


def check_for_updates(force_refresh=False):
    if force_refresh:
        clear_update_cache(include_all_channels=True)

    _ensure_user_config_template()
    cfg = _load_config()

    if not cfg.get("enabled"):
        return None

    manifest_urls = _candidate_manifest_urls(cfg)
    if not manifest_urls:
        return None

    timeout = float(cfg.get("timeout_sec", 120))
    state = _load_state()

    try:
        manifest_bytes, manifest_url = _download_first(manifest_urls, timeout)
        manifest = json.loads(manifest_bytes.decode("utf-8"))
    except Exception as exc:
        return {
            "ok": False,
            "status": "error",
            "error": "manifest_download_failed",
            "detail": f"{type(exc).__name__}: {exc}",
            "update_dir": load_prepared_update_dir(),
        }

    user_dir = _user_data_dir()
    update_dir = os.path.join(user_dir, _runtime_profile()["updates_dirname"])
    state_version = state.get("version")
    manifest_version = manifest.get("version")
    existing_update_dir = state.get("update_dir") or update_dir
    manifest_files = manifest.get("files", [])
    requested_from_bundle_build = _requested_from_bundle_build()
    manifest_build = _extract_build_number(manifest_version)
    if (
        _bundle_build_is_detected()
        and requested_from_bundle_build is not None
        and manifest_build is not None
        and manifest_build < requested_from_bundle_build
    ):
        return {
            "ok": True,
            "status": "up_to_date",
            "version": manifest_version,
            "update_dir": load_prepared_update_dir(),
        }
    if state_version == manifest_version and _existing_update_matches_manifest(existing_update_dir, manifest_files):
        state.update({
            "last_check": time.time(),
            "version": manifest_version,
            "update_dir": existing_update_dir,
            "manifest_url": manifest_url,
            "files": [item.get("path") for item in manifest_files if item.get("path")],
        })
        if requested_from_bundle_build is not None:
            state["requested_from_bundle_build"] = requested_from_bundle_build
        _save_state(state)
        return {"ok": True, "status": "up_to_date", "version": manifest_version, "update_dir": existing_update_dir}

    try:
        os.makedirs(user_dir, exist_ok=True)
        tmp_update_dir, uses_standalone_windows_payload = _create_update_workspace(user_dir, manifest_version)
        if uses_standalone_windows_payload:
            update_dir = tmp_update_dir
    except Exception as exc:
        return {
            "ok": False,
            "status": "error",
            "error": "staging_failed",
            "detail": f"{type(exc).__name__}: {exc}",
            "update_dir": load_prepared_update_dir(),
        }

    ok = True
    download_error = None
    for item in manifest_files:
        rel_path = item.get("path")
        url = item.get("url")
        urls = list(item.get("urls") or [])
        sha256 = item.get("sha256")
        if not rel_path or not url:
            if not rel_path:
                ok = False
                break

        target = os.path.join(tmp_update_dir, rel_path)
        os.makedirs(os.path.dirname(target), exist_ok=True)
        try:
            candidate_urls = []
            if url:
                candidate_urls.append(url)
            candidate_urls.extend(urls)
            candidate_urls.extend(_default_repo_file_urls(rel_path))
            data, _used_url = _download_verified(candidate_urls, timeout, sha256)
            _write_staged_file(target, data)
        except Exception as exc:
            ok = False
            download_error = f"{rel_path}: {type(exc).__name__}: {exc}"
            break

    if ok:
        if not uses_standalone_windows_payload:
            # Keep the current payload intact until the complete new directory is in place.
            previous_update_dir = f"{update_dir}.previous"
            try:
                _remove_path(previous_update_dir)
                if os.path.exists(update_dir):
                    os.replace(update_dir, previous_update_dir)
                os.replace(tmp_update_dir, update_dir)
                _remove_path(previous_update_dir)
            except Exception as exc:
                try:
                    if not os.path.exists(update_dir) and os.path.exists(previous_update_dir):
                        os.replace(previous_update_dir, update_dir)
                except Exception:
                    pass
                _remove_path(tmp_update_dir)
                return {
                    "ok": False,
                    "status": "error",
                    "error": "replace_failed",
                    "detail": f"{type(exc).__name__}: {exc}",
                    "update_dir": load_prepared_update_dir(),
                }

        state.update({
            "last_check": time.time(),
            "version": manifest.get("version"),
            "update_dir": update_dir,
            "manifest_url": manifest_url,
            "files": [item.get("path") for item in manifest_files if item.get("path")],
        })
        if requested_from_bundle_build is not None:
            state["requested_from_bundle_build"] = requested_from_bundle_build
        _save_state(state)
        return {"ok": True, "status": "updated", "version": manifest.get("version"), "update_dir": update_dir}

    _remove_path(tmp_update_dir)
    return {
        "ok": False,
        "status": "error",
        "error": "download_failed",
        "detail": download_error or "no_verified_download_urls",
        "update_dir": load_prepared_update_dir(),
    }


def check_and_prepare_updates():
    result = check_for_updates()
    return result.get("update_dir")

    return state.get("update_dir")
