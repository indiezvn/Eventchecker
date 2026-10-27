import json
import re
import subprocess
import threading
import time
import atexit
import requests
import html
import importlib.util
import importlib
from collections import deque
from flask import Flask, render_template_string, request, jsonify, Response
from flask_socketio import SocketIO
from flask_cors import CORS
from openpyxl import load_workbook
import os
import sys
import shutil
import logging
import sqlite3
import socket
import platform
import signal
import uuid
import plistlib
import tempfile
import zipfile
import glob
try:
    import webview
except Exception:
    webview = None
from pathlib import Path
from queue import Empty, Full, Queue

# Khởi tạo ứng dụng Flask và SocketIO
app = Flask(__name__)
CORS(app)
try:
    socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')
except ValueError:
    # Fallback for environments where 'threading' isn't accepted
    socketio = SocketIO(app, cors_allowed_origins="*")

# --- CẤU HÌNH LOAD ADS (GOOGLE SHEET) ---
G_SHEET_URL = "https://script.google.com/macros/s/AKfycbyLMM9nLAjS9Zhwr4-J6ikjqBSpO7ZCNaNeHKTsfKltiIa0OniDBSrzjvqfvpg87Epl/exec"
RECORD_SHEET_CONNECT_TIMEOUT_SEC = 3
RECORD_SHEET_READ_TIMEOUT_SEC = 8
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SDK_CHECK_PRESETS_FILENAME = "sdk_check_presets.json"
SDK_CHECK_PRESETS_REMOTE_URLS = [
    "https://raw.githubusercontent.com/indiezvn/Eventchecker/main/sdk_check_presets.json",
    "https://github.com/indiezvn/Eventchecker/raw/main/sdk_check_presets.json",
    "https://raw.githubusercontent.com/trucbm/Eventchecker/main/sdk_check_presets.json",
    "https://github.com/trucbm/Eventchecker/raw/main/sdk_check_presets.json",
]
PACKAGE_LOG_PRESETS_FILENAME = "package_log_presets.json"
PACKAGE_LOG_PRESETS_REMOTE_URLS = [
    "https://raw.githubusercontent.com/indiezvn/Eventchecker/main/package_log_presets.json",
    "https://github.com/indiezvn/Eventchecker/raw/main/package_log_presets.json",
    "https://raw.githubusercontent.com/trucbm/Eventchecker/main/package_log_presets.json",
    "https://github.com/trucbm/Eventchecker/raw/main/package_log_presets.json",
]


def _sanitize_sdk_check_presets(raw):
    if not isinstance(raw, dict):
        return {}
    presets = {}
    for name, preset in raw.items():
        if not isinstance(preset, dict):
            continue
        preset_name = str(name).strip()
        platform_name = str(preset.get("platform") or "android").strip().lower()
        lines = preset.get("lines")
        if not preset_name or platform_name not in {"android", "ios", "all"} or not isinstance(lines, list):
            continue
        clean_lines = [str(line).rstrip("\r") for line in lines if str(line).strip()]
        clean_preset = {
            "platform": platform_name,
            "lines": clean_lines,
        }
        bright = preset.get("bright")
        if isinstance(bright, dict):
            bright_version = str(bright.get("version") or "").strip()
            bright_identifiers = []
            seen_identifiers = set()
            raw_identifiers = bright.get("skadnetwork_identifiers")
            if isinstance(raw_identifiers, list):
                for value in raw_identifiers:
                    identifier = str(value or "").strip()
                    key = identifier.casefold()
                    if not identifier or key in seen_identifiers:
                        continue
                    seen_identifiers.add(key)
                    bright_identifiers.append(identifier)
            if bright_version or bright_identifiers:
                clean_preset["bright"] = {
                    "version": bright_version,
                    "skadnetwork_identifiers": bright_identifiers,
                }
        presets[preset_name] = clean_preset
    return presets


def _compare_bright_skad_preset(actual_version, actual_identifiers, preset_name="", preset=None):
    """Compare an IPA's BrightData version and app SKAdNetwork list to a preset."""
    bright = preset.get("bright") if isinstance(preset, dict) else None
    if not isinstance(bright, dict):
        return {
            "preset_name": str(preset_name or ""),
            "rows": [],
            "extra_identifiers": [],
            "expected_identifier_count": 0,
            "actual_identifier_count": 0,
            "passed_count": 0,
            "total_count": 0,
        }

    expected_version = str(bright.get("version") or "").strip()
    expected_identifiers = []
    expected_seen = set()
    for value in bright.get("skadnetwork_identifiers") or []:
        identifier = str(value or "").strip()
        key = identifier.casefold()
        if not identifier or key in expected_seen:
            continue
        expected_seen.add(key)
        expected_identifiers.append(identifier)

    actual_version = str(actual_version or "").strip()
    actual_values = []
    actual_seen = set()
    for value in actual_identifiers or []:
        identifier = str(value or "").strip()
        key = identifier.casefold()
        if not identifier or key in actual_seen:
            continue
        actual_seen.add(key)
        actual_values.append(identifier)

    rows = []
    if expected_version:
        version_status = "PASSED" if actual_version == expected_version else ("MISSING" if not actual_version else "FAILED")
        rows.append({
            "kind": "BrightData",
            "label": "BrightData",
            "actual": actual_version or "NOT FOUND",
            "expected": expected_version,
            "status": version_status,
        })

    for identifier in expected_identifiers:
        found = identifier.casefold() in actual_seen
        rows.append({
            "kind": "SKAdNetwork",
            "label": identifier,
            "actual": identifier if found else "NOT FOUND",
            "expected": identifier,
            "status": "PASSED" if found else "MISSING",
        })

    extras = [identifier for identifier in actual_values if identifier.casefold() not in expected_seen]
    passed_count = sum(1 for row in rows if row.get("status") == "PASSED")
    return {
        "preset_name": str(preset_name or ""),
        "rows": rows,
        "extra_identifiers": extras,
        "expected_identifier_count": len(expected_identifiers),
        "actual_identifier_count": len(actual_values),
        "passed_count": passed_count,
        "total_count": len(rows),
    }


def _sdk_check_preset_file_candidates():
    candidates = []
    env_path = os.getenv("SDK_CHECK_PRESETS_PATH")
    if env_path:
        candidates.append(env_path)
    candidates.append(os.path.join(SCRIPT_DIR, SDK_CHECK_PRESETS_FILENAME))

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, SDK_CHECK_PRESETS_FILENAME))
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), SDK_CHECK_PRESETS_FILENAME))
    return candidates


def _load_sdk_check_presets():
    for path in _sdk_check_preset_file_candidates():
        if not path or not os.path.exists(path):
            continue
        try:
            with open(path, "r", encoding="utf-8") as handle:
                presets = _sanitize_sdk_check_presets(json.load(handle))
            if presets:
                return presets
        except Exception as exc:
            logging.warning("Failed to load SDK check presets from %s: %s", path, exc)
    return {}


def _fetch_sdk_check_presets(force_remote=False):
    if os.getenv("SDK_CHECK_PRESETS_PATH") and not force_remote:
        local_presets = _load_sdk_check_presets()
        if local_presets:
            return local_presets, "local"

    cache_bust = time.time_ns()
    for base_url in SDK_CHECK_PRESETS_REMOTE_URLS:
        try:
            separator = "&" if "?" in base_url else "?"
            response = requests.get(
                f"{base_url}{separator}cache_bust={cache_bust}",
                headers={
                    "Accept": "application/json",
                    "Cache-Control": "no-cache, no-store, max-age=0",
                    "Pragma": "no-cache",
                    "Accept-Encoding": "identity",
                },
                timeout=8,
            )
            response.raise_for_status()
            presets = _sanitize_sdk_check_presets(response.json())
            if presets:
                # Git is the single source of truth.  Do not merge bundled or
                # stale local preset data into a live response.
                return presets, "github"
        except Exception as exc:
            logging.warning("Failed to fetch SDK check presets from %s: %s", base_url, exc)
    return _load_sdk_check_presets(), "local"


def _user_data_dir():
    configured = os.getenv("EVENTINSPECTOR_DATA_DIR", "").strip()
    if configured:
        return os.path.abspath(os.path.expandvars(os.path.expanduser(configured)))
    if os.name == "nt":
        base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, "EventInspector")
    if sys.platform == "darwin":
        return os.path.join(os.path.expanduser("~/Library/Application Support"), "EventInspector")
    return os.path.join(os.path.expanduser("~"), ".eventinspector")


def _package_log_presets_file_candidates():
    """Return explicit development/offline fallback locations for Package Log data."""
    candidates = []
    env_path = os.getenv("PACKAGE_LOG_PRESETS_PATH")
    if env_path:
        candidates.append(env_path)

    update_dir = os.getenv("EVENTINSPECTOR_UPDATE_DIR")
    if update_dir:
        candidates.append(os.path.join(update_dir, PACKAGE_LOG_PRESETS_FILENAME))

    # These locations are fallback/override paths only.  Normal app startup
    # fetches the live catalog from Git before consulting them.
    candidates.append(os.path.join(_user_data_dir(), PACKAGE_LOG_PRESETS_FILENAME))
    candidates.append(os.path.join(SCRIPT_DIR, PACKAGE_LOG_PRESETS_FILENAME))

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, PACKAGE_LOG_PRESETS_FILENAME))
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.append(os.path.join(exe_dir, PACKAGE_LOG_PRESETS_FILENAME))
        if sys.platform == "darwin":
            resources_dir = os.path.abspath(os.path.join(exe_dir, "..", "Resources"))
            candidates.append(os.path.join(resources_dir, PACKAGE_LOG_PRESETS_FILENAME))

    unique = []
    seen = set()
    for path in candidates:
        path = os.path.abspath(os.path.expanduser(str(path))) if path else ""
        if path and path not in seen:
            seen.add(path)
            unique.append(path)
    return unique


def _package_log_platform_value(entry, platform_name, code):
    raw_value = entry.get(platform_name)
    if raw_value is None:
        raw_value = entry.get(f"{platform_name}_value")
    if raw_value is None:
        raw_value = entry.get(f"{platform_name}_id")

    label = entry.get(f"{platform_name}_label")
    if isinstance(raw_value, dict):
        label = raw_value.get("label") or raw_value.get("name") or label
        raw_value = raw_value.get("id") or raw_value.get("value") or raw_value.get("package") or raw_value.get("bundle")

    identifier = str(raw_value or "").strip()
    if not identifier:
        return None
    label = str(label or "").strip() or (f"{code} - {identifier}" if code else identifier)
    return {"id": identifier, "label": label}


def _sanitize_package_log_presets(raw):
    if isinstance(raw, dict):
        projects = raw.get("projects")
    else:
        projects = raw
    if not isinstance(projects, list):
        return []

    presets = []
    seen = set()
    for project in projects:
        if not isinstance(project, dict):
            continue
        code = str(project.get("code") or project.get("key") or project.get("name") or "").strip()
        clean = {"code": code}
        for platform_name in ("android", "ios"):
            platform_value = _package_log_platform_value(project, platform_name, code)
            if not platform_value:
                continue
            identity = (platform_name, platform_value["id"].casefold())
            if identity in seen:
                continue
            seen.add(identity)
            clean[platform_name] = platform_value
        if len(clean) > 1:
            presets.append(clean)
    return presets


def _load_package_log_presets():
    for path in _package_log_presets_file_candidates():
        if not os.path.isfile(path):
            continue
        try:
            with open(path, "r", encoding="utf-8") as handle:
                presets = _sanitize_package_log_presets(json.load(handle))
            if presets:
                return presets
        except Exception as exc:
            logging.warning("Failed to load Package Log presets from %s: %s", path, exc)
    return []


def _fetch_package_log_presets(force_remote=False):
    local_presets = _load_package_log_presets()
    if os.getenv("PACKAGE_LOG_PRESETS_PATH") and not force_remote and local_presets:
        return local_presets, "local"

    cache_bust = time.time_ns()
    for base_url in PACKAGE_LOG_PRESETS_REMOTE_URLS:
        try:
            separator = "&" if "?" in base_url else "?"
            response = requests.get(
                f"{base_url}{separator}cache_bust={cache_bust}",
                headers={
                    "Accept": "application/json",
                    "Cache-Control": "no-cache, no-store, max-age=0",
                    "Pragma": "no-cache",
                    "Accept-Encoding": "identity",
                },
                timeout=8,
            )
            response.raise_for_status()
            presets = _sanitize_package_log_presets(response.json())
            if presets:
                return presets, "github"
        except Exception as exc:
            logging.warning("Failed to fetch Package Log presets from %s: %s", base_url, exc)
    return local_presets, "offline"

def _resolve_default_params_path():
    env_path = os.getenv("DEFAULT_PARAMS_XLSX_PATH")
    if env_path and os.path.exists(env_path):
        return env_path

    filename = "Default event + Default Params.xlsx"

    update_dir = os.getenv("EVENTINSPECTOR_UPDATE_DIR")
    if update_dir:
        upd = os.path.join(update_dir, filename)
        if os.path.exists(upd):
            return upd

    candidates = []

    # 1) Project folder (dev)
    candidates.append(os.path.join(SCRIPT_DIR, filename))

    # 2) Installed app folder (packaged)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.append(os.path.join(exe_dir, filename))

        # macOS: Resources folder inside .app bundle
        if sys.platform == "darwin":
            resources_dir = os.path.abspath(os.path.join(exe_dir, "..", "Resources"))
            candidates.append(os.path.join(resources_dir, filename))

        candidates.append(os.path.join(_user_data_dir(), "profiles", filename))

    # 3) PyInstaller temp bundle path
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, filename))

    # 4) Common user locations
    candidates.extend([
        os.path.join(os.path.expanduser("~/Downloads"), filename),
        os.path.join(os.path.expanduser("~/Documents"), filename),
    ])

    for p in candidates:
        if p and os.path.exists(p):
            return p

    return candidates[0]

DEFAULT_PARAMS_XLSX = _resolve_default_params_path()
DEFAULT_PARAM_FILL = "FFFCE5CD"
# Keep the full provider order here so the two temporarily hidden providers
# can be restored without changing the parser or table layout again.
DEFAULT_AD_EVENT_PROVIDER_ORDER = ("ironsource", "max", "cloudx", "ascendx")
DEFAULT_AD_EVENT_HIDDEN_PROVIDERS = frozenset({"cloudx", "ascendx"})
DEFAULT_AD_EVENT_PROVIDERS = tuple(
    provider
    for provider in DEFAULT_AD_EVENT_PROVIDER_ORDER
    if provider not in DEFAULT_AD_EVENT_HIDDEN_PROVIDERS
)
DEFAULT_AD_EVENT_PROVIDER_LABELS = {
    "ironsource": "Ironsource",
    "max": "MAX",
    "cloudx": "Cloudx",
    "ascendx": "Ascendx",
}
DEFAULT_AD_EVENT_FORMATS = ("interstitial", "rewarded", "banner", "mrec")
DEFAULT_AD_EVENT_FORMAT_LABELS = {
    "interstitial": "Interstitial",
    "rewarded": "Rewarded",
    "banner": "Banner",
    "mrec": "MREC",
}
DEFAULT_AD_EVENT_PARAM_KEYS = {"ad_platform", "ad_format"}
DEFAULT_AD_EVENT_SIMPLE_CATEGORIES = ("audioads", "inplayads")
DEFAULT_AD_EVENT_SIMPLE_CATEGORY_LABELS = {
    "audioads": "AudioAds",
    "inplayads": "InplayAds",
}
DEFAULT_AD_EVENT_SIMPLE_PROVIDERS = {
    "audioads": ("odeeo", "audiomob"),
    "inplayads": ("adverty", "gadsme"),
}
DEFAULT_AD_EVENT_SIMPLE_PROVIDER_LABELS = {
    "odeeo": "Odeeo",
    "audiomob": "Audiomob",
    "adverty": "Adverty",
    "gadsme": "Gadsme",
}
REMOTE_UPDATE_CONFIG_FILENAME = "remote_update_config_v250.json"
# Repository moved to the company organization (indiezvn); the old trucbm
# URLs stay as fallbacks (GitHub redirects them).
DEFAULT_REMOTE_MANIFEST_URL = "https://raw.githubusercontent.com/indiezvn/Eventchecker/main/Updates_2_5/remote_manifest.json"
DEFAULT_REMOTE_MANIFEST_URLS = [
    "https://raw.githubusercontent.com/indiezvn/Eventchecker/main/Updates_2_5/remote_manifest.json",
    "https://github.com/indiezvn/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
    "https://raw.githubusercontent.com/trucbm/Eventchecker/main/Updates_2_5/remote_manifest.json",
    "https://github.com/trucbm/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
    "https://cdn.jsdelivr.net/gh/indiezvn/Eventchecker@main/Updates_2_5/remote_manifest.json",
    "https://cdn.jsdelivr.net/gh/trucbm/Eventchecker@main/Updates_2_5/remote_manifest.json",
]
APP_AUDIT_CONFIG_FILENAME = "app_audit_config.json"
DEFAULT_AUDIT_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzKzf7KAJvy3gu1vfaQ7wHR0tk2mzR_UapoVIfjjjO-aFG2NHKmJQKYLvfsBwgi7heY/exec"
DEFAULT_AUDIT_SHEET_NAME = "AppOpenAudit"
DEFAULT_AUDIT_TIMEOUT_SEC = 5
DEFAULT_PUBLIC_IP_URLS = [
    "https://api.ipify.org?format=json",
    "https://ifconfig.me/all.json",
]


def _runtime_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return SCRIPT_DIR


def _resolve_profiles_dir():
    base_dir = _user_data_dir() if getattr(sys, "frozen", False) else SCRIPT_DIR
    profiles_dir = os.path.join(base_dir, "profiles")
    try:
        os.makedirs(profiles_dir, exist_ok=True)

        # Migrate editable profiles from older bundles without writing back to
        # the signed .app directory.
        if getattr(sys, "frozen", False):
            legacy_dir = os.path.join(_runtime_app_dir(), "profiles")
            if os.path.isdir(legacy_dir) and os.path.abspath(legacy_dir) != os.path.abspath(profiles_dir):
                for legacy_file in Path(legacy_dir).glob("*.xlsx"):
                    target = os.path.join(profiles_dir, legacy_file.name)
                    if not os.path.exists(target):
                        try:
                            shutil.copyfile(legacy_file, target)
                        except Exception:
                            pass
        return profiles_dir
    except Exception:
        fallback = os.path.join(_user_data_dir(), "profiles")
        os.makedirs(fallback, exist_ok=True)
        return fallback


PROFILE_DIR = _resolve_profiles_dir()
active_profile_name = None
active_profile_path = None
active_profile_game_name = ""
profile_revision = 0


def _app_audit_config_path():
    return os.path.join(_user_data_dir(), APP_AUDIT_CONFIG_FILENAME)


def _machine_id_path():
    return os.path.join(_user_data_dir(), "machine_id.txt")


def _ensure_machine_id():
    path = _machine_id_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                value = f.read().strip()
            if value:
                return value
        except Exception:
            pass
    value = str(uuid.uuid4())
    os.makedirs(_user_data_dir(), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(value)
    return value


def _ensure_app_audit_config_template():
    os.makedirs(_user_data_dir(), exist_ok=True)
    path = _app_audit_config_path()
    data = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    data.setdefault("enabled", True)
    data.setdefault("webhook_url", DEFAULT_AUDIT_WEBHOOK_URL)
    data.setdefault("sheet_name", DEFAULT_AUDIT_SHEET_NAME)
    data.setdefault("timeout_sec", DEFAULT_AUDIT_TIMEOUT_SEC)
    data.setdefault("public_ip_urls", DEFAULT_PUBLIC_IP_URLS)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    return path


def _load_app_audit_config():
    _ensure_app_audit_config_template()
    data = {}
    try:
        with open(_app_audit_config_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = {}
    env_webhook = (os.getenv("EVENTINSPECTOR_AUDIT_WEBHOOK_URL") or "").strip()
    env_sheet_name = (os.getenv("EVENTINSPECTOR_AUDIT_SHEET_NAME") or "").strip()
    data["enabled"] = bool(data.get("enabled", True))
    data["webhook_url"] = env_webhook or str(data.get("webhook_url") or DEFAULT_AUDIT_WEBHOOK_URL).strip()
    data["sheet_name"] = env_sheet_name or str(data.get("sheet_name") or DEFAULT_AUDIT_SHEET_NAME).strip()
    try:
        data["timeout_sec"] = float(data.get("timeout_sec", DEFAULT_AUDIT_TIMEOUT_SEC))
    except Exception:
        data["timeout_sec"] = DEFAULT_AUDIT_TIMEOUT_SEC
    urls = data.get("public_ip_urls") or DEFAULT_PUBLIC_IP_URLS
    data["public_ip_urls"] = [str(url).strip() for url in urls if str(url).strip()]
    if not data["public_ip_urls"]:
        data["public_ip_urls"] = list(DEFAULT_PUBLIC_IP_URLS)
    return data


def _detect_public_ip(urls, timeout_sec):
    session = requests.Session()
    last_error = ""
    for url in urls:
        try:
            response = session.get(url, timeout=timeout_sec)
            response.raise_for_status()
            content_type = response.headers.get("content-type", "").lower()
            if "json" in content_type:
                payload = response.json()
                for key in ("ip", "IP", "address"):
                    value = payload.get(key)
                    if value:
                        return str(value).strip(), ""
            text = response.text.strip()
            if text:
                return text, ""
        except Exception as exc:
            last_error = str(exc)
            continue
    return "", last_error or "ip_lookup_failed"


def _current_app_build():
    try:
        with open(__file__, "r", encoding="utf-8", errors="ignore") as f:
            data = f.read(300000)
        matches = re.findall(r"v\d+\.\d+\.\d+\((\d+)\)", data)
        if matches:
            return max(int(item) for item in matches)
    except Exception:
        pass
    return None


def _record_app_open_audit_once():
    try:
        cfg = _load_app_audit_config()
        if not cfg.get("enabled", True):
            logging.info("App-open audit disabled by config")
            return
        webhook_url = str(cfg.get("webhook_url") or "").strip()
        now = time.time()
        public_ip, public_ip_error = _detect_public_ip(
            cfg.get("public_ip_urls") or DEFAULT_PUBLIC_IP_URLS,
            cfg.get("timeout_sec", DEFAULT_AUDIT_TIMEOUT_SEC),
        )
        payload = {
            "action": "app_open_audit",
            "sheet_name": cfg.get("sheet_name") or DEFAULT_AUDIT_SHEET_NAME,
            "opened_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(now)),
            "opened_at_unix": int(now),
            "public_ip": public_ip or "unavailable",
            "public_ip_error": public_ip_error,
            "hostname": socket.gethostname(),
            "machine_id": _ensure_machine_id(),
            "os_name": platform.system(),
            "os_version": platform.platform(),
            "app_build": _current_app_build(),
        }
        if not webhook_url:
            logging.info("App-open audit payload prepared locally without webhook: %s", json.dumps(payload, ensure_ascii=False))
            return
        requests.post(webhook_url, json=payload, timeout=cfg.get("timeout_sec", DEFAULT_AUDIT_TIMEOUT_SEC))
        logging.info("App-open audit sent successfully: %s", json.dumps(payload, ensure_ascii=False))
    except Exception:
        logging.exception("App-open audit failed")


def _package_history_dir():
    history_dir = os.path.join(_user_data_dir(), "package_log_history")
    try:
        os.makedirs(history_dir, exist_ok=True)
        legacy_db = os.path.join(_runtime_app_dir(), "package_log_history", "package_logs.sqlite3")
        target_db = os.path.join(history_dir, "package_logs.sqlite3")
        if os.path.exists(legacy_db) and not os.path.exists(target_db):
            try:
                shutil.copyfile(legacy_db, target_db)
            except Exception:
                pass
        test_file = os.path.join(history_dir, ".write_test")
        with open(test_file, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test_file)
        return history_dir
    except Exception:
        fallback = os.path.join(tempfile.gettempdir(), "EventInspector", "package_log_history")
        os.makedirs(fallback, exist_ok=True)
        return fallback


PACKAGE_LOG_DB_PATH = os.path.join(_package_history_dir(), "package_logs.sqlite3")


def _normalize_remote_update_config():
    try:
        user_dir = _user_data_dir()
        os.makedirs(user_dir, exist_ok=True)
        cfg_path = os.path.join(user_dir, REMOTE_UPDATE_CONFIG_FILENAME)
        cfg = {}
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = {}
        cfg["enabled"] = True
        cfg["manifest_url"] = DEFAULT_REMOTE_MANIFEST_URL
        cfg["manifest_urls"] = DEFAULT_REMOTE_MANIFEST_URLS
        cfg["timeout_sec"] = 120
        cfg["min_interval_sec"] = 0
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"WARNING: Failed to normalize updater config: {e}")


def _load_remote_update_module():
    candidates = []
    update_dir = os.getenv("EVENTINSPECTOR_UPDATE_DIR")
    if update_dir:
        candidates.append(os.path.join(update_dir, "remote_update.py"))
    candidates.append(os.path.join(SCRIPT_DIR, "remote_update.py"))

    for idx, path in enumerate(candidates):
        if not path or not os.path.exists(path):
            continue
        try:
            spec = importlib.util.spec_from_file_location(f"eventinspector_remote_update_{idx}", path)
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                return module
        except Exception:
            continue

    import remote_update
    return remote_update


SERVICES_CHECKER_DEFAULT_PORT = 5010
SERVICES_CHECKER_EXTERNAL_URL_DEFAULT = "http://127.0.0.1:5000"
SERVICES_CHECKER_APP_RELATIVE_PATH = os.path.join(
    "Shared drives",
    "IndieZ - Tester",
    "Automation & Tools",
    "Checker",
    "Services checker",
    "app.py",
)
SERVICES_CHECKER_KEYSTORE_FILENAME = "my-key.keystore"
SERVICES_CHECKER_EXTERNAL_ENV = "EVENTINSPECTOR_ALLOW_EXTERNAL_SERVICES_CHECKER"
_services_checker_lock = threading.Lock()
_services_checker_state = {
    "module": None,
    "thread": None,
    "port": None,
    "url": None,
    "process": None,
    "source": "",
    "error": "",
}


def _services_checker_external_sources_enabled():
    """Keep Drive/launcher hosts opt-in after the bundled migration."""
    return os.getenv(SERVICES_CHECKER_EXTERNAL_ENV, "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def _services_checker_drive_app_candidates():
    """Find the shared Services Checker app.py before using the bundled copy.

    The shared app owns the matching keystore and bundletool files.  Keep this
    lookup here as a second line of defence in case a launcher is unavailable
    on a portable client.
    """
    candidates = []
    configured = os.getenv("EVENTINSPECTOR_SERVICES_APP_PATH", "").strip()
    if configured:
        candidates.append(os.path.expandvars(os.path.expanduser(configured)))

    if os.name == "nt":
        roots = [
            os.getenv("GOOGLE_DRIVE_ROOT", ""),
            os.getenv("GOOGLE_DRIVE_MOUNT", ""),
            os.getenv("GDRIVE_ROOT", ""),
        ]
        roots.extend(f"{letter}:\\" for letter in "CDEFGHIJKLMNOPQRSTUVWXYZ")
        profile = os.path.expanduser("~")
        roots.extend([
            os.path.join(profile, "Google Drive"),
            os.path.join(profile, "GoogleDrive"),
        ])
        for root in roots:
            if not root:
                continue
            candidates.append(os.path.join(root, SERVICES_CHECKER_APP_RELATIVE_PATH))
            candidates.append(os.path.join(root, "Google Drive", SERVICES_CHECKER_APP_RELATIVE_PATH))
    else:
        cloud_roots = glob.glob(os.path.expanduser("~/Library/CloudStorage/GoogleDrive-*"))
        cloud_roots.extend([
            os.path.expanduser("~/Google Drive"),
            os.path.expanduser("~/GoogleDrive"),
        ])
        for root in cloud_roots:
            candidates.append(os.path.join(root, SERVICES_CHECKER_APP_RELATIVE_PATH))

    seen = set()
    return [
        path for path in candidates
        if path and path not in seen and not seen.add(path) and os.path.isfile(path)
    ]


def _services_checker_launcher_is_drive_aware(path):
    """Return whether a saved launcher resolves the shared Drive app."""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as handle:
            source = handle.read(128 * 1024)
        return (
            "Shared drives" in source
            and "IndieZ - Tester" in source
            and "Services checker" in source
            and "app.py" in source
        )
    except (OSError, UnicodeError):
        return False


def _services_checker_file_candidates():
    candidates = []
    update_dir = os.getenv("EVENTINSPECTOR_UPDATE_DIR")
    if update_dir:
        candidates.append(os.path.join(update_dir, "services_checker", "app.py"))

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, "services_checker", "app.py"))
    candidates.append(os.path.join(SCRIPT_DIR, "services_checker", "app.py"))
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), "services_checker", "app.py"))

    if _services_checker_external_sources_enabled():
        candidates.extend(_services_checker_drive_app_candidates())

    seen = set()
    return [path for path in candidates if path and not (path in seen or seen.add(path))]


def _services_checker_command_candidates():
    if not _services_checker_external_sources_enabled():
        return []

    candidates = []
    configured = os.getenv("EVENTINSPECTOR_SERVICES_COMMAND", "").strip()
    if configured:
        candidates.append(os.path.expanduser(configured))
    if os.name == "nt":
        candidates.extend([
            os.path.expanduser("~/Desktop/Androidchecker.cmd"),
            os.path.expanduser("~/Downloads/Androidchecker.cmd"),
        ])
    else:
        candidates.append(os.path.expanduser("~/Desktop/AndroidTool.command"))

    # A previously imported host may be stale and still launch the bundled
    # checker. Prefer the current Drive-aware launchers; keep a saved host as
    # a fallback only when it also points at the shared app.py.
    saved_host = _services_checker_saved_host_path()
    if os.path.isfile(saved_host) and (
        _services_checker_launcher_is_drive_aware(saved_host)
        or not _services_checker_drive_app_candidates()
    ):
        candidates.append(saved_host)

    seen = set()
    return [path for path in candidates if path and not (path in seen or seen.add(path))]


def _services_checker_saved_host_path():
    filename = "Androidchecker.cmd" if os.name == "nt" else "AndroidTool.command"
    return os.path.join(_user_data_dir(), "services_checker_host", filename)


def _services_checker_host_status():
    path = _services_checker_saved_host_path()
    return {
        "saved": os.path.isfile(path),
        "filename": os.path.basename(path),
    }


def _save_services_checker_host(upload):
    filename = os.path.basename(upload.filename or "")
    suffix = os.path.splitext(filename)[1].lower()
    if suffix not in {".command", ".sh", ".cmd", ".bat"}:
        raise ValueError("host_file_must_be_command_or_batch_script")

    target_dir = os.path.dirname(_services_checker_saved_host_path())
    os.makedirs(target_dir, exist_ok=True)
    target = _services_checker_saved_host_path()
    temp_path = ""
    try:
        with tempfile.NamedTemporaryFile(
            prefix="host_",
            suffix=suffix,
            dir=target_dir,
            delete=False,
        ) as temp_file:
            temp_path = temp_file.name
        upload.save(temp_path)
        if os.path.getsize(temp_path) > 2 * 1024 * 1024:
            raise ValueError("host_file_too_large")
        if os.name != "nt":
            os.chmod(temp_path, 0o755)
        os.replace(temp_path, target)
        return target
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def _stop_services_checker_process(process):
    if not process or process.poll() is not None:
        return

    try:
        if os.name == "nt":
            process.terminate()
        else:
            os.killpg(os.getpgid(process.pid), signal.SIGTERM)
    except (OSError, ProcessLookupError):
        return

    deadline = time.time() + 5
    while time.time() < deadline and process.poll() is None:
        time.sleep(0.1)

    if process.poll() is None:
        try:
            if os.name == "nt":
                process.kill()
            else:
                os.killpg(os.getpgid(process.pid), signal.SIGKILL)
        except (OSError, ProcessLookupError):
            pass


def _reload_services_checker():
    with _services_checker_lock:
        owned_process = _services_checker_state.get("process")
        _stop_services_checker_process(owned_process)
        restarted = owned_process is not None
        _services_checker_state.update({
            "module": None,
            "thread": None,
            "port": None,
            "url": None,
            "process": None,
            "source": "",
            "error": "",
        })

    # Let the old Flask/debug-reloader process release port 5000 before
    # launching the selected host script again.
    if restarted:
        deadline = time.time() + 8
        external_url = _services_checker_external_url()
        while time.time() < deadline and _services_checker_ready(external_url):
            time.sleep(0.15)

    return _start_services_checker(), restarted


def _services_checker_external_url():
    return os.getenv(
        "EVENTINSPECTOR_SERVICES_CHECKER_URL",
        SERVICES_CHECKER_EXTERNAL_URL_DEFAULT,
    ).strip().rstrip("/")


def _services_checker_ready(url):
    try:
        response = requests.get(f"{url.rstrip('/')}/", timeout=0.8)
        if not response.ok:
            return False
        body = response.text.lower()
        return any(marker in body for marker in (
            "apk-upload-form",
            "aab-converter-upload-form",
            "services checker",
        ))
    except Exception:
        return False


def _services_checker_uses_shared_keystore(url):
    """Identify the Drive-backed host without trusting a generic root page."""
    try:
        response = requests.get(f"{url.rstrip('/')}/get_keystore_config", timeout=0.8)
        if not response.ok:
            return False
        payload = response.json()
        return bool(
            payload.get("use_hardcoded")
            and payload.get("filename") == SERVICES_CHECKER_KEYSTORE_FILENAME
        )
    except Exception:
        return False


def _find_services_checker_port(preferred=SERVICES_CHECKER_DEFAULT_PORT):
    for port in (preferred,):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
                probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                probe.bind(("127.0.0.1", port))
            return port
        except OSError:
            continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


SERVICES_CHECKER_REQUIRED_RESOURCES = (
    "bundletool-all-1.18.1.jar",
    "my-key.keystore",
)


def _services_checker_resource_roots(app_path):
    """Return the explicit resource roots used by every packaged artifact.

    The layout is intentionally platform-neutral. Build targets may place the
    same ``services_checker`` directory under different artifact roots, but
    resource selection is based only on the declared layout and required files,
    never on an inferred operating system.
    """
    app_path = os.path.abspath(app_path)
    meipass = getattr(sys, "_MEIPASS", "") or ""
    executable = getattr(sys, "executable", "") or ""
    executable_dir = os.path.dirname(os.path.abspath(executable)) if executable else ""
    configured_root = os.getenv("EVENTINSPECTOR_BUNDLED_SERVICES_ROOT", "").strip()

    roots = []
    if configured_root:
        roots.append(os.path.abspath(os.path.expanduser(configured_root)))
    roots.extend([
        os.path.dirname(app_path),
        os.path.join(meipass, "services_checker") if meipass else "",
        meipass,
        os.path.join(executable_dir, "services_checker") if executable_dir else "",
        os.path.join(executable_dir, "_internal", "services_checker") if executable_dir else "",
        os.path.join(executable_dir, "..", "Resources", "services_checker") if executable_dir else "",
        os.path.join(executable_dir, "..", "Frameworks", "services_checker") if executable_dir else "",
        os.path.join(SCRIPT_DIR, "services_checker"),
    ])
    unique_roots = []
    seen = set()
    for root in roots:
        root = os.path.abspath(root) if root else ""
        if not root or root in seen:
            continue
        seen.add(root)
        unique_roots.append(root)
    return unique_roots


def _services_checker_bundled_resource_root(app_path):
    """Find a complete Services Checker resource set for this artifact."""
    for resource_root in _services_checker_resource_roots(app_path):
        if all(
            os.path.isfile(os.path.join(resource_root, filename))
            for filename in SERVICES_CHECKER_REQUIRED_RESOURCES
        ):
            return resource_root
    return ""


def _prepare_services_checker_runtime(app_path):
    """Make bundled dependencies importable before dynamically loading app.py."""
    app_path = os.path.abspath(app_path)
    roots = [os.path.dirname(os.path.dirname(app_path))]
    meipass = getattr(sys, "_MEIPASS", "") or ""
    if meipass:
        roots.insert(0, meipass)
    for root in roots:
        if root and os.path.isdir(root) and root not in sys.path:
            sys.path.insert(0, root)

    # An update may replace Python code without carrying large binary assets.
    # Pin the updated module to the complete resource set from this artifact.
    resource_root = _services_checker_bundled_resource_root(app_path)
    if resource_root:
        os.environ["EVENTINSPECTOR_BUNDLED_SERVICES_ROOT"] = resource_root
        os.environ["EVENTINSPECTOR_BUNDLETOOL_PATH"] = os.path.join(
            resource_root, "bundletool-all-1.18.1.jar"
        )
        logging.debug("Services Checker bundled resource root: %s", resource_root)
    else:
        os.environ.pop("EVENTINSPECTOR_BUNDLED_SERVICES_ROOT", None)
        os.environ.pop("EVENTINSPECTOR_BUNDLETOOL_PATH", None)
        logging.error(
            "Services Checker resource contract is incomplete for artifact: %s",
            app_path,
        )
    try:
        importlib.invalidate_caches()
        importlib.import_module("androguard.core.axml")
    except Exception as exc:
        logging.debug("Services Checker androguard preload unavailable: %s", exc)


def _start_services_checker():
    with _services_checker_lock:
        existing_url = _services_checker_state.get("url")
        if existing_url and _services_checker_ready(existing_url):
            return existing_url

        external_url = _services_checker_external_url()
        external_enabled = _services_checker_external_sources_enabled()
        drive_app_path = next(iter(_services_checker_drive_app_candidates()), "") if external_enabled else ""
        external_ready = _services_checker_ready(external_url) if external_enabled else False
        if external_enabled and external_ready and (
            not drive_app_path or _services_checker_uses_shared_keystore(external_url)
        ):
            _services_checker_state.update({
                "url": external_url,
                "port": None,
                "source": "external",
                "error": "",
            })
            return external_url

        # If port 5000 is occupied by the old bundled service, do not attach
        # to it. Import the Drive app on the private local port below instead.
        # This makes Reload self-healing without requiring a client-side shell.
        force_drive_import = bool(external_enabled and external_ready and drive_app_path)

        command_path = next(
            (path for path in _services_checker_command_candidates() if os.path.isfile(path)),
            "",
        )
        if command_path and not force_drive_import:
            log_path = os.path.join(tempfile.gettempdir(), "eventinspector-services-checker.log")
            try:
                with open(log_path, "ab") as log_handle:
                    launch_args = ["/bin/bash", command_path]
                    launch_kwargs = {
                        "cwd": os.path.dirname(command_path) or None,
                        "stdin": subprocess.DEVNULL,
                        "stdout": log_handle,
                        "stderr": subprocess.STDOUT,
                    }
                    if os.name == "nt":
                        launch_args = ["cmd.exe", "/d", "/c", command_path]
                        launch_kwargs["creationflags"] = getattr(
                            subprocess,
                            "CREATE_NEW_PROCESS_GROUP",
                            0,
                        )
                    else:
                        launch_kwargs["start_new_session"] = True
                    process = subprocess.Popen(
                        launch_args,
                        **launch_kwargs,
                    )
            except Exception as exc:
                raise RuntimeError(f"Cannot start Services Checker command: {exc}") from exc

            deadline = time.time() + 15
            while time.time() < deadline:
                if _services_checker_ready(external_url):
                    _services_checker_state.update({
                        "process": process,
                        "url": external_url,
                        "port": None,
                        "source": command_path,
                        "error": "",
                    })
                    return external_url
                if process.poll() is not None:
                    break
                time.sleep(0.2)

            _services_checker_state["error"] = f"command_not_ready:{command_path}"
            raise RuntimeError(
                f"Services Checker did not become ready at {external_url}; "
                f"see {log_path}"
            )

        app_path = next((path for path in _services_checker_file_candidates() if os.path.exists(path)), "")
        if not app_path:
            raise FileNotFoundError(
                "Services Checker command/app.py was not found. "
                "Set EVENTINSPECTOR_SERVICES_COMMAND or install the bundled checker."
            )

        _prepare_services_checker_runtime(app_path)
        module_name = f"eventinspector_services_checker_{os.getpid()}"
        spec = importlib.util.spec_from_file_location(module_name, app_path)
        if not spec or not spec.loader:
            raise RuntimeError(f"Cannot load Services Checker from {app_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        try:
            spec.loader.exec_module(module)
        except Exception:
            sys.modules.pop(module_name, None)
            raise
        service_app = getattr(module, "app", None)
        if service_app is None:
            raise RuntimeError("Services Checker did not expose a Flask app")

        port = _find_services_checker_port()
        os.environ["SERVICES_CHECKER_HOST"] = "127.0.0.1"
        os.environ["SERVICES_CHECKER_PORT"] = str(port)

        def serve():
            try:
                service_app.run(
                    host="127.0.0.1",
                    port=port,
                    debug=False,
                    use_reloader=False,
                    threaded=True,
                )
            except Exception:
                logging.exception("Services Checker stopped unexpectedly")

        thread = threading.Thread(target=serve, name="services-checker", daemon=True)
        thread.start()
        deadline = time.time() + 12
        service_url = f"http://127.0.0.1:{port}"
        last_error = "service_start_timeout"
        while time.time() < deadline:
            if _services_checker_ready(service_url):
                _services_checker_state.update({
                    "module": module,
                    "thread": thread,
                    "port": port,
                    "url": service_url,
                    "source": app_path,
                    "error": "",
                })
                return service_url
            last_error = "service_root_not_ready"
            time.sleep(0.15)

        _services_checker_state["error"] = last_error
        raise RuntimeError(last_error)

def _resolve_adb():
    adb_env = os.getenv("ADB_PATH")
    if adb_env:
        return adb_env

    # Common SDK locations
    sdk_roots = [
        os.getenv("ANDROID_HOME"),
        os.getenv("ANDROID_SDK_ROOT"),
        os.path.expanduser("~/Library/Android/sdk"),
        os.path.expanduser("~/Android/Sdk"),
    ]
    for root in [r for r in sdk_roots if r]:
        candidate = os.path.join(root, "platform-tools", "adb")
        if os.path.exists(candidate):
            return candidate

    extra_paths = "/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
    current_path = os.environ.get("PATH", "")
    if extra_paths not in current_path:
        os.environ["PATH"] = f"{extra_paths}:{current_path}" if current_path else extra_paths

    found = shutil.which("adb")
    return found or "adb"

ADB_EXECUTABLE = _resolve_adb()

if os.name == 'nt':
    creation_flags = 0x08000000  # CREATE_NO_WINDOW
else:
    creation_flags = 0

# Bảng tra cứu tên thiết bị
DEVICE_NAMES = {
    "R94Y40MZJ5T": "SS A16",
    "RFCX30DPW6D": "SS A35",
    "FMAEAYHIQ8FEYXPN": "OPPO A18",
    "ZY226DFRH2": "Moto Z4",
    "R5CR1282PAK": "SS S21",
    "F6QCCAGIRSQOVGFQ": "Redmi A3",
    "7ec7bca6": "Xiaomi 13 Lite"
}

IOS_DEVICE_NAMES = {
    "eb6b13cc453f5a53ef07ff7149858a8635d18c10": "Iphone X",
    "00008030-000D54412EC0802E": "Iphone 11 pro",
    "00008120-001C0D3C0C61A01E": "Iphone 15",
    "00008020-001D70442113802E": "Ipad Air",
    "00008030-0009398422F3C02E": "Ipad Gen 9",
    "c1ae4200716a97b5b92a377a255d80fc002e7908": "7 Plus white",
    "2327f3aecd8ddc636571c6a8572f87f4fdbcfac0": "7 Plus black",
}

ANDROID_PACKAGE_GAME_CODES = {
    "com.indiez.nonogram": "NG",
    "com.indiez.train.miner": "TM",
    "com.indiez.idletycoon.horse.racing": "HR",
    "com.indiez.solitaire.word.card.puzzle": "SW",
    "com.nostel.dot.line.puzzle": "KN",
    "com.nostel.parking.car": "CP",
    "com.afk.idle.cat.food.restaurent": "CR",
    "tap.monster.block.away": "TP",
}

# --- DỮ LIỆU TOÀN CỤC ---

# Giới hạn cache để UI không giữ quá nhiều log trong RAM.
MAX_LOAD_ADS_LOGS = 1000
MAX_VALIDATOR_LOGS = 1500
MAX_SPECIFIC_EVENT_LOGS = 1500
MAX_CALLBACK_AD_LOGS = 1500
MAX_ADREVENUE_LOGS = 1500
MAX_PRICE_ROTATION_LOGS = 1500

# 1. Dữ liệu cho Tab Load Ads
load_ads_events = deque(maxlen=MAX_LOAD_ADS_LOGS)
unique_load_ads = set()

# 2. Dữ liệu cho Tab Load Ads Ext
load_ads_ext_events = deque(maxlen=MAX_LOAD_ADS_LOGS)
unique_load_ads_ext = set()

# Trạng thái Recording (Google Sheet) - RIÊNG BIỆT CHO TỪNG TAB
recording_states = {
    "LoadAds": {"is_recording": False, "current_sheet": None, "record_request_id": None},
    "LoadAdsExt": {"is_recording": False, "current_sheet": None, "record_request_id": None}
}

# 3. Dữ liệu cho Tab Validator
validator_results = deque(maxlen=MAX_VALIDATOR_LOGS)
# Validator rows not yet sent to the UI, and the per-row counter.
validator_pending = []
validator_seq = 0
required_params = []  # Manual extra params from UI
default_params = []   # Default params from sheet (apply to all events)
event_specific_params = {}  # event_name -> list of params
validator_active = False

# 3A. Dữ liệu coverage cho Tab Default Ad Events
default_ad_event_config = {
    ad_format: {provider: [] for provider in DEFAULT_AD_EVENT_PROVIDERS}
    for ad_format in DEFAULT_AD_EVENT_FORMATS
}
default_ad_event_hits = set()  # (device_id, provider, ad_format, event_name)
default_ad_event_simple_config = {
    category: {provider: [] for provider in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS[category]}
    for category in DEFAULT_AD_EVENT_SIMPLE_CATEGORIES
}
default_ad_event_simple_hits = set()  # (device_id, category, provider, event_name)
default_ad_event_clients = set()

# 4. Dữ liệu cho Tab Specific Event
event_log_cache = deque(maxlen=MAX_SPECIFIC_EVENT_LOGS)
specific_event_results = []
# Specific Validator rows not yet sent to the UI, and the per-event counter.
specific_event_pending = []
specific_event_seq = 0
# Live tables send new rows at most this often (coalesced deltas).
TABLE_DELTA_INTERVAL_SECONDS = 0.25
specific_event_name_filters = []
specific_event_params_filters = []

# 5. Dữ liệu cho Tab Package Logcat
# The server-side cache is only a replay window for (re)connecting clients.
# Live rows reach the UI as deltas ordered by `seq`, so its size no longer
# limits what the UI can show.
PACKAGE_LOG_SERVER_CACHE_ROWS = 50000
package_log_cache = deque(maxlen=PACKAGE_LOG_SERVER_CACHE_ROWS)
PACKAGE_LOG_UI_MAX_ROWS = 8000
PACKAGE_LOG_UI_MAX_ROWS_IOS = 12000
PACKAGE_LOG_DELTA_INTERVAL_SECONDS = 0.2
PACKAGE_LOG_DELTA_CHUNK_ROWS = 2000
# Fields sent to the UI.  The raw record stays on the server (cache + SQLite)
# and is fetched on demand, so each row is not transmitted twice.
PACKAGE_LOG_UI_FIELDS = (
    "seq", "device_id", "device_name", "time", "time_display",
    "level", "tag", "message", "pid", "tid", "is_error",
)
# `seq` increases for every accepted row and never resets while the app runs.
# `package_log_stream_id` changes whenever capture starts, stops or is
# cleared, so a client can tell a new stream from a gap in the current one.
package_log_seq = 0
package_log_stream_id = 0
package_log_last_emitted_seq = 0
PACKAGE_LOG_DB_QUEUE_MAX_ITEMS = 50000
PACKAGE_LOG_DB_BATCH_ROWS = 500
PACKAGE_LOG_SEQ_INDEX_SQL = (
    "CREATE INDEX IF NOT EXISTS idx_package_log_entries_session_seq ON package_log_entries(session_id, seq)"
)
PACKAGE_LOG_SEQ_INDEX_STARTUP_MAX_ROWS = 20000
# History retention: whole sessions older than `days`, then the oldest
# sessions while the database uses more than `max_bytes`.  Nothing is
# deleted automatically until the user has confirmed a cleanup once.
PACKAGE_LOG_RETENTION_FILENAME = "package_log_retention.json"
PACKAGE_LOG_RETENTION_DEFAULTS = {"days": 14, "max_bytes": 2 * 1024 ** 3, "confirmed": False}
PACKAGE_LOG_RETENTION_DELETE_CHUNK_ROWS = 5000
PACKAGE_LOG_MAINTENANCE_FIRST_DELAY_SECONDS = 30
PACKAGE_LOG_MAINTENANCE_INTERVAL_SECONDS = 3600
MAX_IOS_PACKAGE_LOG_BUFFER_CHARS = 500000
MAX_PACKAGE_LOG_RECORD_CHARS = 200000
target_package_name = ""
active_package_pids = {}
active_logcat_processes = {}
active_ios_package_log_buffers = {}
active_ios_package_log_frame_buffers = {}
active_ios_package_log_stream_modes = {}
active_package_log_session_id = None
package_log_db_queue = Queue(maxsize=PACKAGE_LOG_DB_QUEUE_MAX_ITEMS)
package_log_db_dropped_rows = 0
# Last SQLite write error seen by the writer (empty when healthy).
package_log_db_last_error = ""
# Only tests set this; the production writer runs for the life of the app.
package_log_db_writer_stop = threading.Event()

# 6. Dữ liệu cho Tab Callback & Ads Event
callback_ad_logs = deque(maxlen=MAX_CALLBACK_AD_LOGS)

# 7. Dữ liệu cho Tab AdRevenue
adrevenue_log_cache = []
adrevenue_logs = deque(maxlen=MAX_ADREVENUE_LOGS)
# Rendered Revenue rows (mirrors adrevenue_logs), rows not yet sent, counter.
adrevenue_rendered = []
adrevenue_pending = []
adrevenue_seq = 0
adrevenue_default_params = []
adrevenue_source_params = {}

# 8. Dữ liệu cho Tab Price Rotation
price_rotation_logs = deque(maxlen=MAX_PRICE_ROTATION_LOGS)
# Live tables that send new rows as coalesced deltas (see _live_table_append).
LIVE_TABLE_ROWS = {
    "load_ads": load_ads_events,
    "load_ads_ext": load_ads_ext_events,
    "callback": callback_ad_logs,
    "price_rotation": price_rotation_logs,
}
LIVE_TABLE_DELTA_EVENTS = {
    "load_ads": "load_ads_delta",
    "load_ads_ext": "load_ads_ext_delta",
    "callback": "callback_ad_delta",
    "price_rotation": "price_rotation_delta",
}
LIVE_TABLE_PENDING = {name: [] for name in LIVE_TABLE_ROWS}
live_table_seq = 0
PRICE_ROTATION_PREFIX = "[Ad,RewardedBidding"
PRICE_ROTATION_PREFIXES = (
    ("Rewarded", "[Ad,RewardedBidding"),
    ("Interstitial", "[Ad,InterstitialBidding"),
)
PRICE_ROTATION_HIDDEN_TYPES = frozenset({"ascendx", "cloudx"})
PRICE_ROTATION_TYPE_ALIASES = {
    "max": "max",
    "applovin": "max",
    "applovinmax": "max",
    "ascendx": "ascendx",
    "cloudx": "cloudx",
}


def _normalize_price_rotation_type(value):
    token = re.sub(r"[^a-z0-9]+", "", str(value or "").strip().casefold())
    return PRICE_ROTATION_TYPE_ALIASES.get(token, token)

# 9. Dữ liệu cho Tab SDK Check
sdk_check_search_list = []
sdk_check_results = {}
sdk_check_input_list = []
sdk_check_active = False
sdk_check_expected_map = {}
sdk_check_runtime_state = {}
sdk_check_current_network = {}
sdk_check_expected_order = []
sdk_max_ios_pending_lines = {}
sdk_max_ios_core_pending_lines = {}
active_platform = "android"
platform_generation = 0

# Dữ liệu hệ thống chung
active_log_readers = {}
active_ios_log_readers = {}
active_ios_log_processes = {}
active_ios_log_started_at = {}
active_ios_log_last_seen = {}
active_ios_log_commands = {}
ios_device_missing_polls = {}
connected_devices_info = []
installation_id_state = {}
is_paused = False
# Several tab refresh/emit paths legitimately call a helper that takes the
# same runtime lock (for example profile refresh -> default-ad-event payload).
# A plain Lock turns that harmless re-entry into a permanent server-wide
# deadlock, leaving Package Log alive while the page/profile APIs hang.
lock = threading.RLock()
record_dedup_lock = threading.Lock()

# A single physical log record can reach the iOS readers more than once when
# the relay reconnects or when a generic event is dispatched through both the
# callback and event paths.  Keep a bounded exact-record index per tab so the
# UI/cache/export pipeline stores one copy of the same device record.
RECORD_DEDUP_LIMITS = {
    "load_ads": MAX_LOAD_ADS_LOGS * 2,
    "load_ads_ext": MAX_LOAD_ADS_LOGS * 2,
    "validator": MAX_VALIDATOR_LOGS * 2,
    "specific": MAX_SPECIFIC_EVENT_LOGS * 2,
    "callback": MAX_CALLBACK_AD_LOGS * 2,
    "adrevenue": MAX_ADREVENUE_LOGS * 2,
    "price_rotation": MAX_PRICE_ROTATION_LOGS * 2,
    "package": 30000,
}
record_dedup_seen = {name: set() for name in RECORD_DEDUP_LIMITS}
record_dedup_order = {
    name: deque(maxlen=limit)
    for name, limit in RECORD_DEDUP_LIMITS.items()
}


def _accept_exact_record(bucket, device_id, raw_log, *identity_parts):
    """Return False when this exact device record was already accepted.

    Callers may hold ``lock``; this helper uses its own small lock so it is
    safe for the background iOS reader and the socket/UI threads together.
    The raw log includes the source timestamp in normal syslog formats, so
    genuinely separate events remain distinct while replayed identical lines
    collapse to one record.
    """
    if bucket not in record_dedup_seen:
        return True
    key = (
        str(device_id or "").strip(),
        *(str(part or "").strip() for part in identity_parts),
        str(raw_log or "").strip(),
    )
    with record_dedup_lock:
        seen = record_dedup_seen[bucket]
        if key in seen:
            return False
        order = record_dedup_order[bucket]
        if order.maxlen and len(order) >= order.maxlen:
            seen.discard(order[0])
        order.append(key)
        seen.add(key)
        return True


def _clear_record_dedup(bucket=None):
    names = (bucket,) if bucket else tuple(record_dedup_seen)
    with record_dedup_lock:
        for name in names:
            if name not in record_dedup_seen:
                continue
            record_dedup_seen[name].clear()
            record_dedup_order[name].clear()


def _append_callback_ad_row(row):
    """Append and emit a callback row only once for an exact source record."""
    with lock:
        if not _accept_exact_record(
            "callback",
            row.get("device_id", ""),
            row.get("raw_log", ""),
            row.get("type", ""),
            row.get("event_name", ""),
        ):
            return False
        _live_table_append("callback", row)
        snapshot = True
    return True


incomplete_impression_logs = {} # Buffer cho logs bị ngắt dòng
incomplete_ios_adrevenue_logs = {} # Buffer cho iOS AdRevenue logs bị ngắt dòng
incomplete_ios_load_ads_ext_logs = {} # Buffer riêng cho Load Ads đọc AppMetrica AdRevenue iOS
incomplete_ios_max_load_ads_logs = {} # Buffer cho MAX Load Ads viewability logs nhiều dòng
incomplete_adjust_adrevenue_logs = {} # Buffer cho Adjust callback/partner params bị ngắt dòng
incomplete_ios_event_logs = {} # Buffer cho TrackingService event JSON bị xuống dòng trên iOS
incomplete_android_event_logs = {} # Buffer cho TrackingService event JSON bị adb logcat wrap dòng
adb_error_counter = 0
IOS_LOG_STALL_TIMEOUT_SECONDS = 60
IOS_DEVICE_MISSING_GRACE_POLLS = 2
# A reader that delivered data this recently is never stopped just because
# device discovery missed the phone (usbmux on Windows often returns an
# empty list for a moment while the syslog stream keeps working).
IOS_LOG_HEALTHY_SECONDS = 20
# iOS syslog is chatty (~170 lines/s measured), so an alive reader with no
# data for this long while the phone is still connected is stuck: restart it.
IOS_LOG_STALL_RESTART_SECONDS = 30
# While every reader is healthy, run `tidevice list` only this often instead
# of every pass, to keep fewer usbmux connections open at the same time.
IOS_DISCOVERY_SLOW_INTERVAL_SECONDS = 15
ios_discovery_cache = {"at": 0.0, "ids": set()}


def _get_package_db_connection():
    conn = sqlite3.connect(PACKAGE_LOG_DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _init_package_log_db():
    conn = _get_package_db_connection()
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS package_log_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                package_id TEXT NOT NULL,
                started_at REAL NOT NULL,
                ended_at REAL,
                status TEXT NOT NULL DEFAULT 'running'
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS package_log_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                created_at REAL NOT NULL,
                time_display TEXT,
                device_id TEXT,
                device_name TEXT,
                level TEXT,
                tag TEXT,
                message TEXT,
                raw_log TEXT,
                is_error INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY(session_id) REFERENCES package_log_sessions(id)
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_package_log_entries_session ON package_log_entries(session_id, id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_package_log_sessions_started ON package_log_sessions(started_at DESC)")
        # Live rows carry a `seq`; older databases gain the column in place.
        # ADD COLUMN only rewrites the schema, so it stays instant on a large
        # history file.  No index is built here: that would scan the whole
        # table at startup and block the writer.
        entry_columns = {row[1] for row in conn.execute("PRAGMA table_info(package_log_entries)")}
        if "seq" not in entry_columns:
            conn.execute("ALTER TABLE package_log_entries ADD COLUMN seq INTEGER")
        # Sessions keep their own row count so the history list never has to
        # count millions of rows.  Older databases get NULL here and are
        # backfilled once in the background (_backfill_package_log_row_counts).
        session_columns = {row[1] for row in conn.execute("PRAGMA table_info(package_log_sessions)")}
        if "row_count" not in session_columns:
            conn.execute("ALTER TABLE package_log_sessions ADD COLUMN row_count INTEGER")
        # Platform of the capture (android / ios).  Older sessions keep NULL
        # and are resolved from the package id when listed.
        if "platform" not in session_columns:
            conn.execute("ALTER TABLE package_log_sessions ADD COLUMN platform TEXT")
        # A session still marked 'running' at startup was left behind by a
        # crash or force-quit.  Close it at its last saved row.
        conn.execute(
            """
            UPDATE package_log_sessions
            SET status = 'interrupted',
                ended_at = COALESCE(
                    (SELECT e.created_at FROM package_log_entries e
                     WHERE e.session_id = package_log_sessions.id
                     ORDER BY e.id DESC LIMIT 1),
                    started_at
                )
            WHERE status = 'running'
            """
        )
        # The (session_id, seq) index speeds up scroll-back.  Build it now
        # only while the table is small; a large history gets it from
        # Compact so startup never scans the whole file.
        small_table = conn.execute(
            "SELECT COUNT(*) FROM (SELECT 1 FROM package_log_entries LIMIT ?)",
            (PACKAGE_LOG_SEQ_INDEX_STARTUP_MAX_ROWS + 1,),
        ).fetchone()[0] <= PACKAGE_LOG_SEQ_INDEX_STARTUP_MAX_ROWS
        if small_table:
            conn.execute(PACKAGE_LOG_SEQ_INDEX_SQL)
        conn.commit()
    finally:
        conn.close()


def _backfill_package_log_row_counts():
    """Fill row_count for sessions recorded before the column existed.

    One UPDATE computes every missing count under the write lock, so it
    cannot race with the writer's increments: rows written earlier are
    counted here, rows written later are added by the writer.
    """
    conn = _get_package_db_connection()
    try:
        cur = conn.execute(
            """
            UPDATE package_log_sessions
            SET row_count = (SELECT COUNT(*) FROM package_log_entries e WHERE e.session_id = package_log_sessions.id)
            WHERE row_count IS NULL
            """
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def _package_log_retention_path():
    return os.path.join(os.path.dirname(PACKAGE_LOG_DB_PATH), PACKAGE_LOG_RETENTION_FILENAME)


def _load_package_log_retention():
    cfg = dict(PACKAGE_LOG_RETENTION_DEFAULTS)
    try:
        with open(_package_log_retention_path(), "r", encoding="utf-8") as handle:
            stored = json.load(handle)
        if isinstance(stored, dict):
            cfg.update({key: stored[key] for key in PACKAGE_LOG_RETENTION_DEFAULTS if key in stored})
    except FileNotFoundError:
        pass
    except Exception as exc:
        logging.warning("Package Log retention config unreadable, using defaults: %s", exc)
    try:
        cfg["days"] = max(1, int(cfg["days"]))
    except (TypeError, ValueError):
        cfg["days"] = PACKAGE_LOG_RETENTION_DEFAULTS["days"]
    try:
        cfg["max_bytes"] = max(64 * 1024 ** 2, int(cfg["max_bytes"]))
    except (TypeError, ValueError):
        cfg["max_bytes"] = PACKAGE_LOG_RETENTION_DEFAULTS["max_bytes"]
    cfg["confirmed"] = bool(cfg.get("confirmed"))
    return cfg


def _save_package_log_retention(cfg):
    path = _package_log_retention_path()
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump({key: cfg[key] for key in PACKAGE_LOG_RETENTION_DEFAULTS}, handle, indent=2)
    os.replace(tmp_path, path)


def _package_log_db_usage(conn):
    page_size = conn.execute("PRAGMA page_size").fetchone()[0]
    page_count = conn.execute("PRAGMA page_count").fetchone()[0]
    free_pages = conn.execute("PRAGMA freelist_count").fetchone()[0]
    file_bytes = 0
    for suffix in ("", "-wal"):
        try:
            file_bytes += os.path.getsize(PACKAGE_LOG_DB_PATH + suffix)
        except OSError:
            pass
    return {
        "file_bytes": file_bytes,
        "used_bytes": (page_count - free_pages) * page_size,
        "free_bytes": free_pages * page_size,
    }


def _package_log_session_rows(conn):
    return conn.execute(
        """
        SELECT s.id, s.package_id, s.started_at, s.ended_at, s.status,
               CASE WHEN s.row_count IS NULL
                    THEN (SELECT COUNT(*) FROM package_log_entries e WHERE e.session_id = s.id)
                    ELSE s.row_count END AS row_count
        FROM package_log_sessions s
        ORDER BY s.started_at ASC, s.id ASC
        """
    ).fetchall()


def _package_log_retention_plan(conn, cfg, now=None):
    """Sessions the policy would delete, oldest first, with the reason.

    The live session and anything still marked 'running' are never chosen.
    Size is estimated from the average stored row, which is cheap and close
    enough to decide how many old sessions to drop.
    """
    now = time.time() if now is None else now
    with lock:
        protected = {active_package_log_session_id} if active_package_log_session_id else set()
    sessions = _package_log_session_rows(conn)
    usage = _package_log_db_usage(conn)
    total_rows = sum(row["row_count"] or 0 for row in sessions)
    bytes_per_row = (usage["used_bytes"] / total_rows) if total_rows else 0
    cutoff = now - cfg["days"] * 86400
    deletable = [row for row in sessions if row["id"] not in protected and row["status"] != "running"]
    plan = []
    chosen = set()
    for row in deletable:
        if (row["started_at"] or 0) < cutoff:
            plan.append({"id": row["id"], "reason": "age", "rows": row["row_count"] or 0})
            chosen.add(row["id"])
    estimated_used = usage["used_bytes"] - sum(item["rows"] for item in plan) * bytes_per_row
    for row in deletable:
        if estimated_used <= cfg["max_bytes"]:
            break
        if row["id"] in chosen:
            continue
        plan.append({"id": row["id"], "reason": "size", "rows": row["row_count"] or 0})
        chosen.add(row["id"])
        estimated_used -= (row["row_count"] or 0) * bytes_per_row
    by_id = {row["id"]: row for row in sessions}
    for item in plan:
        row = by_id[item["id"]]
        item["package_id"] = row["package_id"]
        item["started_at"] = row["started_at"]
    return {
        "sessions": plan,
        "rows": sum(item["rows"] for item in plan),
        "estimated_bytes": int(sum(item["rows"] for item in plan) * bytes_per_row),
        "usage": usage,
        "total_sessions": len(sessions),
        "total_rows": total_rows,
    }


def _delete_package_log_session(conn, session_id):
    """Delete one session in small transactions so the writer is never blocked long."""
    deleted = 0
    while True:
        cur = conn.execute(
            "DELETE FROM package_log_entries WHERE id IN "
            "(SELECT id FROM package_log_entries WHERE session_id = ? LIMIT ?)",
            (session_id, PACKAGE_LOG_RETENTION_DELETE_CHUNK_ROWS),
        )
        conn.commit()
        deleted += cur.rowcount
        if cur.rowcount < PACKAGE_LOG_RETENTION_DELETE_CHUNK_ROWS:
            break
    conn.execute("DELETE FROM package_log_sessions WHERE id = ?", (session_id,))
    conn.commit()
    return deleted


def _apply_package_log_retention(confirmed_by_user=False, now=None):
    """Run the retention policy.  The first run needs the user's confirmation."""
    cfg = _load_package_log_retention()
    if not cfg["confirmed"] and not confirmed_by_user:
        return {"ok": False, "error": "confirmation_required"}
    if confirmed_by_user and not cfg["confirmed"]:
        cfg["confirmed"] = True
        _save_package_log_retention(cfg)
    conn = _get_package_db_connection()
    try:
        plan = _package_log_retention_plan(conn, cfg, now=now)
        deleted_sessions = 0
        deleted_rows = 0
        for item in plan["sessions"]:
            with lock:
                if item["id"] == active_package_log_session_id:
                    continue
            deleted_rows += _delete_package_log_session(conn, item["id"])
            deleted_sessions += 1
        usage = _package_log_db_usage(conn)
    finally:
        conn.close()
    if deleted_sessions:
        logging.info("Package Log retention removed %s sessions / %s rows", deleted_sessions, deleted_rows)
    return {
        "ok": True,
        "deleted_sessions": deleted_sessions,
        "deleted_rows": deleted_rows,
        "usage": usage,
        "policy": {key: cfg[key] for key in ("days", "max_bytes", "confirmed")},
    }


def _package_log_maintenance_loop():
    """Background upkeep: backfill row counts, then apply retention hourly
    (only after the user has confirmed a cleanup once)."""
    try:
        _backfill_package_log_row_counts()
    except Exception as exc:
        logging.warning("Package Log row-count backfill failed: %s", exc)
    time.sleep(PACKAGE_LOG_MAINTENANCE_FIRST_DELAY_SECONDS)
    while True:
        try:
            if _load_package_log_retention()["confirmed"]:
                _apply_package_log_retention()
        except Exception as exc:
            logging.warning("Package Log retention pass failed: %s", exc)
        time.sleep(PACKAGE_LOG_MAINTENANCE_INTERVAL_SECONDS)


def _start_package_log_session(package_id):
    conn = _get_package_db_connection()
    try:
        cur = conn.execute(
            "INSERT INTO package_log_sessions (package_id, started_at, status, row_count, platform) VALUES (?, ?, 'running', 0, ?)",
            (package_id, time.time(), "ios" if active_platform == "ios" else "android"),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def _finish_package_log_session(session_id):
    if not session_id:
        return
    conn = _get_package_db_connection()
    try:
        conn.execute(
            "UPDATE package_log_sessions SET ended_at = ?, status = 'stopped' WHERE id = ?",
            (time.time(), session_id),
        )
        conn.commit()
    finally:
        conn.close()


def _enqueue_package_log_db_row(item):
    """Put one Package Log row on a bounded queue.

    The live table is already bounded, but the SQLite writer used to have an
    unbounded queue.  A reconnect or a slow/locked database could therefore
    retain every raw log string in RAM even after the UI had been cleared.
    Keep the newest rows and drop the oldest pending write when the writer
    falls behind; rows already committed to SQLite remain available in the
    history view.
    """
    global package_log_db_dropped_rows
    try:
        package_log_db_queue.put_nowait(item)
        return True
    except Full:
        try:
            package_log_db_queue.get_nowait()
            package_log_db_dropped_rows += 1
        except Empty:
            pass
        try:
            package_log_db_queue.put_nowait(item)
            return True
        except Full:
            package_log_db_dropped_rows += 1
            return False


def _drain_package_log_db_queue():
    """Discard writes that belong to a capture the user just stopped/cleared."""
    drained = 0
    while True:
        try:
            package_log_db_queue.get_nowait()
            drained += 1
        except Empty:
            return drained


def _package_log_db_writer():
    """Persist queued Package Log rows; never let one SQLite error stop it.

    A failed write (locked database during Clear History, disk full, ...)
    used to end this thread silently, after which every new row was dropped.
    Now the batch is kept and retried with a fresh connection, and the error
    is exposed through ``package_log_db_last_error`` for the UI.
    """
    global package_log_db_last_error, package_log_db_dropped_rows
    conn = None
    batch = []
    last_flush = time.time()
    retry_delay = 0.5

    def flush():
        nonlocal conn, batch, last_flush
        if not batch:
            return
        if conn is None:
            conn = _get_package_db_connection()
        conn.executemany(
            """
            INSERT INTO package_log_entries
            (session_id, created_at, time_display, device_id, device_name, level, tag, message, raw_log, is_error, seq)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            batch,
        )
        per_session = {}
        for item in batch:
            per_session[item[0]] = per_session.get(item[0], 0) + 1
        conn.executemany(
            "UPDATE package_log_sessions SET row_count = row_count + ? WHERE id = ?",
            [(count, session_id) for session_id, count in per_session.items()],
        )
        conn.commit()
        batch = []
        last_flush = time.time()

    while not package_log_db_writer_stop.is_set() or batch:
        if package_log_db_writer_stop.is_set() and conn is None and package_log_db_last_error:
            break
        try:
            item = package_log_db_queue.get(timeout=0.5)
            batch.append(item)
            # Take whatever else is already waiting, up to one batch.
            while len(batch) < PACKAGE_LOG_DB_BATCH_ROWS:
                try:
                    batch.append(package_log_db_queue.get_nowait())
                except Empty:
                    break
        except Empty:
            pass
        if not batch or (
            len(batch) < PACKAGE_LOG_DB_BATCH_ROWS
            and time.time() - last_flush < 0.5
            and not package_log_db_writer_stop.is_set()
        ):
            continue
        try:
            flush()
            if package_log_db_last_error:
                package_log_db_last_error = ""
            retry_delay = 0.5
        except Exception as exc:
            package_log_db_last_error = f"{type(exc).__name__}: {exc}"
            logging.warning("Package Log DB write failed, will retry: %s", package_log_db_last_error)
            try:
                if conn is not None:
                    conn.rollback()
                    conn.close()
            except Exception:
                pass
            conn = None
            # Bound memory while the database stays unavailable: keep the
            # newest rows and count every row given up.
            overflow = len(batch) - PACKAGE_LOG_DB_QUEUE_MAX_ITEMS
            if overflow > 0:
                del batch[:overflow]
                package_log_db_dropped_rows += overflow
            time.sleep(retry_delay)
            retry_delay = min(retry_delay * 2, 5.0)
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass


# --- REGEX PATTERNS ---

# Pattern cho Load Ads (Unity)
UNITY_TRACKING_PATTERN = re.compile(r'\[\s*Tracking\s*\]\s*TrackingService->Track:\s*(\{"eventName":"ad_impression".*)')

# Pattern cho Load Ads MAX (AppLovin MAX revenue callback).  MAX writes the
# ad format and network name inside MediatedAd; keep the marker exact so
# unrelated AppLovin diagnostics never become Load Ads rows.
MAX_LOAD_ADS_REVENUE_KEYWORD = "MaxAdRevenueListener.onAdRevenuePaid"
MAX_LOAD_ADS_FORMAT_PATTERN = re.compile(r"\bformat\s*=\s*([A-Za-z0-9_-]+)", re.IGNORECASE)
MAX_LOAD_ADS_NETWORK_PATTERN = re.compile(r"\bnetworkName\s*=\s*(['\"])(.*?)\1", re.IGNORECASE)
MAX_LOAD_ADS_IOS_DELEGATE_KEYWORDS = (
    "didDisplayAd:",
    "didPayRevenueForAd:",
)
MAX_LOAD_ADS_IOS_DELEGATE_NETWORK_PATTERN = re.compile(
    r"\bnetworkName\s*=\s*(?:'([^']+)'|\"([^\"]+)\"|([^,\]\r\n]+))",
    re.IGNORECASE,
)
MAX_LOAD_ADS_VIEWABILITY_KEYWORD = "Reporting mediated_ad_viewability_impression_called"
MAX_LOAD_ADS_IOS_FORMAT_PATTERN = re.compile(
    r"[\"']?ad_format[\"']?\s*(?:=|:)\s*(?:\"([^\"]+)\"|'([^']+)'|([A-Za-z0-9_-]+))",
    re.IGNORECASE,
)
MAX_LOAD_ADS_IOS_NETWORK_PATTERN = re.compile(
    r"[\"']?network_name[\"']?\s*(?:=|:)\s*(?:\"([^\"]+)\"|'([^']+)'|([^;\s}]+))",
    re.IGNORECASE,
)
MAX_LOAD_ADS_IOS_ADAPTER_CLASS_PATTERN = re.compile(
    r"[\"']?adapter_class[\"']?\s*(?:=|:)\s*(?:\"([^\"]+)\"|'([^']+)'|([^;\s}]+))",
    re.IGNORECASE,
)
MAX_LOAD_ADS_IOS_ADAPTER_VERSION_PATTERN = re.compile(
    r"[\"']?adapter_version[\"']?\s*(?:=|:)\s*(?:\"([^\"]+)\"|'([^']+)'|([^;\s}]+))",
    re.IGNORECASE,
)

# Pattern cho Load Ads Ext (AppMetrica)
METRICA_TRACKING_PATTERN = re.compile(r'Event sent: ad_impression with value\s*(\{.*\})')
LOAD_ADS_EXT_ADREVENUE_PATTERN = re.compile(r'AdRevenue Received:\s*AdRevenue\{(.*)\}', re.IGNORECASE)
IOS_LOAD_ADS_EXT_ADREVENUE_KEYWORDS = (
    "AppMetricaTrackingHandler->_Handle:",
    "AppMetricaAdRevenueTrackingHandler->Handle:",
    "Appmetrica TrackAdRevenueEventForOther:",
    "AppMetrica TrackAdRevenueEventForOther:",
    "Appmetrica TrackAdRevenueEvent:",
    "AppMetrica TrackAdRevenueEvent:",
)
METRICA_REGULAR_EVENT_PATTERN = re.compile(
    r'Event received on service:\s*EVENT_TYPE_REGULAR\s+with name\s+([A-Za-z0-9_.$-]+)\s+with value\s*(\{.*\})'
)
TRACKING_SERVICE_EVENT_PATTERN = re.compile(
    r'TrackingService\s*->\s*(?:Track|_LogEvent)\s*:',
    re.IGNORECASE,
)
ANDROID_LOG_RECORD_START_PATTERN = re.compile(
    r'^\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?\s+',
)

# Patterns cũ của Log Checker
OLD_EVENT_LOG_PATTERN = re.compile(r'\[\s*Tracking\s*\]\s*TrackingService->Track:\s*(\{"eventName":.*)')
IOS_FIREBASE_EVENT_PATTERN = re.compile(
    r'\[\s*Tracking\s*\]\s*TrackingService->_LogEvent:\s*(\{.*\})',
    re.IGNORECASE
)
LEVELPLAY_IMPRESSION_DATA_MAPPER_KEYWORD = "LevelPlayAdImpressionEventMapper->Map"
ASCENDX_CALLBACK_EVENTS = (
    "_OnAdLoadedEvent",
    "_OnAdImpressionEvent",
    "_OnAdDisplayedEvent",
    "_OnAdClickedEvent",
    "_OnAdClosedEvent",
    "_OnAdReceivedRewardEvent",
)
CLOUDX_CALLBACK_EVENTS = (
    "OnAdLoadFailed",
    "OnAdLoadSuccess",
    "OnAdRevenuePaid",
    "OnAdClicked",
    "OnAdShowSuccess",
    "OnAdRewarded",
    "OnAdHidden",
)
CALLBACK_LOG_PATTERN = re.compile(
    r"(LevelPlayAdImpressionEventMapper->Map|_OnImpressionDataReadyEvent|"
    r"_OnLevelPlayImpressionDataReadyEvent|LevelPlayInterstitialAdListener|"
    r"LevelPlayBannerAdViewListener|LevelPlayRewardedAdListener|"
    r"Receive Ironsource Impression Data LevelPlayImpressionData|"
    r"_OnAdLoadedEvent|_OnAdImpressionEvent|_OnAdDisplayedEvent|"
    r"_OnAdClickedEvent|_OnAdClosedEvent|_OnAdReceivedRewardEvent|"
    r"OnAdLoadFailed|OnAdLoadSuccess|OnAdRevenuePaid|OnAdClicked|"
    r"OnAdShowSuccess|OnAdRewarded|OnAdHidden)"
)
ADREVENUE_LOG_PATTERN = re.compile(r"AdRevenue Received:\s*AdRevenue\{(.*)\}")
APPSFLYER_ADREVENUE_PATTERN = re.compile(r"\b(ADREVENUE)-\d+:\s*preparing data:\s*(\{.*\})", re.IGNORECASE)
SDK_CHECK_SEARCH_PATTERN = re.compile(r'"search_pattern"\s*:\s*["\'](.*?)["\']')
GADSME_SERVICE_KEYWORD = "[InPlayAds,Gadsme]"
ADVERTY5_KEYWORD = "[InPlayAds,Adverty]"
SDK_HEADER_PATTERN = re.compile(r'-{5,}\s*(.*?)\s*-{5,}')
SDK_VERSION_LINE_PATTERN = re.compile(r'SDK\s+Version\s*[-–—]\s*(.+)$', re.IGNORECASE)
SDK_ADAPTER_VERSION_LINE_PATTERN = re.compile(r'Adapter\s+Version\s*[-–—]\s*(.+)$', re.IGNORECASE)
IOS_SDK_VERSION_LINE_PATTERN = re.compile(r'SDK\s*[-–—]\s*Version\s+([0-9]+(?:\.[0-9]+)+)', re.IGNORECASE)
IOS_ADAPTER_VERSION_LINE_PATTERN = re.compile(r'Adapter\s*[-–—]\s*Version\s+([0-9]+(?:\.[0-9]+)+)', re.IGNORECASE)
SDK_ADAPTER_MISSING_PATTERN = re.compile(r'Adapter\s*[-–—]\s*MISSING\b', re.IGNORECASE)
IOS_ADAPTER_MISSING_PATTERN = re.compile(r'Adapter\s+\*+\s*MISSING\s*\*+', re.IGNORECASE)
SDK_VERIFICATION_PATTERN = re.compile(r'>{3,}\s*(.*?)\s*-\s*(VERIFIED|NOT VERIFIED)\b', re.IGNORECASE)
IOS_VERIFICATION_PATTERN = re.compile(r'IntegrationHelper\s+(.+?)\s+\*+\s*(NOT VERIFIED|VERIFIED)\s*\*+', re.IGNORECASE)
SDK_CLOUDX_ADAPTER_METADATA_PATTERN = re.compile(
    r'\[AdapterMetadataResolver\]\s+Discovered adapter metadata:\s*'
    r'network\s*=\s*([^,]+?)\s*,\s*'
    r'adapterVersion\s*=\s*([0-9]+(?:\.[0-9]+)+)',
    re.IGNORECASE,
)
SDK_CLOUDX_NATIVE_METADATA_PATTERN = re.compile(
    r'\bnetworkSdkVersion\s*=\s*([0-9]+(?:\.[0-9]+)+)',
    re.IGNORECASE,
)
SDK_CLOUDX_HTTP_METADATA_PATTERN = re.compile(r'\[CloudXHttpClient\]', re.IGNORECASE)
AUDIOMOB_OPEN_MEASUREMENT_SDK_PATTERN = re.compile(
    r'\bAudiomob\b.*?\bOpen\s+Measurement\s+started\s*'
    r'\(\s*SDK\s*:\s*v?(?P<version>[0-9]+(?:\.[0-9]+)+)\s*\)',
    re.IGNORECASE,
)
SDK_MAX_CORE_INITIALIZATION_PATTERN = re.compile(
    r'\[TaskInitializeSdk\]\s+AppLovin\s+SDK\s+'
    r'(?P<version>[0-9]+(?:\.[0-9]+)+)\s+initialization\s+succeeded\b',
    re.IGNORECASE,
)
SDK_MAX_CORE_VERSION_PATTERN = re.compile(
    r'\bAppLovin\s+SDK\b.*?\bVersion\s*:\s*'
    r'(?P<version>[0-9]+(?:\.[0-9]+)+)',
    re.IGNORECASE | re.DOTALL,
)
SDK_MAX_IOS_CORE_MARKER_PATTERN = re.compile(
    r'={3,}\s*AppLovin\s+SDK\s*={3,}',
    re.IGNORECASE,
)
SDK_MAX_HEALTH_EVENTS_PATTERN = re.compile(
    r'\[(?:AL)?HealthEventsReporter\].*?\bsignal_collection_success\b',
    re.IGNORECASE,
)
SDK_MAX_ADAPTER_VERSION_FIELD_PATTERN = re.compile(
    r'["\']?\badapter_version["\']?\s*(?:=|:)\s*["\']?'
    r'(?P<version>[0-9]+(?:\.[0-9]+)+)',
    re.IGNORECASE,
)
SDK_MAX_NETWORK_NAME_FIELD_PATTERN = re.compile(
    r'["\']?\bnetwork_name["\']?\s*(?:=|:)\s*["\']?'
    r'(?P<network>[A-Za-z0-9_./() -]+?)["\']?\s*(?:[,;\]}])',
    re.IGNORECASE,
)
SDK_MAX_IOS_RECORD_START_PATTERN = re.compile(
    r'^(?:'
    r'[A-Z][a-z]{2}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?'
    r'|\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?'
    r'|\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?'
    r')\s+',
)

# Mapping tên hiển thị cho Callback
CALLBACK_DISPLAY_NAMES = {
    "LevelPlayInterstitialAdListener": "Interstitial",
    "LevelPlayBannerAdViewListener": "Banner",
    "LevelPlayRewardedAdListener": "Rewarded",
    "_OnImpressionDataReadyEvent": "Impression Data",
    "_OnLevelPlayImpressionDataReadyEvent": "LevelPlay Impression Data",
    "Receive Ironsource Impression Data LevelPlayImpressionData": "LevelPlayImpressionData",
    LEVELPLAY_IMPRESSION_DATA_MAPPER_KEYWORD: "LevelPlayImpressionData",
}

IOS_LEVELPLAY_CALLBACK_KEYS = {
    "[Ad,LevelPlay,Reward] LevelPlayRewardAdWrapper->_OnAdRewarded": "Reward _OnAdRewarded",
    "[Ad,LevelPlay,Reward] LevelPlayRewardAdWrapper->_OnAdClosed": "Reward _OnAdClosed",
    "[Ad,LevelPlay,Reward] LevelPlayRewardAdWrapper->_OnAdLoaded": "Reward _OnAdLoaded",
    "[Ad,LevelPlay,Reward] LevelPlayRewardAdWrapper->_OnAdClicked": "Reward _OnAdClicked",
    "[Ad,LevelPlay,Reward] LevelPlayRewardAdWrapper->OnADAdLoadFailed": "Reward OnADAdLoadFailed",
    "[Ad,LevelPlay,Interstitial] LevelPlayInterstitialAdWrapper->_OnAdLoaded": "Interstitial _OnAdLoaded",
    "[Ad,LevelPlay,Interstitial] LevelPlayInterstitialAdWrapper->_OnAdInfoChanged": "Interstitial _OnAdInfoChanged",
    "[Ad,LevelPlay,Interstitial] LevelPlayInterstitialAdWrapper->_OnAdDisplayed": "Interstitial _OnAdDisplayed",
    "[Ad,LevelPlay,Interstitial] LevelPlayInterstitialAdWrapper->_OnAdClosed": "Interstitial _OnAdClosed",
    "[Ad,LevelPlay,Interstitial] LevelPlayInterstitialAdWrapper->_OnAdClicked": "Interstitial _OnAdClicked",
    "[Ad,LevelPlay,Interstitial] LevelPlayInterstitialAdWrapper->OnADAdLoadFailed": "Interstitial OnADAdLoadFailed",
    "[Ad,LevelPlay,Banner] LevelPlayBannerAdService->_OnImpressionDataReady": "Banner _OnImpressionDataReady",
    "[Ad,LevelPlay,Banner] LevelPlayBannerWrapper->OnADAdDisplayed": "Banner OnADAdDisplayed",
    "[Ad,LevelPlay,Banner] LevelPlayBannerWrapper->OnADAdLoaded": "Banner OnADAdLoaded",
    "[Ad,LevelPlay,Banner] LevelPlayBannerWrapper->OnADAdClicked": "Banner OnADAdClicked",
    "[Ad,LevelPlay,Banner] LevelPlayBannerWrapper->OnADAdLoadFailed": "Banner OnADAdLoadFailed",
    "[Ad,LevelPlay,Mrec] LevelPlayMrecWrapper->_OnAdLoaded": "Mrec _OnAdLoaded",
    "[Ad,LevelPlay,Mrec,Aps] LevelPlayMrecAdService->_OnImpressionDataReady": "Mrec APS _OnImpressionDataReady",
    "[Ad,LevelPlay,Mrec] LevelPlayMrecWrapper->_OnAdDisplayed": "Mrec _OnAdDisplayed",
}

def get_device_name(device_id):
    return DEVICE_NAMES.get(device_id) or IOS_DEVICE_NAMES.get(device_id) or device_id

def get_ios_device_name(device_id):
    return IOS_DEVICE_NAMES.get(device_id) or "iOS Device"

def _short_device_id(device_id, length=8):
    if not device_id:
        return ""
    return f"{device_id[:length]}..."


def _game_code_for_package(package_id):
    package_id = (package_id or "").strip()
    return ANDROID_PACKAGE_GAME_CODES.get(package_id) or "Unknown"


def _get_android_foreground_package(device_id):
    commands = [
        [ADB_EXECUTABLE, "-s", device_id, "shell", "dumpsys", "window", "windows"],
        [ADB_EXECUTABLE, "-s", device_id, "shell", "dumpsys", "activity", "activities"],
    ]
    patterns = [
        r"mCurrentFocus=.*? ([A-Za-z0-9_\.]+)/[A-Za-z0-9_\.]+",
        r"mFocusedApp=.*? ([A-Za-z0-9_\.]+)/[A-Za-z0-9_\.]+",
        r"topResumedActivity:.*? ([A-Za-z0-9_\.]+)/[A-Za-z0-9_\.]+",
        r"ResumedActivity:.*? ([A-Za-z0-9_\.]+)/[A-Za-z0-9_\.]+",
        r"mResumedActivity:.*? ([A-Za-z0-9_\.]+)/[A-Za-z0-9_\.]+",
    ]
    for cmd in commands:
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=creation_flags,
                timeout=5,
            )
            output = result.stdout or ""
        except Exception:
            continue
        for pattern in patterns:
            match = re.search(pattern, output)
            if match:
                return match.group(1).strip()
    return ""


def _build_installation_id_payload(device_id):
    state = installation_id_state.get(device_id) or {}
    package_id = state.get("package_id", "")
    return {
        "device_id": device_id,
        "device_name": get_device_name(device_id),
        "platform": active_platform,
        "package_id": package_id,
        "game_code": state.get("game_code") or _game_code_for_package(package_id),
        "installation_id": state.get("installation_id", ""),
    }


def _emit_installation_id_state(device_id=None):
    with lock:
        if device_id:
            payload = _build_installation_id_payload(device_id)
        else:
            target_device_id = ""
            if installation_id_state:
                target_device_id = max(
                    installation_id_state.items(),
                    key=lambda item: item[1].get("updated_at", 0),
                )[0]
            elif connected_devices_info:
                target_device_id = connected_devices_info[0].get("id", "")
            payload = _build_installation_id_payload(target_device_id)
    socketio.emit("installation_id_status", payload)


def _update_installation_id_runtime(device_id, package_id="", installation_id="", force_emit=False):
    package_id = (package_id or "").strip()
    game_code = _game_code_for_package(package_id)
    changed = False
    with lock:
        state = dict(installation_id_state.get(device_id) or {})
        package_changed = state.get("package_id", "") != package_id
        if package_changed:
            state["package_id"] = package_id
            changed = True
        if state.get("game_code", "") != game_code:
            state["game_code"] = game_code
            changed = True
        if package_changed and state.get("installation_id", ""):
            state["installation_id"] = ""
            changed = True
        if installation_id and state.get("installation_id", "") != installation_id:
            state["installation_id"] = installation_id
            changed = True
        if changed or force_emit:
            state["updated_at"] = time.time()
            installation_id_state[device_id] = state
    if changed or force_emit:
        _emit_installation_id_state(device_id)


def process_installation_id_log(line, device_id):
    if active_platform != "android":
        return
    installation_id = ""
    marker = "FirebaseInstallationIdPostInitHandler->_DebugInstallationId:"
    if marker in line:
        _, _, json_part = line.partition(marker)
        json_part = json_part.strip()
        if json_part:
            try:
                payload = json.loads(json_part)
            except Exception:
                payload = {}
            id_task = payload.get("idTask") or {}
            installation_id = str(id_task.get("idTask.Result") or "").strip()
    if not installation_id:
        plain_match = re.search(r"\bInstallations id\s+([A-Za-z0-9_-]+)", line)
        if plain_match:
            installation_id = plain_match.group(1).strip()
    if not installation_id:
        return
    package_id = _get_android_foreground_package(device_id)
    _update_installation_id_runtime(device_id, package_id=package_id, installation_id=installation_id, force_emit=True)


def android_installation_id_monitor():
    while True:
        time.sleep(2)
        if active_platform != "android":
            continue
        with lock:
            devices_snapshot = list(connected_devices_info)
        for device in devices_snapshot:
            device_id = device.get("id", "")
            if not device_id:
                continue
            try:
                package_id = _get_android_foreground_package(device_id)
            except Exception:
                continue
            if not package_id:
                continue
            _update_installation_id_runtime(device_id, package_id=package_id)

def _ios_log_reader_status(device_id):
    now = time.time()
    thread = active_ios_log_readers.get(device_id)
    proc = active_ios_log_processes.get(device_id)
    started_at = active_ios_log_started_at.get(device_id)
    last_seen = active_ios_log_last_seen.get(device_id)

    if not thread:
        return "log reader: starting"
    if thread and not thread.is_alive():
        return "log reader: stopped"
    if proc and proc.poll() is not None:
        return f"log reader: exited {proc.poll()}"
    if last_seen and now - last_seen <= 10:
        return "log OK"
    if started_at and now - started_at <= IOS_LOG_STALL_TIMEOUT_SECONDS:
        return "waiting logs"
    if last_seen:
        return f"no logs {int(now - last_seen)}s"
    return "waiting logs"

def _ios_log_reader_color_class(status):
    status_text = (status or "").lower()
    if "log ok" in status_text:
        return "text-green-600"
    if "waiting" in status_text or "starting" in status_text or "restarting" in status_text:
        return "text-orange-500"
    return "text-red-600"

def _make_device_info(device_id, platform):
    if platform == "ios":
        name = get_ios_device_name(device_id)
        log_status = _ios_log_reader_status(device_id)
        return {
            "id": device_id,
            "name": name,
            "platform": "ios",
            "log_status": log_status,
            "status_class": _ios_log_reader_color_class(log_status),
            "display_name": f"{name} ({_short_device_id(device_id)}) - {log_status}",
        }
    return {
        "id": device_id,
        "name": get_device_name(device_id),
        "platform": "android",
        "display_name": f"{device_id} - {get_device_name(device_id)}",
    }

def _resolve_ios_tool(name):
    candidates = [
        shutil.which(name),
        os.path.join(os.path.expanduser("~/homebrew/bin"), name),
        os.path.join(os.path.expanduser("~/.homebrew/bin"), name),
        os.path.join(os.path.expanduser("~/Homebrew/bin"), name),
        os.path.join(os.path.expanduser("~/Library/Python/3.12/bin"), name),
        os.path.join(os.path.expanduser("~/Library/Python/3.11/bin"), name),
        os.path.join(os.path.expanduser("~/Library/Python/3.10/bin"), name),
        os.path.join("/opt/homebrew/bin", name),
        os.path.join("/usr/local/bin", name),
        os.path.join("/usr/bin", name),
        os.path.join("/usr/sbin", name),
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return None

def _normalize_ios_udid(value):
    if not value:
        return ""
    cleaned = re.sub(r'[^0-9A-Fa-f]', '', str(value)).upper()
    if len(cleaned) == 24:
        return f"{cleaned[:8]}-{cleaned[8:]}"
    if len(cleaned) == 40:
        return cleaned.lower()
    return str(value).strip()

def _list_ios_device_ids():
    ids = []
    tidevice = _resolve_ios_tool("tidevice")
    if tidevice:
        try:
            output = subprocess.run(
                [tidevice, "list"],
                capture_output=True,
                text=True,
                timeout=8,
                creationflags=creation_flags,
            ).stdout
            for line in output.splitlines():
                match = re.search(r'\b([0-9A-Fa-f]{8}-[0-9A-Fa-f]{16}|[0-9A-Fa-f]{24}|[0-9A-Fa-f]{40})\b', line)
                if not match:
                    continue
                if "ConnectionType.USB" in line:
                    ids.append(_normalize_ios_udid(match.group(1)))
        except Exception:
            pass

        # This application intentionally supports iOS over USB only. A
        # network-paired row from tidevice must never create a reader or show
        # up as a connected device.
        seen = set()
        unique_ids = []
        for device_id in ids:
            if device_id not in seen:
                seen.add(device_id)
                unique_ids.append(device_id)
        return unique_ids

    # On macOS, use the USB inventory before any generic libimobiledevice
    # fallback. This keeps network-paired devices out even when tidevice is
    # not installed. Do not fall through to xcrun: it can include network
    # paired devices in its device list.
    system_profiler = _resolve_ios_tool("system_profiler")
    if system_profiler:
        try:
            output = subprocess.run(
                [system_profiler, "SPUSBDataType"],
                capture_output=True,
                text=True,
                timeout=8,
                creationflags=creation_flags,
            ).stdout
            in_ios_block = False
            for line in output.splitlines():
                stripped = line.strip()
                if stripped in {"iPhone:", "iPad:", "iPod:"}:
                    in_ios_block = True
                    continue
                if in_ios_block and stripped.startswith("Serial Number:"):
                    ids.append(_normalize_ios_udid(stripped.split(":", 1)[1].strip()))
                    in_ios_block = False
                elif in_ios_block and stripped.endswith(":") and stripped not in {"iPhone:", "iPad:", "iPod:"}:
                    in_ios_block = False
        except Exception:
            pass

        seen = set()
        unique_ids = []
        for device_id in ids:
            if device_id not in seen:
                seen.add(device_id)
                unique_ids.append(device_id)
        return unique_ids

    idevice_id = _resolve_ios_tool("idevice_id")
    if idevice_id:
        try:
            output = subprocess.run(
                [idevice_id, "-l"],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=creation_flags,
            ).stdout
            ids.extend([_normalize_ios_udid(line.strip()) for line in output.splitlines() if line.strip()])
        except Exception:
            pass

    # Without `tidevice`, `idevice_id` is still our best live-device fallback.
    if idevice_id:
        seen = set()
        unique_ids = []
        for device_id in ids:
            if device_id not in seen:
                seen.add(device_id)
                unique_ids.append(device_id)
        return unique_ids

    seen = set()
    unique_ids = []
    for device_id in ids:
        if device_id not in seen:
            seen.add(device_id)
            unique_ids.append(device_id)
    return unique_ids


def _ios_reader_recently_active(device_id, now, window=IOS_LOG_HEALTHY_SECONDS):
    """True when the reader process is alive and delivered data within `window` seconds."""
    proc = active_ios_log_processes.get(device_id)
    last_seen = active_ios_log_last_seen.get(device_id)
    if proc is None or proc.poll() is not None or not last_seen:
        return False
    return now - last_seen < window


def _observe_ios_device_ids(now=None):
    """Device discovery, skipped (cached) while every reader is streaming."""
    now = time.time() if now is None else now
    with lock:
        readers = set(active_ios_log_readers)
        healthy = bool(readers) and all(_ios_reader_recently_active(did, now, window=10) for did in readers)
        cached = set(ios_discovery_cache["ids"])
        fresh = now - ios_discovery_cache["at"] < IOS_DISCOVERY_SLOW_INTERVAL_SECONDS
    if healthy and fresh and readers <= cached:
        return cached
    ids = set(_list_ios_device_ids())
    with lock:
        ios_discovery_cache["at"] = now
        ios_discovery_cache["ids"] = set(ids)
    return ids


def _ios_readers_to_stop(observed_ids, now):
    """Decide which iOS readers to stop. Call under `lock`.

    Returns (readers_to_stop as [(device_id, reason)], stale_reader_ids).
    - dead: the thread or process exited;
    - missing from discovery: absent beyond the grace polls AND no data for
      IOS_LOG_HEALTHY_SECONDS (a streaming reader survives discovery misses);
    - stalled: the phone is still connected but the alive reader sent nothing
      for IOS_LOG_STALL_RESTART_SECONDS; the device manager starts a new one.
    """
    readers_to_stop = []
    active_reader_ids = set(active_ios_log_readers) | set(active_ios_log_processes)
    stale_reader_ids = {
        did for did in (active_reader_ids - observed_ids)
        if ios_device_missing_polls.get(did, 0) > IOS_DEVICE_MISSING_GRACE_POLLS
        and not _ios_reader_recently_active(did, now)
    }
    for did in sorted(stale_reader_ids):
        readers_to_stop.append((did, 'missing from discovery'))
    # A reader can disappear from the thread map while its process is still
    # alive. Treat that process as orphaned; otherwise the next pass may
    # start a second stream for the same device and duplicate the old buffer.
    for did in sorted(set(active_ios_log_processes) - set(active_ios_log_readers)):
        if did not in stale_reader_ids:
            readers_to_stop.append((did, 'orphan process'))
    for did in list(active_ios_log_readers.keys()):
        if did in stale_reader_ids:
            continue
        proc = active_ios_log_processes.get(did)
        thread = active_ios_log_readers.get(did)
        started_at = active_ios_log_started_at.get(did, now)
        last_seen = active_ios_log_last_seen.get(did, started_at)
        is_dead = (thread and not thread.is_alive()) or (proc and proc.poll() is not None)
        if is_dead:
            readers_to_stop.append((did, 'dead'))
        elif (proc is not None and did in observed_ids
              and now - started_at >= IOS_LOG_STALL_RESTART_SECONDS
              and now - last_seen >= IOS_LOG_STALL_RESTART_SECONDS):
            readers_to_stop.append((did, f'stalled: no logs for {int(now - last_seen)}s'))
    return readers_to_stop, stale_reader_ids


def _stabilize_ios_device_ids(observed_ids, known_ids):
    """Debounce transient empty iOS discovery results.

    ``tidevice list`` is a live usbmux query and can return no rows while its
    service is reconnecting. The device manager calls this under ``lock``;
    keeping the state here lets a single failed poll leave the active reader
    alone while a genuinely disconnected device is still removed after the
    grace window.
    """
    observed = {str(device_id).strip() for device_id in (observed_ids or set()) if str(device_id).strip()}
    known = {str(device_id).strip() for device_id in (known_ids or set()) if str(device_id).strip()}
    stable = set(observed)

    for device_id in observed:
        ios_device_missing_polls.pop(device_id, None)

    for device_id in known - observed:
        misses = ios_device_missing_polls.get(device_id, 0) + 1
        # Keep the miss count after the grace window so the device manager can
        # stop a stale syslog reader instead of re-adding it forever below.
        ios_device_missing_polls[device_id] = misses
        if misses <= IOS_DEVICE_MISSING_GRACE_POLLS:
            stable.add(device_id)

    return stable

def _ios_device_status_message():
    if _resolve_ios_tool("idevice_id") or _resolve_ios_tool("tidevice") or _resolve_ios_tool("system_profiler"):
        return "Waiting... (iOS: connect a trusted USB device)"
    return "Waiting... (iOS: install libimobiledevice/idevice_id or connect a trusted USB device)"

def _resolve_ios_syslog_command(device_id):
    # Prefer tidevice for iOS syslog; discovery has already restricted the
    # device ID to a physical USB transport.
    tidevice = _resolve_ios_tool("tidevice")
    if tidevice:
        return [tidevice, "-u", device_id, "syslog"]
    idevicesyslog = _resolve_ios_tool("idevicesyslog")
    if idevicesyslog:
        return [idevicesyslog, "-u", device_id]
    return None

def _normalize_ios_log_line(raw_line, device_id):
    raw = (raw_line or "").rstrip("\n")
    message = raw.strip()
    tag = ""

    tag_match = re.search(r'\b([A-Za-z0-9_.\-/\[\],]+)\s*[:：]\s*(.*)$', message)
    if tag_match:
        tag = tag_match.group(1).strip()
        message = tag_match.group(2).strip() or message

    return {
        "device_id": device_id,
        "device_name": get_ios_device_name(device_id),
        "tag": tag,
        "message": message,
        "raw_log": raw,
        "platform": "ios",
    }


IOS_LOG_TIMESTAMP_PATTERN = (
    r'(?:'
    r'[A-Z][a-z]{2}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?'
    r'|\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?'
    r'|\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?'
    r')'
)
IOS_LOG_RECORD_START_PATTERN = re.compile(
    rf'^{IOS_LOG_TIMESTAMP_PATTERN}\s+',
)
IOS_LOG_TIMESTAMP_CAPTURE_PATTERN = re.compile(
    rf'^(?P<timestamp>{IOS_LOG_TIMESTAMP_PATTERN})\s+',
)

def _normalize_sdk_search_text(text):
    if text is None:
        return ""
    normalized = str(text).lower()
    normalized = normalized.replace("–", "-").replace("—", "-").replace("−", "-")
    normalized = normalized.replace('"', '').replace("'", '')
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()

def _normalize_sdk_search_compact(text):
    return re.sub(r'[^a-z0-9]+', '', _normalize_sdk_search_text(text))


def _extract_first_sdk_version(line, patterns):
    for pattern in patterns:
        match = re.search(pattern, line, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def _normalize_sdk_network_name(name):
    if not name:
        return ""
    normalized = str(name).lower()
    normalized = normalized.replace("&", "and")
    normalized = re.sub(r'[^a-z0-9]+', '', normalized)
    return normalized


SDK_NETWORK_ALIASES = {
    "bigo": "bigoads",
    "fyber": "digitalturbinefyber",
    "digitalturbine": "digitalturbinefyber",
    "liftoff": "liftoffmonetizationvungle",
    "liftoffmonetization": "liftoffmonetizationvungle",
    "liftoffmonetize": "liftoffmonetizationvungle",
    "vungle": "liftoffmonetizationvungle",
    "google": "googleadmobandadmanager",
    "admob": "googleadmobandadmanager",
    "googleadmanager": "googleadmobandadmanager",
    "line": "lineads",
    "lineadsfivead": "lineads",
    "mytarget": "mytargetvkads",
    "vkads": "mytargetvkads",
    "pubmatic": "pubmaticopenwrap",
    "openwrap": "pubmaticopenwrap",
    "verve": "vervepubnative",
    "pubnative": "vervepubnative",
    "unity": "unityads",
    "meta": "metaaudiencenetwork",
    "facebook": "facebooksdk",
}

SDK_BASE_NETWORK_ALIASES = {
    "levelplay": "ironsource",
    "levelplayironsource": "ironsource",
}


def _sdk_base_match_network(name):
    """Match standard SDK labels while keeping a trailing ``- MAX`` distinct."""
    base_name = re.sub(r'^\s*max\s*/\s*', '', str(name or ''), flags=re.IGNORECASE).strip()
    normalized = _normalize_sdk_network_name(base_name)
    return SDK_BASE_NETWORK_ALIASES.get(
        normalized,
        SDK_NETWORK_ALIASES.get(normalized, normalized),
    )

CLOUDX_NETWORK_ALIASES = {
    "digitalturbine": "digitalturbinefyber",
    "fyber": "digitalturbinefyber",
    "liftoff": "liftoffmonetizationvungle",
    "liftoffmonetization": "liftoffmonetizationvungle",
    "liftoffmonetize": "liftoffmonetizationvungle",
    "vungle": "liftoffmonetizationvungle",
    "meta": "metaaudiencenetwork",
    "facebook": "metaaudiencenetwork",
    "unity": "unityads",
    "verve": "vervepubnative",
    "pubnative": "vervepubnative",
}


def _is_cloudx_sdk_network(name):
    return bool(re.search(r'\s*-\s*cloudx\s*$', str(name or ''), re.IGNORECASE))


def _cloudx_sdk_match_network(name):
    base_name = re.sub(r'\s*-\s*cloudx\s*$', '', str(name or ''), flags=re.IGNORECASE).strip()
    normalized = _normalize_sdk_network_name(base_name)
    return CLOUDX_NETWORK_ALIASES.get(normalized, normalized)


MAX_SDK_NETWORK_ALIASES = {
    "maxapplovin": "maxapplovin",
    "applovin": "maxapplovin",
    "applovinnetwork": "maxapplovin",
    "levelplay": "levelplayironsource",
    "levelplaynetwork": "levelplayironsource",
    "ironsource": "levelplayironsource",
    "ironsourcenetwork": "levelplayironsource",
    "bigo": "bigoads",
    "bigonetwork": "bigoads",
    "bigoadsnetwork": "bigoads",
    "digitalturbine": "digitalturbinefyber",
    "digitalturbinenetwork": "digitalturbinefyber",
    "dtexchange": "digitalturbinefyber",
    "dtexchangenetwork": "digitalturbinefyber",
    "fyber": "digitalturbinefyber",
    "fybernetwork": "digitalturbinefyber",
    "google": "googleadmobandadmanager",
    "googlenetwork": "googleadmobandadmanager",
    "admob": "googleadmobandadmanager",
    "admobnetwork": "googleadmobandadmanager",
    "googleadmob": "googleadmobandadmanager",
    "googleadmobnetwork": "googleadmobandadmanager",
    "googleadmanager": "googleadmobandadmanager",
    "googleadmanagernetwork": "googleadmobandadmanager",
    "facebook": "metaaudiencenetwork",
    "facebooknetwork": "metaaudiencenetwork",
    "meta": "metaaudiencenetwork",
    "metanetwork": "metaaudiencenetwork",
    "line": "lineads",
    "linenetwork": "lineads",
    "lineadsnetwork": "lineads",
    "lineadsfivead": "lineads",
    "liftoff": "liftoffmonetizationvungle",
    "liftoffnetwork": "liftoffmonetizationvungle",
    "liftoffmonetization": "liftoffmonetizationvungle",
    "liftoffmonetizationnetwork": "liftoffmonetizationvungle",
    "vungle": "liftoffmonetizationvungle",
    "vunglenetwork": "liftoffmonetizationvungle",
    "pangle": "pangletiktok",
    "panglenetwork": "pangletiktok",
    "tiktok": "pangletiktok",
    "tiktoknetwork": "pangletiktok",
    "bytedance": "pangletiktok",
    "bytedancenetwork": "pangletiktok",
    "pubmatic": "pubmaticopenwrap",
    "pubmaticnetwork": "pubmaticopenwrap",
    "openwrap": "pubmaticopenwrap",
    "openwrapnetwork": "pubmaticopenwrap",
    "unity": "unityads",
    "unityads": "unityads",
    "unityadsnetwork": "unityads",
    "verve": "vervepubnative",
    "vervenetwork": "vervepubnative",
    "pubnative": "vervepubnative",
    "pubnativenetwork": "vervepubnative",
    "mintegralnetwork": "mintegral",
    "inmobinetwork": "inmobi",
    "chartboostnetwork": "chartboost",
    "mobilefusenetwork": "mobilefuse",
    "moloconetwork": "moloco",
    "ogurynetwork": "ogury",
    "ogurypresage": "ogury",
    "ogurypresagenetwork": "ogury",
    "yandexnetwork": "yandex",
    "ysonetwork": "yso",
    "ascendxnetwork": "ascendx",
    "yeahmobi": "yeahmobimaticoo",
    "yeahmobinetwork": "yeahmobimaticoo",
    "maticoo": "yeahmobimaticoo",
    "maticoonetwork": "yeahmobimaticoo",
    "taurusxnetwork": "taurusx",
    "prado": "prado",
    "pradonetwork": "prado",
}


def _is_max_sdk_network(name):
    return bool(re.search(r'\s*-\s*max\s*$', str(name or ''), re.IGNORECASE))


def _max_sdk_base_name(name):
    return re.sub(r'\s*-\s*max\s*$', '', str(name or ''), flags=re.IGNORECASE).strip()


def _max_sdk_match_network(name):
    normalized = _normalize_sdk_network_name(_max_sdk_base_name(name))
    if normalized in MAX_SDK_NETWORK_ALIASES:
        return MAX_SDK_NETWORK_ALIASES[normalized]
    if normalized.endswith("network") and normalized[:-7] in MAX_SDK_NETWORK_ALIASES:
        return MAX_SDK_NETWORK_ALIASES[normalized[:-7]]
    return normalized


def _max_sdk_network_variants(name):
    normalized = _normalize_sdk_network_name(_max_sdk_base_name(name))
    if not normalized:
        return set()
    variants = {normalized}
    for variant in tuple(variants):
        for suffix in ("nativenetwork", "nativebidding", "bidding", "network"):
            if variant.endswith(suffix) and len(variant) > len(suffix):
                variants.add(variant[:-len(suffix)])
        if variant.endswith("presage") and len(variant) > len("presage"):
            variants.add(variant[:-len("presage")])
    for variant in tuple(variants):
        canonical = MAX_SDK_NETWORK_ALIASES.get(variant)
        if canonical:
            variants.add(canonical)
    return {variant for variant in variants if variant}


def _match_cloudx_sdk_expected_key(actual_name):
    actual_normalized = _normalize_sdk_network_name(actual_name)
    if not actual_normalized:
        return ''
    actual_match_network = CLOUDX_NETWORK_ALIASES.get(actual_normalized, actual_normalized)
    matches = [
        key for key, expected in sdk_check_expected_map.items()
        if expected.get('source') == 'cloudx'
        and expected.get('match_network') == actual_match_network
    ]
    return matches[0] if len(matches) == 1 else ''


def _match_max_sdk_expected_key(actual_name):
    actual_variants = _max_sdk_network_variants(actual_name)
    if not actual_variants:
        return ""

    expected_items = [
        (key, expected)
        for key, expected in sdk_check_expected_map.items()
        if expected.get("source") == "max" or _is_max_sdk_network(expected.get("display_name"))
    ]
    exact_matches = []
    for key, expected in expected_items:
        expected_variants = _max_sdk_network_variants(expected.get("display_name") or key)
        if actual_variants.intersection(expected_variants):
            exact_matches.append(key)
    if len(exact_matches) == 1:
        return exact_matches[0]

    fuzzy_matches = []
    for key, expected in expected_items:
        expected_variants = _max_sdk_network_variants(expected.get("display_name") or key)
        if any(
            len(actual_variant) >= 4
            and len(expected_variant) >= 4
            and (actual_variant in expected_variant or expected_variant in actual_variant)
            for actual_variant in actual_variants
            for expected_variant in expected_variants
        ):
            fuzzy_matches.append(key)
    return fuzzy_matches[0] if len(fuzzy_matches) == 1 else ""


def _match_sdk_expected_key(actual_name):
    actual_norm = _normalize_sdk_network_name(actual_name)
    if not actual_norm:
        return ""

    if _is_max_sdk_network(actual_name):
        return _match_max_sdk_expected_key(actual_name)

    is_cloudx = _is_cloudx_sdk_network(actual_name)
    expected_source = "cloudx" if is_cloudx else ""
    expected_items = [
        (key, expected)
        for key, expected in sdk_check_expected_map.items()
        if (expected.get("source") or "") == expected_source
    ]
    if not is_cloudx:
        expected_items = [
            (key, expected)
            for key, expected in sdk_check_expected_map.items()
            if (expected.get("source") or "") != "cloudx"
        ]

    if actual_norm in {key for key, _ in expected_items}:
        return actual_norm

    actual_match = (
        _cloudx_sdk_match_network(actual_name)
        if is_cloudx
        else _sdk_base_match_network(actual_name)
    )
    canonical_matches = []
    for key, expected in expected_items:
        expected_name = expected.get("display_name") or key
        expected_match = (
            _cloudx_sdk_match_network(expected_name)
            if is_cloudx
            else _sdk_base_match_network(expected_name)
        )
        if expected_match == actual_match:
            canonical_matches.append(key)
    if len(canonical_matches) == 1:
        return canonical_matches[0]

    matches = []
    for key, _ in expected_items:
        if actual_norm in key or key in actual_norm:
            matches.append(key)
    if len(matches) == 1:
        return matches[0]
    return ""


def _extract_sdk_comparable_version(value, expected_value=""):
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    match = re.search(r'(\d+(?:\.\d+)+)', text)
    if not match:
        return text
    version = match.group(1)
    expected_text = str(expected_value or "").strip()
    expected_match = re.search(r'(\d+(?:\.\d+)+)', expected_text)
    if expected_match:
        expected_parts = expected_match.group(1).split('.')
        version_parts = version.split('.')
        if len(version_parts) >= len(expected_parts):
            return '.'.join(version_parts[:len(expected_parts)])
    return version


def _sdk_result_status(actual_value, expected_value):
    actual_text = str(actual_value or "").strip()
    expected_text = str(expected_value or "").strip()
    if not expected_text:
        return "FOUND" if actual_text and actual_text.upper() not in {"NOT FOUND", "MISSING"} else "NOT_FOUND"
    if not actual_text or actual_text.upper() in {"NOT FOUND", "MISSING"}:
        if expected_text.casefold() == "removed":
            return "PASSED"
        return "FAILED"
    actual_compare = _extract_sdk_comparable_version(actual_text)
    expected_compare = _extract_sdk_comparable_version(expected_text)
    if expected_compare:
        actual_parts = actual_compare.split(".")
        expected_parts = expected_compare.split(".")
        if actual_parts and expected_parts:
            compare_len = min(len(actual_parts), len(expected_parts))
            if compare_len >= 3 and actual_parts[:compare_len] == expected_parts[:compare_len]:
                return "PASSED"
        return "PASSED" if actual_compare == expected_compare else "FAILED"
    return "FOUND"

def _is_sdk_input_header_line(line):
    cols = [
        _normalize_sdk_network_name(col.strip().strip('"').strip("'"))
        for col in re.split(r'\t+|\s{2,}', str(line or "").strip())
        if col.strip()
    ]
    return bool(cols) and cols[0] in {"adsnetwork", "network", "name"} and any(col in {"adapter", "sdk"} for col in cols[1:])

def _parse_sdk_expected_line(line):
    raw = (line or "").strip()
    if not raw:
        return None
    if _is_sdk_input_header_line(raw):
        return None

    if "\t" in raw:
        cols = [col.strip() for col in raw.split("\t")]
        if not cols or not cols[0]:
            return None
        source = "cloudx" if _is_cloudx_sdk_network(cols[0]) else ("max" if _is_max_sdk_network(cols[0]) else "")
        return {
            "network": cols[0],
            "adapter": cols[1] if len(cols) > 1 else "",
            "sdk": cols[2] if len(cols) > 2 else "",
            "log_search": cols[3] if len(cols) > 3 else "",
            "source": source,
            "match_network": (
                _cloudx_sdk_match_network(cols[0])
                if source == "cloudx"
                else (_max_sdk_match_network(cols[0]) if source == "max" else "")
            ),
        }

    version_matches = list(re.finditer(r'\b\d+(?:\.\d+)+\b', raw))
    if not version_matches:
        if _is_cloudx_sdk_network(raw) or _is_max_sdk_network(raw):
            source = "cloudx" if _is_cloudx_sdk_network(raw) else "max"
            return {
                "network": raw,
                "adapter": "",
                "sdk": "",
                "log_search": "",
                "source": source,
                "match_network": (
                    _cloudx_sdk_match_network(raw)
                    if source == "cloudx"
                    else _max_sdk_match_network(raw)
                ),
            }
        return None

    network = raw[:version_matches[0].start()].strip(" -–—\t")
    if not network:
        return None

    versions = [m.group(0) for m in version_matches]
    adapter = ""
    sdk = ""
    if len(versions) >= 2:
        adapter, sdk = versions[0], versions[1]
    else:
        expected_key = _normalize_sdk_network_name(network)
        if expected_key in {"appmetrica", "firebasecrashlytics"}:
            sdk = versions[0]
        else:
            adapter = versions[0]

    return {
        "network": network,
        "adapter": adapter,
        "sdk": sdk,
        "log_search": "",
        "source": (
            "cloudx"
            if _is_cloudx_sdk_network(network)
            else ("max" if _is_max_sdk_network(network) else "")
        ),
        "match_network": (
            _cloudx_sdk_match_network(network)
            if _is_cloudx_sdk_network(network)
            else (_max_sdk_match_network(network) if _is_max_sdk_network(network) else "")
        ),
    }

def _extract_sdk_search_pattern(line):
    match = re.search(r'"search_pattern"\s*:\s*(.+)$', line or "", re.IGNORECASE)
    if not match:
        return ""
    pattern = match.group(1).strip().rstrip(",")
    if (pattern.startswith('"') and pattern.endswith('"')) or (pattern.startswith("'") and pattern.endswith("'")):
        pattern = pattern[1:-1]
    else:
        pattern = pattern.strip('"').strip("'")
    return pattern.strip()

def _parse_sdk_search_pattern_line(line, current_label=""):
    pattern = _extract_sdk_search_pattern(line)
    if not pattern:
        return None

    prefix = line[:re.search(r'"search_pattern"\s*:', line, re.IGNORECASE).start()].rstrip(", ").strip().strip('"').strip("'")
    version_match = re.search(r'\b\d+(?:\.\d+)+\b', prefix)
    if version_match:
        name_part = prefix[:version_match.start()].strip(" -–—\t")
        version = version_match.group(0)
    else:
        name_part = prefix or current_label or pattern
        version_match = re.search(r'\b\d+(?:\.\d+)+\b', pattern)
        version = version_match.group(0) if version_match else ""

    network = current_label or name_part or pattern
    name_norm = _normalize_sdk_network_name(name_part)
    network_norm = _normalize_sdk_network_name(network)

    field = "adapter"
    if name_norm in {"sdk", "sdkversion"} or "sdkversion" in name_norm:
        field = "sdk"
    elif network_norm.startswith("firebase") or network_norm in {"appmetrica"}:
        field = "sdk"
    elif network_norm == "adquality":
        field = "both"
    elif name_norm in {"unity", "plugin", "pluginversion"}:
        field = "adapter"

    return {
        "network": network,
        "display_name": network,
        "version_label": name_part or network,
        "version": version,
        "field": field,
        "search_pattern": pattern,
        "search_pattern_normalized": _normalize_sdk_search_text(pattern),
        "search_pattern_compact": _normalize_sdk_search_compact(pattern),
        "source": "cloudx" if _is_cloudx_sdk_network(network) else "",
        "match_network": _cloudx_sdk_match_network(network) if _is_cloudx_sdk_network(network) else "",
    }

def _register_sdk_expected(network, adapter="", sdk="", log_search="", source="", match_network=""):
    expected_key = _normalize_sdk_network_name(network)
    if not expected_key:
        return ""
    existing = sdk_check_expected_map.setdefault(expected_key, {
        "display_name": network,
        "adapter": "",
        "sdk": "",
        "log_search": "",
        "source": "",
        "match_network": "",
    })
    existing["display_name"] = existing.get("display_name") or network
    if adapter:
        existing["adapter"] = adapter
    if sdk:
        existing["sdk"] = sdk
    if log_search:
        existing["log_search"] = log_search
    if source:
        existing["source"] = source
    if match_network:
        existing["match_network"] = match_network
    if expected_key not in sdk_check_expected_order:
        sdk_check_expected_order.append(expected_key)
    return expected_key


def _sdk_state_network_keys(device_id):
    keys = []
    seen = set()
    device_state = sdk_check_runtime_state.get(device_id, {})
    for expected_key in sdk_check_expected_order:
        if expected_key not in seen:
            keys.append(expected_key)
            seen.add(expected_key)
    for network_key in device_state.keys():
        if network_key not in seen:
            keys.append(network_key)
            seen.add(network_key)
    return keys


def _ensure_sdk_expected_blocks_for_device(device_id):
    device_state = sdk_check_runtime_state.setdefault(device_id, {})
    for expected_key in sdk_check_expected_order:
        expected = sdk_check_expected_map.get(expected_key, {})
        if expected_key not in device_state:
            device_state[expected_key] = {
                "display_name": expected.get("display_name", expected_key),
                "sdk_version": "",
                "adapter_version": "",
                "native_version": "",
                "observed_network_name": "",
                "adapter_missing": False,
                "verification": "",
                "expected_key": expected_key,
                "updated_at": time.time(),
            }


def _extract_json_field_version(line, field_name):
    pattern = re.compile(rf'"{re.escape(field_name)}"\s*:\s*"([^"]+)"', re.IGNORECASE)
    match = pattern.search(line)
    return match.group(1).strip() if match else ""


def _process_sdk_external_line(line, device_id):
    normalized_line = _normalize_sdk_search_text(line)
    compact_line = _normalize_sdk_search_compact(line)
    platform = active_platform
    changed = False

    def update_block(network_name, sdk_version=None, adapter_version=None, adapter_missing=False):
        nonlocal changed
        block = _sdk_check_block_for_device(device_id, network_name)
        if sdk_version is not None and block.get("sdk_version") != sdk_version:
            block["sdk_version"] = sdk_version
            changed = True
        if adapter_version is not None and (block.get("adapter_version") != adapter_version or block.get("adapter_missing")):
            block["adapter_version"] = adapter_version
            block["adapter_missing"] = False
            changed = True
        if adapter_missing and (not block.get("adapter_missing") or block.get("adapter_version")):
            block["adapter_missing"] = True
            block["adapter_version"] = ""
            changed = True

    def update_single_expected_version(network_name, version):
        """Put a one-version log into the version column used by its preset row."""
        expected_key = _match_sdk_expected_key(network_name) or _normalize_sdk_network_name(network_name)
        expected = sdk_check_expected_map.get(expected_key, {})
        if expected.get("sdk") and not expected.get("adapter"):
            update_block(network_name, sdk_version=version)
        else:
            update_block(network_name, adapter_version=version)

    if platform == "ios":
        for item in sdk_check_search_list:
            pattern_norm = item.get("search_pattern_normalized", "")
            pattern_compact = item.get("search_pattern_compact", "")
            if pattern_norm and (pattern_norm in normalized_line or (pattern_compact and pattern_compact in compact_line)):
                version = item.get("version", "")
                field = item.get("field", "adapter")
                network_name = item.get("display_name") or item.get("network") or item.get("version_label") or item.get("search_pattern")
                expected_key = _match_sdk_expected_key(network_name) or _normalize_sdk_network_name(network_name)
                if expected_key == "firebasecrashlytics":
                    continue
                if field == "sdk":
                    update_block(network_name, sdk_version=version)
                elif field == "both":
                    update_block(network_name, sdk_version=version, adapter_version=version)
                else:
                    update_block(network_name, adapter_version=version)

    if "advertysdk" in normalized_line:
        match = re.search(r'AdvertySDK\s+([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("Adverty", adapter_version=match.group(1))
    if "gadsmeservice->initialize" in normalized_line:
        version = _extract_json_field_version(line, "Version")
        if version:
            update_block("Gadsme", adapter_version=version)
    if "gadsme sdk version" in normalized_line:
        match = re.search(r'Gadsme\s+sdk\s+version\s*:\s*([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("Gadsme", adapter_version=match.group(1))
    if "audiomob" in normalized_line:
        match = AUDIOMOB_OPEN_MEASUREMENT_SDK_PATTERN.search(line)
        if not match:
            match = re.search(r'Audiomob\s+v?([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            version = match.groupdict().get("version") or match.group(1)
            update_single_expected_version("AudioMob", version)
    if "adquality" in normalized_line:
        match = re.search(r'AdQuality\s+([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            version = match.group(1)
            update_block("AdQuality", adapter_version=version, sdk_version=version)
    if "facebook unity sdk" in normalized_line:
        match = re.search(r'Facebook\s+Unity\s+SDK\s+v?([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("Facebook SDK", adapter_version=match.group(1))
    if "fbandroidsdk/" in normalized_line:
        match = re.search(r'FBAndroidSDK/([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("Facebook SDK", sdk_version=match.group(1))
    if platform == "ios" and "fbiossdk/" in normalized_line:
        match = re.search(r'FBiOSSDK/([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("Facebook SDK", sdk_version=match.group(1))
    if "appmetrica" in normalized_line and "release type" in normalized_line and "version" in normalized_line:
        match = re.search(r'AppMetrica.*?Version\s+([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("AppMetrica", sdk_version=match.group(1))
    if "appsflyer" in normalized_line or "android_unity" in normalized_line:
        adapter_match = re.search(r'"pluginVersion"\s*:\s*"([0-9]+(?:\.[0-9]+)+)"', line, re.IGNORECASE)
        if not adapter_match and "android_unity" in normalized_line:
            adapter_match = re.search(r'"platform"\s*:\s*"android_unity"[^{}]*?"version"\s*:\s*"([0-9]+(?:\.[0-9]+)+)"', line, re.IGNORECASE)
        if adapter_match:
            update_block("Appsflyer", adapter_version=adapter_match.group(1))
        appsflyer_sdk = ""
        sdk_json_match = re.search(r'"AppsFlyer\.getSdkVersion\(\)"\s*:\s*"version\s*:\s*([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if sdk_json_match:
            appsflyer_sdk = sdk_json_match.group(1)
        if not appsflyer_sdk:
            sdk_field = _extract_json_field_version(line, "AppsFlyer.getSdkVersion()")
            if sdk_field:
                sdk_match = re.search(r'version\s*:\s*([0-9]+(?:\.[0-9]+)+)', sdk_field, re.IGNORECASE)
                if sdk_match:
                    appsflyer_sdk = sdk_match.group(1)
        if not appsflyer_sdk and "appsflyer:" in normalized_line:
            match = re.search(r'AppsFlyer:\s*\(v?([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
            if match:
                appsflyer_sdk = match.group(1)
        if appsflyer_sdk:
            update_block("Appsflyer", sdk_version=appsflyer_sdk)
    if "clientsdk:" in normalized_line and "unity" in normalized_line:
        if platform == "ios":
            match = re.search(r'unity\s*([0-9]+(?:\.[0-9]+)+)\s*@\s*ios\s*([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        else:
            match = re.search(r'unity\s*([0-9]+(?:\.[0-9]+)+)\s*@\s*(?:android|adroid)\s*([0-9]+(?:\.[0-9]+)+)', line, re.IGNORECASE)
        if match:
            update_block("Adjust", adapter_version=match.group(1), sdk_version=match.group(2))
    if platform == "ios":
        firebase_analytics_version = _extract_first_sdk_version(line, [
            r'Analytics\s+v\.?([0-9]+(?:\.[0-9]+)+)\s+started',
            r'FirebaseAnalytics[^0-9]*([0-9]+(?:\.[0-9]+)+)',
        ])
        if firebase_analytics_version:
            update_block("Firebase Analytics", sdk_version=firebase_analytics_version)
        firebase_messaging_version = _extract_first_sdk_version(line, [
            r'([0-9]+(?:\.[0-9]+)+)\s*-\s*\[FirebaseMessaging\]',
            r'FirebaseMessaging[^0-9]*([0-9]+(?:\.[0-9]+)+)',
        ])
        if firebase_messaging_version:
            update_block("Firebase Cloud Messaging", sdk_version=firebase_messaging_version)
    firebase_crashlytics_version = ""
    if platform == "ios":
        firebase_crashlytics_version = _extract_first_sdk_version(line, [
            r'\[Firebase/Crashlytics\]\s*Version\s*([0-9]+(?:\.[0-9]+)+)',
        ])
        if firebase_crashlytics_version and "[firebase/crashlytics]" in normalized_line:
            update_block("Firebase Crashlytics", sdk_version=firebase_crashlytics_version)
    else:
        firebase_crashlytics_match = re.search(
            r'Initializing Firebase Crashlytics\s+([0-9]+(?:\.[0-9]+)+)',
            line,
            re.IGNORECASE,
        )
        if firebase_crashlytics_match and "initializing firebase crashlytics" in normalized_line:
            update_block("Firebase Crashlytics", sdk_version=firebase_crashlytics_match.group(1))
    if platform == "ios":
        firebase_performance_version = _extract_first_sdk_version(line, [
            r'([0-9]+(?:\.[0-9]+)+)\s*-\s*\[FirebasePerformance\]',
            r'FirebasePerformance[^0-9]*([0-9]+(?:\.[0-9]+)+)',
        ])
        if firebase_performance_version:
            update_block("Firebase Performance Monitoring", sdk_version=firebase_performance_version)

    return changed


def _update_sdk_max_block(device_id, expected_key, sdk_version=None, adapter_version=None, observed_network_name=""):
    expected = sdk_check_expected_map.get(expected_key, {})
    if not expected:
        return False
    display_name = expected.get("display_name") or expected_key
    block = _sdk_check_block_for_device(device_id, display_name)
    changed = False
    if sdk_version is not None and block.get("sdk_version") != sdk_version:
        block["sdk_version"] = sdk_version
        changed = True
    if adapter_version is not None and (
        block.get("adapter_version") != adapter_version or block.get("adapter_missing")
    ):
        block["adapter_version"] = adapter_version
        block["adapter_missing"] = False
        changed = True
    if observed_network_name and block.get("observed_network_name") != observed_network_name:
        block["observed_network_name"] = observed_network_name
        changed = True
    block["expected_key"] = expected_key
    block["display_name"] = display_name
    block["updated_at"] = time.time()
    return changed


def _process_sdk_max_line(line, device_id):
    """Parse MAX core version logs and MAX adapter health telemetry."""
    if active_platform not in {"android", "ios"}:
        return False
    raw_line = str(line or "")
    changed = False

    core_matches = (
        SDK_MAX_CORE_INITIALIZATION_PATTERN.search(raw_line),
        SDK_MAX_CORE_VERSION_PATTERN.search(raw_line),
    )
    for core_match in core_matches:
        if not core_match:
            continue
        expected_key = _match_max_sdk_expected_key("MAX / AppLovin - MAX")
        if expected_key:
            changed = _update_sdk_max_block(
                device_id,
                expected_key,
                sdk_version=core_match.group("version"),
            ) or changed

    if SDK_MAX_HEALTH_EVENTS_PATTERN.search(raw_line):
        adapter_match = SDK_MAX_ADAPTER_VERSION_FIELD_PATTERN.search(raw_line)
        network_match = SDK_MAX_NETWORK_NAME_FIELD_PATTERN.search(raw_line)
        if adapter_match and network_match:
            observed_network_name = network_match.group("network").strip().strip("'\"")
            expected_key = _match_max_sdk_expected_key(observed_network_name)
            if expected_key:
                changed = _update_sdk_max_block(
                    device_id,
                    expected_key,
                    adapter_version=adapter_match.group("version"),
                    observed_network_name=observed_network_name,
                ) or changed

    return changed


def _process_sdk_max_stream_line(line, device_id):
    """Parse MAX SDK telemetry without changing the shared iOS log reader."""
    raw_line = str(line or "")
    if active_platform != "ios":
        return _process_sdk_max_line(raw_line, device_id)

    pending = sdk_max_ios_pending_lines.get(device_id, "")
    if pending:
        if SDK_MAX_IOS_RECORD_START_PATTERN.match(raw_line):
            sdk_max_ios_pending_lines.pop(device_id, None)
            changed = _process_sdk_max_line(pending, device_id)
            return _process_sdk_max_stream_line(raw_line, device_id) or changed

        combined = f"{pending}\n{raw_line}"
        if "}" in raw_line:
            sdk_max_ios_pending_lines.pop(device_id, None)
            return _process_sdk_max_line(combined, device_id)
        sdk_max_ios_pending_lines[device_id] = combined
        return False

    core_pending = sdk_max_ios_core_pending_lines.get(device_id, "")
    if core_pending:
        if SDK_MAX_IOS_RECORD_START_PATTERN.match(raw_line):
            sdk_max_ios_core_pending_lines.pop(device_id, None)
            changed = _process_sdk_max_line(core_pending, device_id)
            return _process_sdk_max_stream_line(raw_line, device_id) or changed

        combined = f"{core_pending}\n{raw_line}"
        if SDK_MAX_CORE_VERSION_PATTERN.search(combined):
            sdk_max_ios_core_pending_lines.pop(device_id, None)
            return _process_sdk_max_line(combined, device_id)
        sdk_max_ios_core_pending_lines[device_id] = combined
        return False

    if (
        SDK_MAX_HEALTH_EVENTS_PATTERN.search(raw_line)
        and "{" in raw_line
        and "}" not in raw_line
    ):
        sdk_max_ios_pending_lines[device_id] = raw_line
        return False

    if SDK_MAX_IOS_CORE_MARKER_PATTERN.search(raw_line):
        if SDK_MAX_CORE_VERSION_PATTERN.search(raw_line):
            return _process_sdk_max_line(raw_line, device_id)
        sdk_max_ios_core_pending_lines[device_id] = raw_line
        return False

    return _process_sdk_max_line(raw_line, device_id)


# --- HELPER FUNCTIONS FOR FORMATTING ---
def format_json_html(data):
    """Format JSON object to HTML string with indentation and colors"""
    try:
        if isinstance(data, str):
            # Try to parse string as JSON if possible
            if data.strip().startswith('{') or data.strip().startswith('['):
                try:
                    data = json.loads(data)
                except: pass
        
        # If it's a dict or list, dump it pretty
        if isinstance(data, (dict, list)):
            # ensure_ascii=False để hiển thị tiếng Việt đúng
            json_str = json.dumps(data, indent=2, ensure_ascii=False)
            return f'<pre class="text-xs bg-gray-50 p-2 rounded border border-gray-200 overflow-x-auto font-mono text-gray-700">{html.escape(json_str)}</pre>'
        return str(data)
    except:
        return str(data)


def format_adjust_config_html(data):
    try:
        if isinstance(data, str):
            data = json.loads(data)
        if not isinstance(data, dict):
            return format_json_html(data)

        special_keys = {
            "IsSendingInBackgroundEnabled",
            "IsCostDataInAttributionEnabled",
            "IsAdServicesEnabled",
            "IsIdfaReadingEnabled",
        }

        def render_value(value):
            if isinstance(value, dict):
                return render_dict(value, 0)
            if isinstance(value, list):
                return html.escape(json.dumps(value, ensure_ascii=False))
            if isinstance(value, str):
                return f'"{html.escape(value)}"'
            return html.escape(json.dumps(value, ensure_ascii=False))

        def render_dict(obj, indent_level):
            indent = "&nbsp;" * (indent_level * 4)
            inner_indent = "&nbsp;" * ((indent_level + 1) * 4)
            lines = [f"{indent}{{"]
            entries = list(obj.items())
            for idx, (key, value) in enumerate(entries):
                comma = "," if idx < len(entries) - 1 else ""
                line_class = ""
                if key in special_keys:
                    line_class = "text-green-600 font-semibold" if value is True else "text-red-600 font-semibold"

                if isinstance(value, dict):
                    value_html = render_dict(value, indent_level + 1)
                    lines.append(
                        f'{inner_indent}<span class="{line_class}">"{html.escape(str(key))}": {value_html}{comma}</span>'
                    )
                else:
                    lines.append(
                        f'{inner_indent}<span class="{line_class}">"{html.escape(str(key))}": {render_value(value)}{comma}</span>'
                    )
            lines.append(f"{indent}}}")
            return "<br>".join(lines)

        rendered = render_dict(data, 0)
        return f'<pre class="text-xs bg-gray-50 p-2 rounded border border-gray-200 overflow-x-auto font-mono text-gray-700">{rendered}</pre>'
    except Exception:
        return format_json_html(data)


def format_param_issue_html(title, items, color_class, chunk_size=4):
    if not items:
        return ""
    ordered = [html.escape(str(x)) for x in sorted(items)]
    lines = "<br>".join(f"&nbsp;&nbsp;{item}" for item in ordered)
    return (
        f'<div class="{color_class} text-xs mb-2 break-words leading-5">'
        f'<div>{html.escape(title)}:</div>'
        f'<div class="font-normal">{lines}</div>'
        f'</div>'
    )

def extract_json_object_from_text(text):
    """Extract first JSON object substring from text by brace matching."""
    try:
        start_idx = text.find('{')
        if start_idx == -1:
            return None
        open_braces = 0
        for i in range(start_idx, len(text)):
            if text[i] == '{':
                open_braces += 1
            elif text[i] == '}':
                open_braces -= 1
                if open_braces == 0:
                    return text[start_idx:i+1]
        return None
    except:
        return None


def _parse_price_rotation_log(log_entry, device_id):
    """Parse exact RewardedBidding and InterstitialBidding markers."""
    raw_log = str(log_entry or "").strip()
    if not raw_log:
        return None

    marker_candidates = []
    for bidding_name, marker_prefix in PRICE_ROTATION_PREFIXES:
        marker_start = raw_log.find(marker_prefix)
        if marker_start == -1:
            continue
        type_start = marker_start + len(marker_prefix)
        if type_start < len(raw_log) and raw_log[type_start] in ",]":
            marker_candidates.append((marker_start, bidding_name, marker_prefix))

    if not marker_candidates:
        return None

    marker_start, bidding_name, marker_prefix = min(marker_candidates, key=lambda item: item[0])
    type_start = marker_start + len(marker_prefix)
    marker_end = raw_log.find("]", type_start)
    if marker_end == -1:
        return None

    rotation_type = raw_log[type_start:marker_end].strip().lstrip(",").strip()
    if not rotation_type:
        service_text = raw_log[marker_end + 1:].lstrip()
        cap_service = f"{bidding_name}CapService"
        rotation_type = f"{bidding_name}Cap" if service_text.startswith(cap_service) else f"{bidding_name}Bidding"
    payload_text = raw_log[marker_end + 1:].strip()
    json_text = extract_json_object_from_text(payload_text)
    parsed_json = None
    details = payload_text or "No JSON payload"

    if json_text:
        try:
            parsed_json = json.loads(json_text)
            details = json.dumps(parsed_json, ensure_ascii=False, indent=2)
        except (TypeError, ValueError):
            details = json_text

    return {
        "device_id": device_id,
        "device_name": get_device_name(device_id),
        "bidding_name": bidding_name,
        "type": rotation_type,
        "details": details,
        "raw_log": raw_log,
        "json_data": json.dumps(parsed_json, ensure_ascii=False) if parsed_json is not None else "{}",
    }


def process_price_rotation_log(log_entry, device_id):
    if is_paused:
        return

    parsed = _parse_price_rotation_log(log_entry, device_id)
    if not parsed:
        return
    if _normalize_price_rotation_type(parsed.get("type")) in PRICE_ROTATION_HIDDEN_TYPES:
        return

    with lock:
        if not _accept_exact_record(
            "price_rotation",
            device_id,
            parsed.get("raw_log", ""),
            parsed.get("bidding_name", ""),
            parsed.get("type", ""),
        ):
            return
        _live_table_append("price_rotation", parsed)
        snapshot = True

def _decode_ios_levelplay_json_text(text):
    # iOS syslog sometimes writes \" as \134", which is not valid JSON.
    return (text or "").replace(r"\134", "\\")

def _parse_callback_json_payload(json_str):
    for candidate in (json_str, _decode_ios_levelplay_json_text(json_str)):
        try:
            data = json.loads(candidate)
            break
        except Exception:
            data = None
    if not isinstance(data, dict):
        return None

    for key in ("adInfo", "error", "impressionData", "LevelPlayImpressionData"):
        value = data.get(key)
        if isinstance(value, str):
            nested_text = _decode_ios_levelplay_json_text(value).strip()
            if nested_text.startswith("{") and nested_text.endswith("}"):
                try:
                    data[key] = json.loads(nested_text)
                except Exception:
                    pass
    return data

def _find_first_key_deep(data, keys):
    if isinstance(data, dict):
        for key, value in data.items():
            if key in keys:
                return value
            found = _find_first_key_deep(value, keys)
            if found is not None:
                return found
    elif isinstance(data, list):
        for value in data:
            found = _find_first_key_deep(value, keys)
            if found is not None:
                return found
    return None

def _normalize_ad_format_label(value, fallback=""):
    raw = str(value or fallback or "").strip().lower()
    if raw == "mrec":
        return "MREC"
    if raw == "banner":
        return "Banner"
    if raw in {"reward", "rewarded", "rewarded_video", "rewardedvideo"}:
        return "Reward"
    if raw in {"interstitial", "inter"}:
        return "Interstitial"
    return str(value or fallback or "").strip()

def _levelplay_impression_event_name(data, fallback_format=""):
    ad_format = _find_first_key_deep(data, {"AdFormat", "adFormat", "ad_format"})
    label = _normalize_ad_format_label(ad_format, fallback_format)
    return f"_OnLevelPlayImpressionDataReadyEvent - {label}" if label else "_OnLevelPlayImpressionDataReadyEvent"


def _normalize_levelplay_ad_unit_label(value):
    raw = str(value or "").strip()
    normalized = raw.lower().replace("_", "").replace("-", "")
    if "banner" in normalized:
        return "Banner"
    if "reward" in normalized:
        return "Rewarded"
    if "interstitial" in normalized or normalized == "inter":
        return "Interstitial"
    return raw


def _levelplay_impression_data_event_name(data):
    ad_unit_name = ""
    for key in ("mediationAdUnitName", "adUnitName", "adUnit", "adFormat"):
        value = _find_first_key_deep(data, {key})
        if str(value or "").strip():
            ad_unit_name = _normalize_levelplay_ad_unit_label(value)
            break
    return f"LevelPlayImpressionData - {ad_unit_name}" if ad_unit_name else "LevelPlayImpressionData"


def _find_named_callback_event(log_entry, event_names):
    raw_log = str(log_entry or "")
    return next((event_name for event_name in event_names if event_name in raw_log), "")


def _named_callback_ad_unit_label(log_entry, event_name=""):
    """Resolve the Reward/Interstitial suffix used by AscendX and CloudX callbacks."""
    if event_name in {"_OnAdReceivedRewardEvent", "OnAdRewarded"}:
        return "Reward"

    raw_log = str(log_entry or "")
    json_str = extract_json_object_from_text(raw_log)
    if json_str:
        data = _parse_callback_json_payload(json_str)
        if data is not None:
            for key in ("adUnitName", "mediationAdUnitName", "adUnit", "adFormat", "ad_format", "format"):
                value = _find_first_key_deep(data, {key})
                label = _normalize_ad_format_label(value)
                if label in {"Reward", "Interstitial"}:
                    return label

    lowered = raw_log.lower()
    if re.search(r"\binterstitial\b", lowered):
        return "Interstitial"
    if re.search(r"\breward(?:ed)?\b|rewarded[_ -]?video", lowered):
        return "Reward"
    return ""


def _record_named_provider_callback(log_entry, device_id, event_name, callback_type):
    after_keyword = str(log_entry or "").split(event_name, 1)[1].strip()
    json_str = extract_json_object_from_text(after_keyword)
    details = "Callback fired"
    json_data_for_log = "{}"
    if json_str:
        data = _parse_callback_json_payload(json_str)
        if data is not None:
            details = format_json_html(data)
            json_data_for_log = json.dumps(data, ensure_ascii=False)
        else:
            details = (
                '<div class="text-xs font-mono break-all text-red-600">JSON Parse Error</div>'
                f'<div class="text-xs font-mono break-all">{html.escape(json_str)}</div>'
            )
    elif after_keyword:
        details = f'<div class="text-xs font-mono break-all">{html.escape(after_keyword.lstrip(":").strip())}</div>'

    ad_unit_label = _named_callback_ad_unit_label(log_entry, event_name)
    display_name = f"{event_name} - {ad_unit_label}" if ad_unit_label else event_name
    _append_callback_ad_row({
        "device_id": device_id,
        "device_name": get_device_name(device_id),
        "type": callback_type,
        "event_name": display_name,
        "details": details,
        "raw_log": str(log_entry or "").strip(),
        "json_data": json_data_for_log,
    })

def split_top_level_csv(text):
    parts = []
    current = []
    depth = 0
    in_single = False
    in_double = False
    escape = False
    for ch in text:
        current.append(ch)
        if escape:
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == "'" and not in_double:
            in_single = not in_single
            continue
        if ch == '"' and not in_single:
            in_double = not in_double
            continue
        if in_single or in_double:
            continue
        if ch in '{[(':
            depth += 1
        elif ch in '}])':
            depth = max(0, depth - 1)
        elif ch == ',' and depth == 0:
            current.pop()
            part = ''.join(current).strip()
            if part:
                parts.append(part)
            current = []
    tail = ''.join(current).strip()
    if tail:
        parts.append(tail)
    return parts


def parse_levelplay_impression_text(payload):
    payload = (payload or '').strip()
    marker = 'LevelPlayImpressionData'
    if payload.startswith(marker):
        payload = payload[len(marker):].strip()
    if not (payload.startswith('{') and payload.endswith('}')):
        return None
    inner = payload[1:-1].strip()
    result = {}
    for item in split_top_level_csv(inner):
        if '=' not in item:
            continue
        key, raw = item.split('=', 1)
        key = key.strip()
        raw = raw.strip()
        if raw.startswith("'") and raw.endswith("'"):
            value = raw[1:-1]
        elif raw in ('null', 'None'):
            value = None
        elif raw == '':
            value = ''
        else:
            try:
                if any(ch in raw for ch in '.eE'):
                    value = float(raw)
                else:
                    value = int(raw)
            except:
                value = raw
        result[key] = value
    return result or None


def parse_appmetrica_adrevenue_text(payload):
    payload = (payload or "").strip()
    marker = "AdRevenue{"
    if payload.startswith(marker) and payload.endswith("}"):
        payload = payload[len(marker):-1].strip()

    result = {}
    for item in split_top_level_csv(payload):
        if "=" not in item:
            continue
        key, raw = item.split("=", 1)
        key = key.strip()
        raw = raw.strip()

        if key == "payload" and raw.startswith("{") and raw.endswith("}"):
            try:
                result[key] = json.loads(raw)
            except:
                result[key] = raw
            continue

        if raw.startswith("'") and raw.endswith("'"):
            value = raw[1:-1]
        elif raw in ("null", "None", "<null>"):
            value = None
        elif raw == "":
            value = ""
        elif raw.lower() == "true":
            value = True
        elif raw.lower() == "false":
            value = False
        else:
            try:
                if any(ch in raw for ch in ".eE"):
                    value = float(raw)
                else:
                    value = int(raw)
            except:
                value = raw
        result[key] = value
    return result or None

def _read_profile_sheet(ws, allow_default_fill=True):
    event_map = {}
    current_event = None
    header_row = 1
    game_name = ""

    if str(ws.cell(1, 2).value or "").strip().lower() == "game":
        game_name = str(ws.cell(1, 3).value or "").strip()
        header_row = 2

    for r in range(header_row + 1, ws.max_row + 1):
        event_val = ws.cell(r, 2).value
        param_cell = ws.cell(r, 3)
        param_val = param_cell.value

        if event_val:
            current_event = str(event_val).strip()
            if current_event:
                event_map.setdefault(current_event, {"specific": [], "default": []})

        if current_event and param_val:
            param = str(param_val).strip()
            fill = param_cell.fill
            fg = fill.fgColor.rgb if fill and fill.fgColor else None
            is_default = allow_default_fill and (fg == DEFAULT_PARAM_FILL)
            bucket = "default" if is_default else "specific"
            event_map[current_event][bucket].append(param)

    seen = set()
    merged_default = []
    for data in event_map.values():
        for p in data["default"]:
            if p not in seen:
                seen.add(p)
                merged_default.append(p)
    return game_name, merged_default, {k: v["specific"] for k, v in event_map.items()}


def _normalize_default_ad_event_token(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().casefold())


def _normalize_default_ad_event_provider(value):
    aliases = {
        "ironsource": "ironsource",
        "iron": "ironsource",
        "max": "max",
        "applovin": "max",
        "applovinmax": "max",
        "cloudx": "cloudx",
        "ascendx": "ascendx",
    }
    return aliases.get(_normalize_default_ad_event_token(value), "")


def _normalize_default_ad_event_simple_category(value):
    aliases = {
        "audioads": "audioads",
        "audioad": "audioads",
        "inplayads": "inplayads",
        "inplayad": "inplayads",
        "inplay": "inplayads",
    }
    return aliases.get(_normalize_default_ad_event_token(value), "")


def _normalize_default_ad_event_simple_provider(value):
    aliases = {
        "odeeo": "odeeo",
        "audiomob": "audiomob",
        "adverty": "adverty",
        "gadsme": "gadsme",
    }
    return aliases.get(_normalize_default_ad_event_token(value), "")


def _normalize_default_ad_event_format(value):
    aliases = {
        "interstitial": "interstitial",
        "rewarded": "rewarded",
        "rewardedvideo": "rewarded",
        "banner": "banner",
        "mrec": "mrec",
        "mediumrectangle": "mrec",
    }
    return aliases.get(_normalize_default_ad_event_token(value), "")


def _normalize_default_ad_event_param_key(value):
    token = _normalize_default_ad_event_token(value)
    return {
        "adplatform": "ad_platform",
        "adformat": "ad_format",
    }.get(token, "")


def _normalize_default_ad_event_name(value):
    return str(value or "").strip().casefold()


def _is_default_ad_event_name(value):
    text = str(value or "").strip()
    if not text.casefold().startswith("ad_"):
        return False
    return _normalize_default_ad_event_param_key(text) == ""


def _new_default_ad_event_config():
    return {
        ad_format: {provider: [] for provider in DEFAULT_AD_EVENT_PROVIDERS}
        for ad_format in DEFAULT_AD_EVENT_FORMATS
    }


def _new_default_ad_event_simple_config():
    return {
        category: {provider: [] for provider in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS[category]}
        for category in DEFAULT_AD_EVENT_SIMPLE_CATEGORIES
    }


def _read_default_ad_event_sheet(ws, provider):
    """Read one provider sheet laid out as horizontal or vertical format blocks."""
    records = {}

    # The reference template has four merged headings on one row. The first
    # occurrence of each format is the heading; later occurrences are values
    # in the ad_format parameter rows and must not become new blocks.
    headings_by_format = {}
    for row in ws.iter_rows():
        for cell in row:
            ad_format = _normalize_default_ad_event_format(cell.value)
            if ad_format and ad_format not in headings_by_format:
                headings_by_format[ad_format] = (cell.row, cell.column)

    headings = sorted(
        (
            ad_format,
            position[0],
            position[1],
        )
        for ad_format, position in headings_by_format.items()
    )

    for ad_format, heading_row, heading_column in headings:
        # A following heading in the same column means vertically stacked
        # blocks. A following heading in the same row means side-by-side
        # blocks. This keeps the parser compatible with both spreadsheet forms.
        vertical_boundaries = [
            row
            for _, row, column in headings
            if column == heading_column and row > heading_row
        ]
        horizontal_boundaries = [
            column
            for _, row, column in headings
            if row == heading_row and column > heading_column
        ]
        row_end = min(vertical_boundaries) - 1 if vertical_boundaries else ws.max_row
        column_end = min(horizontal_boundaries) - 1 if horizontal_boundaries else ws.max_column
        current_event = None

        for row in range(heading_row + 1, row_end + 1):
            event_value = ws.cell(row, heading_column).value
            if _is_default_ad_event_name(event_value):
                current_event = str(event_value).strip()
                records.setdefault((ad_format, current_event), {})

            if not current_event:
                continue

            record = records.setdefault((ad_format, current_event), {})
            for column in range(heading_column + 1, column_end + 1):
                key = _normalize_default_ad_event_param_key(ws.cell(row, column).value)
                if not key:
                    continue
                value = None
                for value_column in range(column + 1, column_end + 1):
                    value_cell = ws.cell(row, value_column)
                    if value_cell.value is not None and str(value_cell.value).strip():
                        value = str(value_cell.value).strip()
                        break
                if value:
                    record[key] = value

    parsed = {ad_format: [] for ad_format in DEFAULT_AD_EVENT_FORMATS}
    for (ad_format, event_name), params in records.items():
        configured_provider = _normalize_default_ad_event_provider(params.get("ad_platform"))
        configured_format = _normalize_default_ad_event_format(params.get("ad_format"))
        if configured_provider != provider or configured_format != ad_format:
            continue
        if event_name not in parsed[ad_format]:
            parsed[ad_format].append(event_name)
    return parsed


def _read_default_ad_event_config(wb):
    config = _new_default_ad_event_config()
    for ws in wb.worksheets:
        provider = _normalize_default_ad_event_provider(ws.title)
        if provider not in DEFAULT_AD_EVENT_PROVIDERS:
            continue
        parsed = _read_default_ad_event_sheet(ws, provider)
        for ad_format, event_names in parsed.items():
            config[ad_format][provider].extend(event_names)
    return config


def _read_default_ad_event_simple_sheet(ws, category):
    """Read AudioAds/InplayAds blocks where each block is identified by platform."""
    expected_providers = set(DEFAULT_AD_EVENT_SIMPLE_PROVIDERS.get(category, ()))
    headings = []

    # The supplied workbook uses merged platform headings such as B2:D2 and
    # F2:H2. Reading only merged headings prevents ad_platform values in the
    # parameter rows from being mistaken for another block.
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row != merged_range.max_row or merged_range.min_row > 5:
            continue
        provider = _normalize_default_ad_event_simple_provider(
            ws.cell(merged_range.min_row, merged_range.min_col).value
        )
        if provider in expected_providers:
            headings.append((provider, merged_range.min_row, merged_range.min_col))

    # Keep compatibility with a simple unmerged layout whose platform names
    # are placed in the first few rows.
    if not headings:
        for row in range(1, min(ws.max_row, 3) + 1):
            for column in range(1, ws.max_column + 1):
                provider = _normalize_default_ad_event_simple_provider(ws.cell(row, column).value)
                if provider in expected_providers:
                    headings.append((provider, row, column))

    headings = sorted(set(headings), key=lambda item: (item[1], item[2], item[0]))
    parsed = {provider: [] for provider in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS.get(category, ())}
    if not headings:
        return parsed

    for provider, heading_row, heading_column in headings:
        vertical_boundaries = [
            row
            for other_provider, row, column in headings
            if column == heading_column and row > heading_row
        ]
        horizontal_boundaries = [
            column
            for other_provider, row, column in headings
            if row == heading_row and column > heading_column
        ]
        row_end = min(vertical_boundaries) - 1 if vertical_boundaries else ws.max_row
        column_end = min(horizontal_boundaries) - 1 if horizontal_boundaries else ws.max_column

        for row in range(heading_row + 1, row_end + 1):
            event_name = ws.cell(row, heading_column).value
            if not _is_default_ad_event_name(event_name):
                continue
            event_name = str(event_name).strip()
            if event_name and event_name not in parsed[provider]:
                parsed[provider].append(event_name)

    return parsed


def _read_default_ad_event_simple_config(wb):
    config = _new_default_ad_event_simple_config()
    for ws in wb.worksheets:
        category = _normalize_default_ad_event_simple_category(ws.title)
        if not category:
            continue
        parsed = _read_default_ad_event_simple_sheet(ws, category)
        for provider, event_names in parsed.items():
            config[category][provider].extend(event_names)
    return config


def _default_ad_event_catalog_payload():
    formats = []
    for ad_format in DEFAULT_AD_EVENT_FORMATS:
        event_names = []
        for provider in DEFAULT_AD_EVENT_PROVIDERS:
            for event_name in default_ad_event_config.get(ad_format, {}).get(provider, []):
                if event_name not in event_names:
                    event_names.append(event_name)
        formats.append({
            "key": ad_format,
            "label": DEFAULT_AD_EVENT_FORMAT_LABELS[ad_format],
            "events": [
                {
                    "name": event_name,
                    "providers": [
                        provider
                        for provider in DEFAULT_AD_EVENT_PROVIDERS
                        if event_name in default_ad_event_config.get(ad_format, {}).get(provider, [])
                    ],
                }
                for event_name in event_names
            ],
        })
    return {
        "providers": [
            {"key": provider, "label": DEFAULT_AD_EVENT_PROVIDER_LABELS[provider]}
            for provider in DEFAULT_AD_EVENT_PROVIDERS
        ],
        "formats": formats,
    }


def _default_ad_event_simple_catalog_payload():
    categories = []
    for category in DEFAULT_AD_EVENT_SIMPLE_CATEGORIES:
        providers = [
            {
                "key": provider,
                "label": DEFAULT_AD_EVENT_SIMPLE_PROVIDER_LABELS[provider],
            }
            for provider in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS[category]
            if default_ad_event_simple_config.get(category, {}).get(provider) is not None
        ]
        event_names = []
        for provider in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS[category]:
            for event_name in default_ad_event_simple_config.get(category, {}).get(provider, []):
                if event_name not in event_names:
                    event_names.append(event_name)
        if not event_names:
            continue
        categories.append({
            "key": category,
            "label": DEFAULT_AD_EVENT_SIMPLE_CATEGORY_LABELS[category],
            "providers": providers,
            "events": [
                {
                    "name": event_name,
                    "providers": [
                        provider
                        for provider in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS[category]
                        if event_name in default_ad_event_simple_config.get(category, {}).get(provider, [])
                    ],
                }
                for event_name in event_names
            ],
        })
    return categories


def _default_ad_event_payload_unlocked():
    format_order = {value: index for index, value in enumerate(DEFAULT_AD_EVENT_FORMATS)}
    provider_order = {value: index for index, value in enumerate(DEFAULT_AD_EVENT_PROVIDERS)}
    category_order = {value: index for index, value in enumerate(DEFAULT_AD_EVENT_SIMPLE_CATEGORIES)}
    hits = []
    for device_id, provider, ad_format, event_name in sorted(
        default_ad_event_hits,
        key=lambda item: (
            format_order.get(item[2], 99),
            provider_order.get(item[1], 99),
            _normalize_default_ad_event_name(item[3]),
            str(item[0]),
        ),
    ):
        hits.append({
            "device_id": device_id,
            "provider": provider,
            "ad_format": ad_format,
            "event_name": event_name,
        })
    simple_hits = []
    for device_id, category, provider, event_name in sorted(
        default_ad_event_simple_hits,
        key=lambda item: (
            category_order.get(item[1], 99),
            DEFAULT_AD_EVENT_SIMPLE_PROVIDERS.get(item[1], ()).index(item[2])
            if item[2] in DEFAULT_AD_EVENT_SIMPLE_PROVIDERS.get(item[1], ()) else 99,
            _normalize_default_ad_event_name(item[3]),
            str(item[0]),
        ),
    ):
        simple_hits.append({
            "device_id": device_id,
            "category": category,
            "provider": provider,
            "event_name": event_name,
        })
    payload = _default_ad_event_catalog_payload()
    payload["hits"] = hits
    payload["simple_categories"] = _default_ad_event_simple_catalog_payload()
    payload["simple_hits"] = simple_hits
    return payload


def _default_ad_event_payload():
    with lock:
        return _default_ad_event_payload_unlocked()


def _clear_default_ad_event_history():
    """Clear only runtime ad-event hits; keep the imported profile catalog intact."""
    with lock:
        default_ad_event_hits.clear()
        default_ad_event_simple_hits.clear()


atexit.register(_clear_default_ad_event_history)


def _record_default_ad_event_hit(event_name, actual_params, device_id):
    if is_paused or not isinstance(actual_params, dict):
        return
    provider = _normalize_default_ad_event_provider(actual_params.get("ad_platform"))
    ad_format = _normalize_default_ad_event_format(actual_params.get("ad_format"))
    simple_provider = _normalize_default_ad_event_simple_provider(actual_params.get("ad_platform"))
    event_key = _normalize_default_ad_event_name(event_name)
    if not event_key:
        return

    with lock:
        # Keep coverage independent of the Default Ad Events tab connection.
        # iOS can emit TrackingService events before the tab is opened; the
        # next tab sync should still show those matches.
        changed = False

        if provider and ad_format:
            configured_event = next(
                (
                    configured
                    for configured in default_ad_event_config.get(ad_format, {}).get(provider, [])
                    if _normalize_default_ad_event_name(configured) == event_key
                ),
                None,
            )
            if configured_event:
                hit = (str(device_id or ""), provider, ad_format, configured_event)
                if hit not in default_ad_event_hits:
                    default_ad_event_hits.add(hit)
                    changed = True

        if simple_provider:
            for category in DEFAULT_AD_EVENT_SIMPLE_CATEGORIES:
                configured_event = next(
                    (
                        configured
                        for configured in default_ad_event_simple_config.get(category, {}).get(simple_provider, [])
                        if _normalize_default_ad_event_name(configured) == event_key
                    ),
                    None,
                )
                if not configured_event:
                    continue
                simple_hit = (str(device_id or ""), category, simple_provider, configured_event)
                if simple_hit in default_ad_event_simple_hits:
                    continue
                default_ad_event_simple_hits.add(simple_hit)
                changed = True

        if not changed:
            return
        payload = _default_ad_event_payload_unlocked()
    socketio.emit("update_default_ad_events", payload)


def _normalize_adrevenue_sheet_key(name):
    text = re.sub(r'[^a-z0-9]+', '', str(name or "").lower())
    aliases = {
        "appmetrica": {"appmetrica", "adrevenueappmetrica", "appmetricaadrevenue"},
        "appsflyer": {"appsflyer", "adrevenueappsflyer", "appsflyeradrevenue"},
        "adjust": {"adjust", "adrevenueadjust", "adjustadrevenue"},
        "all": {"adrevenue", "all", "default", "common", "shared"},
    }
    for key, values in aliases.items():
        if text in values:
            return key
    return text


def _read_adrevenue_sheet(ws):
    """Read AdRevenue params from a dedicated sheet with source-specific columns."""
    params_by_source = {"appmetrica": [], "appsflyer": [], "adjust": []}
    source_cols = {"appmetrica": None, "appsflyer": None, "adjust": None}
    header_row = None

    # The adrevenue sheet can have headers on row 1 or row 2.
    # Find the first row near the top that declares source columns.
    for row in range(1, min(ws.max_row, 5) + 1):
        candidate_cols = {"appmetrica": None, "appsflyer": None, "adjust": None}
        for col in range(1, ws.max_column + 1):
            header = _normalize_adrevenue_sheet_key(ws.cell(row, col).value)
            if header in candidate_cols and candidate_cols[header] is None:
                candidate_cols[header] = col
        if any(candidate_cols.values()):
            header_row = row
            source_cols = candidate_cols
            break

    if header_row is None:
        return params_by_source

    for row in range(header_row + 1, ws.max_row + 1):
        for source, col in source_cols.items():
            if not col:
                continue
            raw = ws.cell(row, col).value
            if raw is None:
                continue
            value = str(raw).strip()
            if not value:
                continue
            params_by_source[source].append(value)

    return params_by_source


def load_default_params_config():
    """Load default params, event-specific params, and AdRevenue params from XLSX."""
    global default_params, event_specific_params, active_profile_game_name
    global adrevenue_default_params, adrevenue_source_params
    global default_ad_event_config
    global default_ad_event_simple_config
    path = active_profile_path or DEFAULT_PARAMS_XLSX
    if not path or not os.path.exists(path):
        print(f"INFO: Default params sheet not found: {path}")
        default_params = []
        event_specific_params = {}
        active_profile_game_name = ""
        adrevenue_default_params = []
        adrevenue_source_params = {}
        with lock:
            default_ad_event_config = _new_default_ad_event_config()
            default_ad_event_simple_config = _new_default_ad_event_simple_config()
            default_ad_event_hits.clear()
            default_ad_event_simple_hits.clear()
        return
    try:
        wb = load_workbook(path)
        profile_sheet = wb.active
        if _normalize_default_ad_event_provider(profile_sheet.title):
            profile_sheet = next(
                (
                    ws for ws in wb.worksheets
                    if not _normalize_default_ad_event_provider(ws.title)
                ),
                profile_sheet,
            )
        active_profile_game_name, default_params, event_specific_params = _read_profile_sheet(profile_sheet, allow_default_fill=True)
        parsed_ad_event_config = _read_default_ad_event_config(wb)
        parsed_ad_event_simple_config = _read_default_ad_event_simple_config(wb)
        with lock:
            default_ad_event_config = parsed_ad_event_config
            default_ad_event_simple_config = parsed_ad_event_simple_config
            default_ad_event_hits.clear()
            default_ad_event_simple_hits.clear()

        revenue_defaults = []
        revenue_specific = {}
        revenue_sheet = None
        for ws in wb.worksheets:
            title = _normalize_adrevenue_sheet_key(ws.title)
            if title == "adrevenue":
                revenue_sheet = ws
                break
        if revenue_sheet is None:
            for ws in wb.worksheets[1:]:
                title = _normalize_adrevenue_sheet_key(ws.title)
                if "revenue" in title or title in {"appmetrica", "appsflyer", "all"}:
                    revenue_sheet = ws
                    break

        if revenue_sheet:
            revenue_specific = _read_adrevenue_sheet(revenue_sheet)

        adrevenue_default_params = revenue_defaults
        adrevenue_source_params = revenue_specific
        print(
            f"INFO: Loaded default params: {len(default_params)}; events: {len(event_specific_params)}; "
            f"adrevenue defaults: {len(adrevenue_default_params)}; adrevenue sources: {len(adrevenue_source_params)}"
        )
    except Exception as e:
        print(f"ERROR: Failed to load default params sheet: {e}")
        default_params = []
        event_specific_params = {}
        active_profile_game_name = ""
        adrevenue_default_params = []
        adrevenue_source_params = {}
        with lock:
            default_ad_event_config = _new_default_ad_event_config()
            default_ad_event_simple_config = _new_default_ad_event_simple_config()
            default_ad_event_hits.clear()
            default_ad_event_simple_hits.clear()


def _sanitize_profile_filename(filename):
    name = os.path.basename((filename or "").strip())
    name = name.replace("\\", "_").replace("/", "_")
    if not name.lower().endswith(".xlsx"):
        raise ValueError("Only .xlsx files are supported")
    return name


def _sanitize_ipa_filename(filename):
    name = os.path.basename((filename or "").strip())
    name = name.replace("\\", "_").replace("/", "_")
    if not name.lower().endswith(".ipa"):
        raise ValueError("Only .ipa files are supported")
    return name


def _extract_brightsdk_version_from_ipa(ipa_path):
    if not ipa_path or not os.path.exists(ipa_path):
        return {"app_name": "", "brightsdk_version": "", "skadnetwork_identifiers": []}
    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="eventinspector_ipa_")
        with zipfile.ZipFile(ipa_path, "r") as zf:
            zf.extractall(tmp_dir)
        payload_dir = os.path.join(tmp_dir, "Payload")
        if not os.path.isdir(payload_dir):
            return {"app_name": "", "brightsdk_version": "", "skadnetwork_identifiers": []}
        for app_name in os.listdir(payload_dir):
            app_dir = os.path.join(payload_dir, app_name)
            if not app_name.endswith(".app") or not os.path.isdir(app_dir):
                continue
            app_display_name = ""
            app_bundle_name = ""
            skadnetwork_identifiers = []
            app_plist_path = os.path.join(app_dir, "Info.plist")
            if os.path.exists(app_plist_path):
                with open(app_plist_path, "rb") as f:
                    app_plist = plistlib.load(f)
                app_display_name = str(app_plist.get("CFBundleDisplayName") or "").strip()
                app_bundle_name = str(app_plist.get("CFBundleName") or "").strip()
                # SKAdNetworkItems belongs to the app's Info.plist. Do not
                # read the separate BrightSDK framework plist for this list.
                for item in app_plist.get("SKAdNetworkItems") or []:
                    if not isinstance(item, dict):
                        continue
                    identifier = str(item.get("SKAdNetworkIdentifier") or "").strip()
                    if identifier:
                        skadnetwork_identifiers.append(identifier)
            for root, _dirs, files in os.walk(app_dir):
                if os.path.basename(root) == "brdsdk.framework":
                    plist_path = os.path.join(root, "Info.plist")
                    if not os.path.exists(plist_path):
                        continue
                    with open(plist_path, "rb") as f:
                        plist_data = plistlib.load(f)
                    version = str(plist_data.get("CFBundleShortVersionString") or "").strip()
                    return {
                        "app_name": app_display_name or app_bundle_name or app_name[:-4],
                        "brightsdk_version": version,
                        "skadnetwork_identifiers": skadnetwork_identifiers,
                    }
            # Keep app metadata useful even when the IPA has no BrightSDK
            # framework, while still reporting the app plist SKAdNetwork list.
            if app_display_name or app_bundle_name or skadnetwork_identifiers:
                return {
                    "app_name": app_display_name or app_bundle_name or app_name[:-4],
                    "brightsdk_version": "",
                    "skadnetwork_identifiers": skadnetwork_identifiers,
                }
        return {"app_name": "", "brightsdk_version": "", "skadnetwork_identifiers": []}
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _list_profile_names():
    try:
        files = [
            p.name for p in Path(PROFILE_DIR).glob("*.xlsx")
            if p.is_file()
        ]
        return sorted(files, key=lambda x: x.lower())
    except Exception:
        return []


def _ensure_default_profile_seed():
    source = DEFAULT_PARAMS_XLSX
    if not source or not os.path.exists(source):
        return
    target = os.path.join(PROFILE_DIR, os.path.basename(source))
    if os.path.abspath(source) == os.path.abspath(target):
        return
    if not os.path.exists(target):
        try:
            shutil.copyfile(source, target)
        except Exception as e:
            print(f"WARNING: Failed to seed default profile: {e}")


def _set_active_profile(profile_name=None):
    global active_profile_name, active_profile_path, profile_revision
    _ensure_default_profile_seed()
    names = _list_profile_names()
    if not names:
        active_profile_name = None
        active_profile_path = None
        load_default_params_config()
        profile_revision += 1
        return False

    selected = None
    if profile_name:
        clean = _sanitize_profile_filename(profile_name)
        if clean in names:
            selected = clean

    if not selected:
        preferred = os.path.basename(DEFAULT_PARAMS_XLSX) if DEFAULT_PARAMS_XLSX else None
        if preferred in names:
            selected = preferred
        else:
            selected = names[0]

    active_profile_name = selected
    active_profile_path = os.path.join(PROFILE_DIR, selected)
    load_default_params_config()
    profile_revision += 1
    return True


def _profile_payload():
    return {
        "profiles": _list_profile_names(),
        "current_profile": active_profile_name,
        "profile_revision": profile_revision,
        "game_name": active_profile_game_name,
        "profile_dir": PROFILE_DIR,
        "default_event_names": sorted(event_specific_params.keys()),
        "default_ad_event_data": _default_ad_event_payload(),
    }


def _emit_profile_runtime_refresh():
    """Re-render every tab that derives results from the active profile."""
    with lock:
        validator_results.clear(); validator_pending.clear()
        _clear_record_dedup("validator")
    _apply_specific_filter_and_emit()
    _apply_adrevenue_filter_and_emit()
    socketio.emit('update_validator_table', [])
    socketio.emit('profile_updated', _profile_payload())

def _levenshtein_distance_limit(a, b, limit=2):
    """Compute Levenshtein distance with early exit if > limit."""
    if a == b:
        return 0
    if not a or not b:
        return max(len(a), len(b))
    if abs(len(a) - len(b)) > limit:
        return limit + 1

    # Initialize previous row
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        curr = [i]
        min_in_row = curr[0]
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            curr_val = min(
                prev[j] + 1,      # deletion
                curr[j - 1] + 1,  # insertion
                prev[j - 1] + cost  # substitution
            )
            curr.append(curr_val)
            if curr_val < min_in_row:
                min_in_row = curr_val
        if min_in_row > limit:
            return limit + 1
        prev = curr
    return prev[-1]

# --- GIAO DIỆN WEB (HTML/JS) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi" data-ui-zoom="{{ ui_zoom }}" style="font-size: {{ ui_zoom_px }}px;">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="icon" href="data:,"> <!-- Fix lỗi Favicon 404 -->
    <title>Event Inspector v2.5.0(85)</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/socket.io/4.7.4/socket.io.js"></script>
    <style>
        body { font-family: 'Inter', sans-serif; }
        .log-cell { max-width: 500px; word-wrap: break-word; font-family: monospace; font-size: 0.75rem; color: #6b7280; }
        .message-cell { white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        #packageLogTable { table-layout: fixed; width: 100%; }
        #packageLogTable col.col-time { width: 110px; }
        #packageLogTable col.col-tag { width: 90px; }
        .details-cell { font-family: monospace; font-size: 0.8rem; line-height: 1.4; min-width: 260px; max-width: 640px; white-space: normal; overflow-wrap: anywhere; word-break: break-word; }
        .tag-cell { max-width: 90px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .tag-header { max-width: 90px; }
        .time-cell { white-space: nowrap; width: 1%; max-width: 110px; font-size: 0.75rem; }
        .time-header { white-space: nowrap; width: 1%; max-width: 110px; }
        .resizable { position: relative; }
        .resizer {
            position: absolute;
            right: -3px;
            top: 0;
            width: 12px;
            height: 100%;
            cursor: col-resize;
            user-select: none;
            z-index: 20;
        }
        .resizer:hover { background: rgba(59, 130, 246, 0.15); }
        #packageLogTableBody tr.selected, #packageHistoryTableBody tr.selected { background-color: #bfdbfe !important; }
        /* Live Package Log: virtualized, one line per row (double-click shows the full record). */
        #packageLogTableLive { table-layout: fixed; width: 100%; }
        #packageLogTableLive th.col-time { width: 92px; }
        #packageLogTableLive th.col-tag { width: 140px; }
        #packageLogTableLive thead th { padding-top: 4px; padding-bottom: 4px; font-size: 11px; }
        /* Compact rows: 10px text, 18px rows (the virtual scroller measures the real height). */
        #packageLogTableBody tr.package-log-row { height: 18px; }
        #packageLogTableBody tr.package-log-row td {
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
            padding-top: 0; padding-bottom: 0; height: 18px; line-height: 18px; font-size: 10px;
            border-bottom: 1px solid #f3f4f6; box-sizing: border-box;
        }
        #packageLogTableBody tr.package-log-spacer td { padding: 0; border: 0; }
        #packageNewRowsBtn { position: absolute; right: 24px; bottom: 16px; z-index: 20; }
        .resizer.disabled { cursor: not-allowed; background: transparent; }
        .details-cell pre { margin: 0; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .adrevenue-panel { border: 1px solid #e5e7eb; background: #f9fafb; border-radius: 0.5rem; padding: 0.5rem; }
        .adrevenue-details-panel { max-height: none; overflow: visible; }
        .adrevenue-details-panel pre { margin: 0 !important; background: transparent !important; border: 0 !important; padding: 0 !important; border-radius: 0 !important; overflow: visible !important; }
        .adrevenue-raw-panel { height: 16rem; overflow: auto; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .price-rotation-details,
        .price-rotation-details * {
            user-select: text !important;
            -webkit-user-select: text !important;
            cursor: text;
        }
        .installation-id-value {
            user-select: text !important;
            -webkit-user-select: text !important;
            cursor: text;
        }
        #logDetailContent, #logDetailContent * { user-select: text; -webkit-user-select: text; }
        @keyframes pulse { 50% { opacity: .6; } }
        .animate-pulse-green { animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite; }
        .animate-record { animation: pulse 1.5s infinite; }
        .tab-btn { transition: all 0.2s ease-in-out; }
        .tab-btn.active { border-color: #4f46e5; color: #4f46e5; background-color: #eef2ff; }
        .log-row.selected { background-color: #dbeafe; }
        .log-cell,
        .details-cell,
        .details-cell *,
        #loadAdsTableBody td,
        #loadAdsExtTableBody td,
        #validatorTableBody td,
        #specificEventTableBody td,
        #adRevenueTableBody td,
        #callbackAdTableBody td,
        #sdkCheckTableBody td {
            user-select: text;
            -webkit-user-select: text;
        }
        
        /* Custom scrollbar for pre blocks */
        pre::-webkit-scrollbar { height: 6px; width: 6px; }
        pre::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        pre::-webkit-scrollbar-track { background: #f1f5f9; }

        /* Full-window layout: the visible tab fills the space below the tab bar */
        #tabScrollArea { display: flex; flex-direction: column; }
        #tabScrollArea > [id^="tabContent"]:not(.hidden) { flex: 1 0 auto; display: flex; flex-direction: column; }
        #tabScrollArea > [id^="tabContent"]:not(.hidden) > div:last-child { flex: 1 0 auto; }
    </style>
</head>
<body class="bg-gray-100 text-gray-800 h-screen overflow-hidden">

    <div id="appShell" class="w-full max-w-none p-3 sm:p-4 lg:p-6 h-full flex flex-col">
        <!-- HEADER -->
        <div id="appHeader" class="bg-white rounded-xl shadow-md p-4 mb-4 flex-shrink-0{% if ui_hide_top %} hidden{% endif %}">
            <div class="flex justify-between items-center flex-wrap gap-4">
                <div class="flex items-center gap-3">
                    <div>
                        <div class="flex items-center gap-2.5">
                            <h1 class="text-xl font-bold text-gray-700">Event Inspector</h1>
                            <span class="text-xs font-semibold bg-indigo-100 text-indigo-700 px-2 py-1 rounded-full">v2.5.0(85)</span>
                        </div>
                        <p class="text-sm text-gray-500">Integrates Load Ads & Event Validation.</p>
                    </div>
                    <div class="flex items-center gap-2">
                    <button id="restartAppBtn" class="bg-blue-500 hover:bg-blue-600 text-white text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm">Check Update</button>
                    <button id="manualRestartBtn" class="bg-slate-500 hover:bg-slate-600 text-white text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm">Restart</button>
                    <div id="uiZoomControls" class="flex items-center border border-slate-300 rounded-lg overflow-hidden shadow-sm bg-white text-slate-700 text-sm font-semibold" title="Text size (Ctrl/Cmd − / = / 0)">
                        <button id="uiZoomOut" type="button" class="px-2.5 py-2 hover:bg-slate-100" title="Smaller text (Ctrl/Cmd −)">A−</button>
                        <button id="uiZoomReset" type="button" class="px-2 py-2 border-x border-slate-300 hover:bg-slate-100 min-w-[3.5rem]" title="Reset to the default size (Ctrl/Cmd 0)">{{ (ui_zoom * 100)|round|int }}%</button>
                        <button id="uiZoomIn" type="button" class="px-2.5 py-2 hover:bg-slate-100" title="Larger text (Ctrl/Cmd =)">A+</button>
                    </div>
                    <button id="platformBtn" class="bg-white hover:bg-gray-50 text-slate-700 border border-slate-300 text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm">Platform: Android</button>
                    <div id="installationIdPanel" class="min-w-[300px] max-w-[420px] bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 shadow-sm">
                        <div class="flex items-start justify-between gap-3">
                            <div class="min-w-0 flex-1">
                                <div id="installationIdGame" class="text-xs font-semibold text-slate-700 truncate">-</div>
                                <label class="flex items-center gap-1 mt-1 min-w-0 text-[11px] text-slate-500 font-mono">
                                    <span class="shrink-0">Installation ID:</span>
                                    <input id="installationIdValue" type="text" readonly value="" placeholder="-" aria-label="Installation ID" class="installation-id-value min-w-0 flex-1 bg-transparent border-0 p-0 text-[11px] text-slate-500 font-mono leading-4 focus:outline-none focus:ring-0">
                                </label>
                            </div>
                            <button id="copyInstallationIdBtn" type="button" class="shrink-0 text-[11px] font-semibold px-2 py-1 rounded border border-slate-300 bg-white hover:bg-slate-100 text-slate-700">Copy</button>
                        </div>
                    </div>
                </div>
                </div>
                <div class="flex items-center gap-4">
                    <div class="flex items-center gap-2">
                        <button id="pauseBtn" class="text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm bg-yellow-500 hover:bg-yellow-600 text-white">Pause</button>
                        <button id="clearAllBtn" class="text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm bg-red-500 hover:bg-red-600 text-white">Clear All</button>
                    </div>
                    <div class="text-right p-2.5 rounded-lg bg-gray-50 border min-w-[360px] max-w-[520px]">
                        <p class="text-sm font-semibold text-gray-700 mb-1 border-b pb-1">Connected Devices:</p>
                        <div id="deviceList" class="text-sm text-gray-600 whitespace-nowrap overflow-hidden text-ellipsis">
                             <p class="text-orange-500">WAITING...</p>
                        </div>
                    </div>
                     <div>
                        <label for="deviceFilter" class="block text-xs font-medium text-gray-700">Filter by Device:</label>
                        <select id="deviceFilter" class="mt-1 block w-full pl-3 pr-10 py-1.5 text-sm border-gray-300 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 rounded-md">
                            <option value="all">All Devices</option>
                        </select>
                    </div>
                </div>
            </div>
        </div>

        <div id="platformModal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/40 px-4">
            <div class="bg-white rounded-xl shadow-xl border w-full max-w-md p-5">
                <h2 class="text-lg font-bold text-slate-800 mb-2">Select Platform</h2>
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-4">
                    <button data-platform-choice="android" class="platform-choice-btn rounded-lg border border-slate-200 bg-slate-50 hover:bg-indigo-50 hover:border-indigo-300 text-slate-800 font-semibold py-4 px-4 transition-colors">Android</button>
                    <button data-platform-choice="ios" class="platform-choice-btn rounded-lg border border-slate-200 bg-slate-50 hover:bg-indigo-50 hover:border-indigo-300 text-slate-800 font-semibold py-4 px-4 transition-colors">iOS</button>
                </div>
            </div>
        </div>

        <!-- TABS & ACTIONS -->
        <div class="mb-3 flex-shrink-0">
            <div class="flex justify-between items-end border-b border-gray-200">
                <div class="flex flex-wrap">
                    <button id="tabBtnLoadAdsExt" class="tab-btn active text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('LoadAdsExt')">Load Ads</button>
                    
                    <button id="tabBtnValidator" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('Validator')">Default Events/Params</button>
                    <button id="tabBtnDefaultAdEvents" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('DefaultAdEvents')">Default Ad Events</button>
                    <button id="tabBtnSpecific" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('Specific')">Specific Validator</button>
                    <button id="tabBtnAdRevenue" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('AdRevenue')">Revenue</button>
                    <button id="tabBtnCallbackAd" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('CallbackAd')">CallBack & Ads</button>
                    <button id="tabBtnPriceRotation" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('PriceRotation')">Bidding</button>
                    <button id="tabBtnSdkCheck" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('SdkCheck')">SDK Check</button>
                    <button id="tabBtnBrightSDK" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('BrightSDK')">Bright/SKAds</button>
                    <button id="tabBtnPackage" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('Package')">Package Logcat</button>
                    <button id="tabBtnServicesChecker" class="tab-btn text-sm font-semibold py-2 px-4 -mb-px border-b-2 border-transparent" onclick="switchTab('ServicesChecker')">Services Checker</button>
                </div>
                <button id="toggleTopBtn" type="button" class="shrink-0 mb-1 ml-2 text-xs font-semibold py-1.5 px-3 rounded-md border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 shadow-sm" title="Hide or show the top bar to give the tables more room">{% if ui_hide_top %}Show Top{% else %}Hide Top{% endif %}</button>
            </div>
        </div>

        <!-- CONTENT -->
        <div id="tabScrollArea" class="flex-grow overflow-auto">

            <!-- TAB 1: Load Ads -->
            <div id="tabContentLoadAds" class="hidden">
                 <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="flex items-center gap-2 bg-gray-50 p-2.5 rounded-lg border mb-3">
                        <span class="text-sm font-semibold text-gray-700">Record Load Ads:</span>
                        <input type="text" id="sheetName_LoadAds" placeholder="Tên Sheet..." class="border p-2 rounded text-sm w-48 outline-none">
                        <button id="btnRecord_LoadAds" onclick="toggleRecord('LoadAds')" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 px-3 rounded shadow text-xs">Start Record</button>
                    </div>

                    <div class="overflow-x-auto">
                        <table id="packageLogTable" class="min-w-full bg-white">
                            <colgroup>
                                <col style="width:88px">
                                <col style="width:100px">
                                <col style="width:112px">
                                <col style="width:96px">
                                <col>
                            </colgroup>
                            <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-2 border-b">Provider</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Ad_network</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Format</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                </tr>
                            </thead>
                            <tbody id="loadAdsTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 2: Load Ads -->
            <div id="tabContentLoadAdsExt">
                 <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="flex items-center gap-2 bg-gray-50 p-2.5 rounded-lg border mb-3">
                        <span class="text-sm font-semibold text-gray-700">Record Load Ads:</span>
                        <input type="text" id="sheetName_LoadAdsExt" placeholder="Tên Sheet..." class="border p-2 rounded text-sm w-48 outline-none">
                        <button id="btnRecord_LoadAdsExt" onclick="toggleRecord('LoadAdsExt')" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 px-3 rounded shadow text-xs">Start Record</button>
                    </div>

                    <div class="overflow-x-auto">
                        <table class="min-w-full w-full bg-white table-fixed">
                            <colgroup>
                                <col style="width:88px">
                                <col style="width:100px">
                                <col style="width:96px">
                                <col style="width:112px">
                                <col>
                            </colgroup>
                            <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-2 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-2 border-b">Provider</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-2 border-b">Ad_network</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-2 border-b">Format</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                </tr>
                            </thead>
                            <tbody id="loadAdsExtTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 3: Validator -->
            <div id="tabContentValidator" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4 mb-4">
                    <div class="grid grid-cols-1 lg:grid-cols-3 gap-4">
                        <div class="lg:col-span-1">
                            <div class="space-y-3">
                                <div>
                                    <div class="flex items-center gap-2 mb-1.5">
                                        <label for="profileSelect" class="text-[11px] font-medium text-gray-700">Game Profile:</label>
                                        <p id="profileGameText" class="text-[11px] font-medium text-indigo-700"></p>
                                    </div>
                                    <div class="space-y-2">
                                        <select id="profileSelect" class="w-full h-9 px-3 border rounded-md shadow-sm text-xs"></select>
                                        <input type="file" id="profileFileInput" accept=".xlsx" class="hidden">
                                        <div class="flex flex-wrap items-center gap-3">
                                            <button id="importProfileBtn" class="bg-slate-700 hover:bg-slate-800 text-white font-medium text-xs px-3 rounded-lg h-9 min-w-[116px]">Import Profile</button>
                                            <button id="reloadProfileBtn" class="bg-slate-200 hover:bg-slate-300 text-gray-800 font-medium text-xs px-3 rounded-lg h-9 min-w-[116px]">Reload Profile</button>
                                            <button id="openProfileFolderBtn" class="bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium text-xs px-3 rounded-lg h-9 min-w-[72px]" title="Open profile folder">Open Folder</button>
                                        </div>
                                    </div>
                                </div>
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                    <div>
                                        <label for="validatorEventFilterInput" class="block text-[11px] font-medium text-gray-700 mb-1">Filter by Event Name:</label>
                                        <input type="text" id="validatorEventFilterInput" class="w-full h-9 px-3 border rounded-md shadow-sm text-xs" placeholder="Type to filter events...">
                                    </div>
                                    <div>
                                        <label for="validatorRawFilterInput" class="block text-[11px] font-medium text-gray-700 mb-1">Filter by Raw Log:</label>
                                        <input type="text" id="validatorRawFilterInput" class="w-full h-9 px-3 border rounded-md shadow-sm text-xs" placeholder="Search raw log...">
                                    </div>
                                </div>
                                <div class="flex items-center gap-3 pt-1">
                                    <button id="startValidationBtn" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs px-4 rounded-lg h-9">Start Checking</button>
                                    <button id="clearValidatorFilterBtn" class="bg-gray-200 hover:bg-gray-300 text-gray-800 font-medium text-xs px-4 rounded-lg h-9">Clear Filter</button>
                                    <div class="flex items-center gap-3 text-[11px] text-gray-700 ml-1">
                                        <label class="inline-flex items-center gap-1.5"><input type="radio" name="validatorSourceFilter" value="all" checked> <span>All</span></label>
                                        <label class="inline-flex items-center gap-1.5" data-ios-hidden-source="appmetrica"><input type="radio" name="validatorSourceFilter" value="appmetrica"> <span>Appmetrica</span></label>
                                        <label class="inline-flex items-center gap-1.5"><input type="radio" name="validatorSourceFilter" value="firebase"> <span>Firebase</span></label>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="lg:col-span-2">
                            <div class="text-xs font-semibold text-gray-700 mb-2">Default Events Status</div>
                            <div id="defaultEventStatusList" class="flex flex-wrap gap-1.5 text-xs"></div>
                        </div>
                    </div>
                </div>
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="overflow-x-auto">
                        <table class="min-w-full bg-white">
                           <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Status</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Event Name</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Details</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Action</th>
                                </tr>
                            </thead>
                            <tbody id="validatorTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 3A: Default Ad Events -->
            <div id="tabContentDefaultAdEvents" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4 mb-4">
                    <div class="flex flex-wrap items-center gap-3">
                        <label for="defaultAdEventProfileSelect" class="text-xs font-medium text-gray-700">Default File:</label>
                        <select id="defaultAdEventProfileSelect" class="h-9 min-w-[260px] px-3 border rounded-md shadow-sm text-xs"></select>
                        <span id="defaultAdEventProfileGameText" class="text-xs font-medium text-indigo-700"></span>
                        <span class="text-xs text-gray-500">✓ = khớp ad_platform + ad_format; các bảng platform riêng khớp ad_platform; ô trống = chưa thấy log</span>
                    </div>
                </div>
                <div id="defaultAdEventMatrix" class="space-y-4"></div>
            </div>

            <!-- TAB 3A: BrightSDK -->
            <div id="tabContentBrightSDK" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="flex flex-col gap-3">
                        <div class="flex flex-wrap items-center gap-2 bg-indigo-50 p-2.5 rounded-lg border border-indigo-100">
                            <label for="brightSkadPresetSelect" class="text-sm font-semibold text-gray-700 whitespace-nowrap">Bright/SKAds preset:</label>
                            <select id="brightSkadPresetSelect" class="h-9 min-w-[220px] px-3 border rounded-md shadow-sm text-xs bg-white">
                                <option value="">Loading presets...</option>
                            </select>
                            <button id="reloadBrightSkadPresetsBtn" type="button" class="h-9 shrink-0 rounded border border-indigo-200 bg-white px-3 text-xs font-semibold text-indigo-700 hover:bg-indigo-100">Reload</button>
                            <span id="brightSkadPresetStatus" class="text-xs text-slate-500">C-192-iOS sẽ được chọn mặc định.</span>
                        </div>
                        <div class="flex items-center gap-2 bg-gray-50 p-2.5 rounded-lg border">
                            <span class="text-sm font-semibold text-gray-700 whitespace-nowrap">Upload IPA:</span>
                            <input type="file" id="brightSdkFileInput" accept=".ipa" class="hidden">
                            <button id="importBrightSdkBtn" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 px-3 rounded shadow text-xs">Choose IPA</button>
                            <button id="clearBrightSdkBtn" class="bg-slate-200 hover:bg-slate-300 text-gray-800 font-semibold py-2 px-3 rounded shadow text-xs">Clear</button>
                        </div>
                        <div class="bg-gray-50 p-2.5 rounded-lg border text-sm text-slate-600 space-y-1">
                            <div>
                                <span class="font-semibold">App Name:</span>
                                <span id="brightSdkAppNameValue" class="font-mono text-slate-800 ml-1">Not Found</span>
                            </div>
                            <div>
                                <span class="font-semibold">BrightSDK version:</span>
                                <span id="brightSdkVersionValue" class="font-mono text-slate-800 ml-1">Not Found</span>
                            </div>
                        </div>
                        <div class="bg-white rounded-lg border border-slate-200 overflow-hidden">
                            <div class="flex flex-wrap items-center justify-between gap-2 bg-slate-50 px-3 py-2 border-b">
                                <div class="font-semibold text-sm text-slate-700">Bright/SKAds comparison</div>
                                <div id="brightSkadCompareSummary" class="text-xs text-slate-500">Upload IPA để bắt đầu compare.</div>
                            </div>
                            <div id="brightSkadExtraWarning" class="hidden px-3 py-2 text-xs text-amber-800 bg-amber-50 border-b border-amber-200"></div>
                            <div class="overflow-auto max-h-[32rem]">
                                <table class="min-w-full text-xs">
                                    <thead class="bg-gray-50 sticky top-0 z-10">
                                        <tr>
                                            <th class="text-center font-semibold text-gray-600 py-2 px-3 border-b">Status</th>
                                            <th class="text-left font-semibold text-gray-600 py-2 px-3 border-b">Item</th>
                                            <th class="text-left font-semibold text-gray-600 py-2 px-3 border-b">Actual</th>
                                            <th class="text-left font-semibold text-gray-600 py-2 px-3 border-b">Expected</th>
                                        </tr>
                                    </thead>
                                    <tbody id="brightSkadCompareBody">
                                        <tr><td colspan="4" class="py-3 px-3 text-center text-slate-500 italic">Chưa có IPA để compare.</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- TAB 4: Specific -->
            <div id="tabContentSpecific" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4 mb-4">
                    <div class="grid grid-cols-1 lg:grid-cols-[minmax(0,1fr)_150px_minmax(0,1fr)] gap-4 items-start">
                         <div>
                            <label for="specificEventInput" class="block text-xs font-medium text-gray-700 mb-1">Filter by Event Names:</label>
                            <textarea id="specificEventInput" class="w-full h-[4.25rem] resize-none p-2 border rounded-md shadow-sm" placeholder="Leave empty to show all events..."></textarea>
                         </div>
                         <div>
                            <label class="block text-xs font-medium text-gray-700 mb-1">Source Filter:</label>
                            <div class="flex flex-col gap-2 text-xs text-gray-700 pt-1">
                                <label class="inline-flex items-center gap-1.5">
                                    <input type="radio" name="specificSourceFilter" value="all" checked>
                                    <span>All</span>
                                </label>
                                <label class="inline-flex items-center gap-1.5">
                                    <input type="radio" name="specificSourceFilter" value="firebase">
                                    <span>Firebase</span>
                                </label>
                                <label class="inline-flex items-center gap-1.5">
                                    <input type="radio" name="specificSourceFilter" value="appmetrica">
                                    <span>Appmetrica</span>
                                </label>
                            </div>
                         </div>
                         <div>
                            <label for="specificParamInput" class="block text-xs font-medium text-gray-700 mb-1">Filter by Text:</label>
                            <textarea id="specificParamInput" class="w-full h-[4.25rem] resize-none p-2 border rounded-md shadow-sm" placeholder="Leave empty to show all raw logs..."></textarea>
                         </div>
                    </div>
                </div>
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="overflow-x-auto">
                        <table class="min-w-full bg-white">
                           <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Status</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Event Name</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Details</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Action</th>
                                </tr>
                            </thead>
                            <tbody id="specificEventTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 5: Revenue -->
            <div id="tabContentAdRevenue" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4 mb-4">
                     <div class="grid grid-cols-1 md:grid-cols-[minmax(260px,380px)_minmax(0,1fr)] gap-4 items-start">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <label for="adRevenueProfileSelect" class="text-xs font-medium text-gray-700 whitespace-nowrap">Game Profile:</label>
                                <p id="adRevenueProfileGameText" class="text-xs font-medium text-indigo-700"></p>
                            </div>
                            <select id="adRevenueProfileSelect" class="w-full h-10 px-3 border rounded-md shadow-sm text-sm"></select>
                        </div>
                        <div class="space-y-3">
                            <div>
                                <label class="block text-xs font-medium text-gray-700 mb-1">Source Filter:</label>
                                <div class="flex flex-wrap items-center gap-x-4 gap-y-2 text-sm">
                                    <label class="inline-flex items-center whitespace-nowrap">
                                        <input name="adRevenueSourceFilter" type="radio" value="all" checked class="h-4 w-4 text-indigo-600 border-gray-300 focus:ring-indigo-500">
                                        <span class="ml-2 text-sm text-gray-900">All</span>
                                    </label>
                                    <label class="inline-flex items-center whitespace-nowrap">
                                        <input name="adRevenueSourceFilter" type="radio" value="appmetrica" class="h-4 w-4 text-indigo-600 border-gray-300 focus:ring-indigo-500">
                                        <span class="ml-2 text-sm text-gray-900">Appmetrica</span>
                                    </label>
                                    <label class="inline-flex items-center whitespace-nowrap">
                                        <input name="adRevenueSourceFilter" type="radio" value="appsflyer" class="h-4 w-4 text-indigo-600 border-gray-300 focus:ring-indigo-500">
                                        <span class="ml-2 text-sm text-gray-900">Appsflyer</span>
                                    </label>
                                    <label class="inline-flex items-center whitespace-nowrap">
                                        <input name="adRevenueSourceFilter" type="radio" value="adjust" class="h-4 w-4 text-indigo-600 border-gray-300 focus:ring-indigo-500">
                                        <span class="ml-2 text-sm text-gray-900">Adjust</span>
                                    </label>
                                </div>
                            </div>
                            <div>
                                <label for="adRevenueFilterInput" class="block text-xs font-medium text-gray-700 mb-1">Filter logs by text:</label>
                                <input type="text" id="adRevenueFilterInput" class="w-full p-2 border rounded-md shadow-sm" placeholder="Search in raw log...">
                            </div>
                        </div>
                    </div>
                </div>
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="overflow-x-auto">
                        <table class="min-w-full bg-white">
                           <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Status</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Event Name</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Details</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Action</th>
                                </tr>
                            </thead>
                            <tbody id="adRevenueTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 6: CallbackAd -->
            <div id="tabContentCallbackAd" class="hidden">
                 <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="mb-3 grid grid-cols-1 lg:grid-cols-[1fr_minmax(320px,420px)] gap-4 items-end">
                        <div>
                            <label class="block text-xs font-medium text-gray-700">Filter by Type:</label>
                            <div class="mt-2 flex flex-wrap lg:flex-nowrap items-center gap-x-4 gap-y-2 text-sm">
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeAll" name="callbackType" type="checkbox" value="all" checked class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">All</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeCallback" name="callbackType" type="checkbox" value="callback" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Callback Levelplay</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeAscendx" name="callbackType" type="checkbox" value="ascendx_callback" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Callback Ascendx</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeCloudx" name="callbackType" type="checkbox" value="cloudx_callback" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Callback Cloudx</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeGadsme" name="callbackType" type="checkbox" value="gadsme_callback" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Callback Gadsme</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeAdverty5" name="callbackType" type="checkbox" value="adverty5_callback" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Callback Adverty5</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="callbackTypeAdEvent" name="callbackType" type="checkbox" value="ad_event" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Ad Event</span>
                                </label>
                            </div>
                        </div>
                        <div class="lg:justify-self-end w-full lg:max-w-[420px]">
                            <label for="callbackAdFilterInput" class="block text-xs font-medium text-gray-700">Filter (in raw log):</label>
                            <input type="text" id="callbackAdFilterInput" class="mt-1 block w-full p-2 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500" placeholder="Search...">
                        </div>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="min-w-full bg-white">
                            <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Type</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Event / Key</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Details</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Action</th>
                                </tr>
                            </thead>
                            <tbody id="callbackAdTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 7: Bidding -->
            <div id="tabContentPriceRotation" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="mb-3 grid grid-cols-1 lg:grid-cols-[1fr_minmax(320px,520px)] gap-4 items-end">
                        <div>
                            <label class="block text-xs font-medium text-gray-700">Filter by Type:</label>
                            <div class="mt-2 flex flex-wrap items-center gap-x-4 gap-y-2 text-sm">
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypeAll" name="priceRotationTypeFilter" type="checkbox" value="all" checked class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">All</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypeLevelPlay" name="priceRotationTypeFilter" type="checkbox" value="levelplay" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">LevelPlay</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypeMax" name="priceRotationTypeFilter" type="checkbox" value="max" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">MAX</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypePriceCompare" name="priceRotationTypeFilter" type="checkbox" value="pricecompare" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">PriceCompare</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypeWaterfall" name="priceRotationTypeFilter" type="checkbox" value="waterfall" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Waterfall</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypeRewardedCap" name="priceRotationTypeFilter" type="checkbox" value="rewardedcap" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">RewardedCap</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationTypeInterstitialCap" name="priceRotationTypeFilter" type="checkbox" value="interstitialcap" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">InterstitialCap</span>
                                </label>
                            </div>
                            <label class="block text-xs font-medium text-gray-700 mt-3">Filter by Ad Type:</label>
                            <div class="mt-2 flex flex-wrap items-center gap-x-4 gap-y-2 text-sm">
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationAdTypeAll" name="priceRotationAdTypeFilter" type="checkbox" value="all" checked class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">All</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationAdTypeRewarded" name="priceRotationAdTypeFilter" type="checkbox" value="rewarded" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Rewarded</span>
                                </label>
                                <label class="inline-flex items-center whitespace-nowrap">
                                    <input id="priceRotationAdTypeInterstitial" name="priceRotationAdTypeFilter" type="checkbox" value="interstitial" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                                    <span class="ml-2 text-sm text-gray-900">Interstitial</span>
                                </label>
                            </div>
                        </div>
                        <div class="lg:justify-self-end w-full lg:max-w-[520px]">
                            <label for="priceRotationFilterInput" class="block text-xs font-medium text-gray-700">Filter (in raw log):</label>
                            <input type="text" id="priceRotationFilterInput" class="mt-1 block w-full p-2 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500" placeholder="Search...">
                        </div>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full min-w-[1100px] table-fixed bg-white">
                            <colgroup>
                                <col style="width: 13%;">
                                <col style="width: 14%;">
                                <col style="width: 29%;">
                                <col style="width: 36%;">
                                <col style="width: 8%;">
                            </colgroup>
                            <thead class="bg-gray-50 sticky top-0 z-10">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Device</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Type</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Details</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Raw Log</th>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Action</th>
                                </tr>
                            </thead>
                            <tbody id="priceRotationTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 8: SDK Check -->
            <div id="tabContentSdkCheck" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4 mb-4">
                    <h2 class="text-lg font-semibold mb-2">SDK Check Setup</h2>
                    <div class="grid grid-cols-1 md:grid-cols-[minmax(0,4fr)_minmax(280px,5fr)] gap-4 items-end">
                         <div>
                            <label for="sdkCheckInput" class="block text-xs font-medium text-gray-700 mb-1">SDKs to Check (manual input):</label>
                            <textarea id="sdkCheckInput" rows="8" class="w-full p-2 border rounded-md shadow-sm font-mono text-sm" placeholder="Chọn Saved SDK List hoặc nhập danh sách thủ công..."></textarea>
                         </div>
                         <div class="flex flex-col gap-3">
                            <div id="sdkCheckPresetPanel" class="rounded-lg border border-indigo-100 bg-indigo-50 p-3">
                                <div class="mb-2 flex items-center justify-between gap-2">
                                    <label class="block text-xs font-semibold text-indigo-900">Saved SDK List</label>
                                    <button id="reloadSdkCheckPresetsBtn" type="button" class="shrink-0 rounded border border-indigo-200 bg-white px-2 py-1 text-[11px] font-semibold text-indigo-700 hover:bg-indigo-100">Reload</button>
                                </div>
                                <div id="sdkCheckPresetList" class="flex flex-wrap gap-2"></div>
                                <div id="sdkCheckPresetStatus" class="mt-2 text-[11px] text-slate-600">Chọn preset để tự nạp danh sách.</div>
                            </div>
                            <button id="startSdkCheckBtn" class="bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold py-2 px-4 rounded-lg w-full h-10">Start Checking</button>
                         </div>
                    </div>
                </div>
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="flex items-center justify-between gap-3 mb-2">
                        <h2 class="text-lg font-semibold">SDK Check Results</h2>
                        <span id="sdkCheckPlatformBadge" class="text-xs font-semibold bg-slate-100 text-slate-700 px-3 py-1 rounded-full">Android Table</span>
                    </div>
                    <div id="sdkCheckAndroidPanel" class="sdk-check-panel overflow-x-auto" data-sdk-platform="android">
                        <div class="text-xs font-semibold text-slate-500 mb-2">Android SDK Check</div>
                        <table class="min-w-full bg-white">
                            <tbody id="sdkCheckTableBody" class="divide-y divide-gray-200"></tbody>
                        </table>
                    </div>
                    <div id="sdkCheckIosPanel" class="sdk-check-panel overflow-x-auto hidden" data-sdk-platform="ios">
                        <div class="text-xs font-semibold text-slate-500 mb-2">iOS SDK Check</div>
                        <table class="min-w-full bg-white">
                            <tbody id="sdkCheckIosTableBody" class="divide-y divide-gray-200">
                                <tr><td class="py-2 px-4 text-sm text-gray-500 italic">iOS table ready. Device connection will be added in the next step.</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 8A: Services Checker -->
            <div id="tabContentServicesChecker" class="hidden">
                <div class="bg-white rounded-xl shadow-md p-4">
                    <div class="flex items-center justify-between gap-3 mb-3">
                        <span id="servicesCheckerStatus" class="text-sm text-slate-500" role="status">Ready to start.</span>
                    </div>
                    <div id="servicesCheckerError" class="hidden mb-3 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700" role="alert"></div>
                    <iframe id="servicesCheckerFrame" title="Services Checker" class="hidden w-full rounded-lg border border-slate-200 bg-white" style="height: 78vh; min-height: 620px;"></iframe>
                </div>
            </div>
            
            <!-- TAB 8: Package -->
            <div id="tabContentPackage" class="hidden">
                 <div class="bg-white rounded-xl shadow-md p-5 mb-4">
                     <div class="grid grid-cols-1 lg:grid-cols-[0.85fr_1.7fr] gap-6 items-start">
                        <div class="max-w-md">
                            <label for="packageIdInput" id="packageIdInputLabel" class="block text-[11px] font-medium text-gray-700 mb-1">Package ID:</label>
                            <input type="text" id="packageIdInput" class="w-full p-2 text-[11px] border rounded-md shadow-sm" placeholder="com.example.app">
                            <div class="flex justify-center items-center gap-2 mt-3">
                                <button id="startPackageLogBtn" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs py-2 px-4 rounded-lg h-9">Start</button>
                                <button id="openPackageHistoryBtn" class="text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm bg-slate-600 hover:bg-slate-700 text-white h-9">Recorded Log</button>
                            </div>
                        </div>
                        <div class="max-w-[980px] min-h-[120px] flex flex-col justify-start">
                            <div class="flex items-center justify-between gap-3 mb-1">
                                <label class="block text-[11px] font-medium text-gray-700">Quick Select:</label>
                                <button id="reloadPackagePresetsBtn" type="button" class="text-[11px] font-semibold text-indigo-700 border border-indigo-200 bg-white hover:bg-indigo-50 rounded px-2 py-1">Reload list</button>
                            </div>
                            <div id="packageQuickSelectList" class="grid grid-cols-1 xl:grid-cols-2 gap-x-16 gap-y-1 text-[11px] text-gray-700 pt-1">
                                <span id="packageQuickSelectStatus" class="text-slate-500 italic">Loading projects...</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        </div>
        <!-- FOOTER -->
        <div id="appFooter" class="flex-shrink-0 flex items-center justify-center pt-2">
            <img src="{{ indiez_logo }}" alt="INDIEZ" class="h-10 w-auto select-none" draggable="false">
        </div>
    </div>
    
    <!-- MODALS -->
    <div id="jsonModal" class="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 hidden">
        <div class="bg-white rounded-xl shadow-lg p-6 w-1/2 h-2/3 flex flex-col">
            <div class="flex justify-between items-center border-b pb-3 mb-4">
                <h2 class="text-xl font-bold text-gray-800">Formatted JSON</h2>
                <button id="closeJsonModal" class="text-gray-500 hover:text-gray-800 text-2xl font-bold">&times;</button>
            </div>
            <div class="flex-grow overflow-auto bg-gray-800 text-white p-4 rounded-md">
                <pre><code id="jsonContent"></code></pre>
            </div>
        </div>
    </div>

    <div id="packageHistoryModal" class="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 hidden">
        <div class="bg-white shadow-lg p-4 w-screen h-screen flex flex-col">
            <div class="flex justify-between items-start border-b pb-3 mb-4">
                <div class="flex items-center gap-3 flex-wrap">
                    <h2 class="text-xl font-bold text-gray-800">Recorded Package Logs</h2>
                    <button id="loadPackageHistoryBtn" class="bg-slate-600 hover:bg-slate-700 text-white font-semibold text-xs py-2 px-4 rounded-lg h-9">Load</button>
                    <button id="loadMorePackageHistoryBtn" class="bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-xs py-2 px-4 rounded-lg h-9 hidden">Load More</button>
                    <button id="refreshPackageSessionsBtn" class="bg-gray-200 hover:bg-gray-300 text-gray-800 font-semibold text-xs py-2 px-4 rounded-lg h-9">Refresh Sessions</button>
                    <button id="exportPackageHistoryAllBtn" class="bg-emerald-500 hover:bg-emerald-600 text-white font-semibold text-xs py-2 px-4 rounded-lg h-9">Export Full</button>
                    <button id="exportPackageHistoryFilteredBtn" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-xs py-2 px-4 rounded-lg h-9">Export Filtered</button>
                </div>
                <button id="closePackageHistoryModal" class="text-gray-500 hover:text-gray-800 text-2xl font-bold">&times;</button>
            </div>
            <div class="flex flex-col gap-4">
                <div class="grid grid-cols-1 xl:grid-cols-[260px_110px_110px_150px_1fr_1fr_1fr_1fr] gap-3 items-end">
                    <div>
                        <div class="flex items-center justify-between gap-2 mb-1">
                            <label for="packageHistorySessionSelect" class="block text-[11px] font-medium text-gray-700">Saved Sessions:</label>
                            <div id="packageHistoryPlatformSwitch" class="inline-flex rounded-md border border-slate-300 overflow-hidden text-[11px] font-semibold" title="Android and iOS sessions are listed separately">
                                <button type="button" data-platform="android" class="px-2 py-0.5">Android</button>
                                <button type="button" data-platform="ios" class="px-2 py-0.5 border-l border-slate-300">iOS</button>
                            </div>
                        </div>
                        <select id="packageHistorySessionSelect" class="w-full p-2 text-[11px] border rounded-md shadow-sm"></select>
                    </div>
                    <div>
                        <label class="block text-[11px] font-medium text-gray-700 mb-1">&nbsp;</label>
                        <button id="clearPackageHistoryBtn" class="w-full bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 font-semibold text-xs py-2 px-4 rounded-lg h-9">Clear Data</button>
                    </div>
                    <div>
                        <label class="block text-[11px] font-medium text-gray-700 mb-1">&nbsp;</label>
                        <button id="clearPackageHistoryFiltersBtn" class="w-full bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-200 font-semibold text-xs py-2 px-4 rounded-lg h-9">Clear Filter</button>
                    </div>
                    <div>
                        <label for="packageHistoryQuickFilter" class="block text-[11px] font-medium text-gray-700 mb-1">Quick Filter:</label>
                        <select id="packageHistoryQuickFilter" class="w-full p-2 text-[11px] border rounded-md shadow-sm h-9"><option value="">All</option></select>
                    </div>
                    <div>
                        <label for="packageHistoryFilterInput3" class="block text-[11px] font-medium text-gray-700 mb-1">Tag Filter:</label>
                        <input type="text" id="packageHistoryFilterInput3" class="w-full p-2 text-[11px] border rounded-md shadow-sm" placeholder="Search saved log tag...">
                    </div>
                    <div>
                        <label for="packageHistoryFilterInput" class="block text-[11px] font-medium text-gray-700 mb-1">Filter 1:</label>
                        <input type="text" id="packageHistoryFilterInput" class="w-full p-2 text-[11px] border rounded-md shadow-sm" placeholder="Search saved log text 1...">
                    </div>
                    <div>
                        <label for="packageHistoryFilterInput2" class="block text-[11px] font-medium text-gray-700 mb-1">Filter 2:</label>
                        <input type="text" id="packageHistoryFilterInput2" class="w-full p-2 text-[11px] border rounded-md shadow-sm" placeholder="Search saved log text 2...">
                    </div>
                    <div>
                        <label for="packageHistoryFilterInput4" class="block text-[11px] font-medium text-gray-700 mb-1">Filter 3:</label>
                        <input type="text" id="packageHistoryFilterInput4" class="w-full p-2 text-[11px] border rounded-md shadow-sm" placeholder="Search saved log text 3...">
                    </div>
                </div>
                <p id="packageHistoryMeta" class="text-[11px] text-gray-500">No session selected.</p>
                <div id="packageHistoryMaintenance" class="flex flex-wrap items-center gap-3 text-[11px] text-gray-600 border rounded-md px-3 py-2 bg-slate-50">
                    <span id="packageHistoryStorageInfo" class="flex-1 min-w-[280px]">Đang đọc dung lượng lịch sử…</span>
                    <button id="applyPackageRetentionBtn" type="button" class="bg-amber-500 hover:bg-amber-600 text-white font-semibold text-xs py-1.5 px-3 rounded-lg">Dọn lịch sử</button>
                    <button id="compactPackageHistoryBtn" type="button" class="bg-slate-600 hover:bg-slate-700 disabled:opacity-50 disabled:cursor-not-allowed text-white font-semibold text-xs py-1.5 px-3 rounded-lg">Compact</button>
                </div>
            </div>
            <div id="packageHistoryTableWrap" class="overflow-auto border rounded-md mt-3 flex-1">
                <table class="min-w-full bg-white">
                    <thead class="bg-gray-50 sticky top-0 z-10">
                        <tr>
                            <th class="text-left text-xs font-semibold text-gray-600 py-2 px-2 border-b">Time</th>
                            <th class="text-left text-xs font-semibold text-gray-600 py-2 px-2 border-b">Device</th>
                            <th class="text-left text-xs font-semibold text-gray-600 py-2 px-2 border-b w-[88px] max-w-[88px]">Tag</th>
                            <th class="text-left text-xs font-semibold text-gray-600 py-2 px-2 border-b">Message</th>
                        </tr>
                    </thead>
                    <tbody id="packageHistoryTableBody" class="divide-y divide-gray-200"></tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="packageStreamModal" class="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 hidden">
        <div class="bg-white shadow-lg p-4 w-screen h-screen flex flex-col">
            <div class="border-b pb-2 mb-2">
                <div class="flex flex-wrap items-center gap-x-10 gap-y-1 mb-2">
                    <h2 class="text-xl font-bold text-gray-800">Package Log Stream</h2>
                    <div class="flex items-center gap-2">
                        <label for="packageStreamProjectSelect" class="text-xs font-semibold text-gray-700">Project:</label>
                        <select id="packageStreamProjectSelect" class="py-1 px-2 text-xs border rounded-md shadow-sm min-w-[240px] max-w-[360px]" title="Switch project: saves the current session and starts capturing the selected project"></select>
                    </div>
                    <div class="flex flex-wrap items-center gap-4">
                        <label class="inline-flex items-center gap-2">
                            <input id="showErrorsOnly" type="checkbox" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                            <span class="block text-xs text-gray-900">Show errors only</span>
                        </label>
                        <label class="inline-flex items-center gap-2">
                            <input id="showWarningsOnly" type="checkbox" class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                            <span class="block text-xs text-gray-900">Show only Warning</span>
                        </label>
                        <label class="inline-flex items-center gap-2">
                            <input id="autoScroll" type="checkbox" checked class="h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500">
                            <span class="block text-xs text-gray-900">Auto-scroll</span>
                        </label>
                    </div>
                </div>
                <div class="flex flex-wrap items-end gap-x-3 gap-y-2">
                    <div class="flex items-center gap-3 shrink-0 mr-3">
                        <button id="pausePackageLogBtn" class="text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm bg-amber-500 hover:bg-amber-600 text-white">Pause</button>
                        <button id="stopPackageLogBtn" class="text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm bg-red-500 hover:bg-red-600 text-white">Stop</button>
                        <button id="searchPackageSessionBtn" type="button" class="text-sm font-semibold py-2 px-3 rounded-lg transition-colors shadow-sm bg-slate-600 hover:bg-slate-700 text-white" title="Search every row of this session in saved history (Filter 1-3, Tag and quick filters)">Tìm toàn session</button>
                    </div>
                    <div class="w-44">
                        <label for="packageTagFilterInput" class="block text-[11px] font-medium text-gray-700 mb-1">Tag Filter:</label>
                        <input type="text" id="packageTagFilterInput" class="w-full py-1.5 px-2 text-sm border rounded-md shadow-sm" placeholder="Tag...">
                    </div>
                    <div class="w-56">
                        <label for="packageFilterInput" class="block text-[11px] font-medium text-gray-700 mb-1">Filter 1:</label>
                        <input type="text" id="packageFilterInput" class="w-full py-1.5 px-2 text-sm border rounded-md shadow-sm" placeholder="Search text 1...">
                    </div>
                    <div class="w-56">
                        <label for="packageFilterInput2" class="block text-[11px] font-medium text-gray-700 mb-1">Filter 2:</label>
                        <input type="text" id="packageFilterInput2" class="w-full py-1.5 px-2 text-sm border rounded-md shadow-sm" placeholder="Search text 2...">
                    </div>
                    <div class="w-56">
                        <label for="packageFilterInput3" class="block text-[11px] font-medium text-gray-700 mb-1">Filter 3:</label>
                        <input type="text" id="packageFilterInput3" class="w-full py-1.5 px-2 text-sm border rounded-md shadow-sm" placeholder="Search text 3...">
                    </div>
                    <button id="clearPackageFiltersBtn" type="button" class="shrink-0 text-xs font-semibold py-2 px-3 rounded-md border border-gray-200 bg-gray-100 hover:bg-gray-200 text-gray-700" title="Clear Tag, Filter 1-3 and the quick filter">Clear filter</button>
                </div>
            </div>
            <div id="packageDbWarning" class="hidden mb-2 rounded-md border border-red-300 bg-red-50 px-3 py-2 text-xs font-semibold text-red-700"></div>
            <div class="flex-1 flex gap-3" style="min-height: 0;">
                <div id="packageQuickFilterPanel" class="w-40 shrink-0 overflow-y-auto border-r pr-2 pt-1 flex flex-col gap-y-2 text-xs text-gray-700">
                    <label class="inline-flex items-center gap-2">
                        <input type="radio" name="tagQuickFilter" value="" checked class="h-4 w-4 text-indigo-600 border-gray-300 focus:ring-indigo-500">
                        <span>All</span>
                    </label>
                    <div id="packageQuickFilterList" class="flex flex-col gap-y-2"></div>
                    <button id="editPackageQuickFiltersBtn" type="button" class="self-start mt-1 text-[11px] font-semibold text-indigo-600 hover:text-indigo-800 hover:underline">Edit filters…</button>
                </div>
            <div class="relative flex-1 flex flex-col min-w-0" style="min-height: 0;">
            <div id="packageOlderStatus" class="hidden absolute left-4 top-9 z-20 rounded bg-slate-700 px-2 py-1 text-[11px] text-white shadow"></div>
            <button id="packageNewRowsBtn" type="button" class="hidden bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold py-1.5 px-3 rounded-full shadow">New rows</button>
            <div id="packageLogContainer" class="overflow-auto flex-1 border rounded-md" style="min-height: 0;">
                <table id="packageLogTableLive" class="min-w-full bg-white">
                   <thead class="bg-gray-50 sticky top-0 z-10">
                        <tr>
                            <th class="text-left text-xs font-semibold text-gray-600 py-1.5 px-2 border-b time-header resizable col-time">Time<div class="resizer" data-col="time"></div></th>
                            <th class="text-left text-xs font-semibold text-gray-600 py-1.5 pr-1 pl-2 border-b tag-header resizable col-tag">Tag<div class="resizer" data-col="tag"></div></th>
                            <th class="text-left text-xs font-semibold text-gray-600 py-1.5 pl-1 pr-3 border-b resizable col-message">Message<div class="resizer" data-col="message"></div></th>
                        </tr>
                    </thead>
                    <tbody id="packageLogTableBody"></tbody>
                </table>
            </div>
            </div>
            </div>
        </div>
    </div>

    <div id="packageQuickFilterEditor" class="fixed inset-0 flex items-center justify-center bg-black bg-opacity-50 hidden" style="z-index: 70;">
        <div class="bg-white rounded-xl shadow-lg p-5 w-[min(1000px,96vw)] max-h-[90vh] flex flex-col">
            <div class="flex justify-between items-start gap-4 border-b pb-3 mb-3">
                <div>
                    <h2 class="text-lg font-bold text-gray-800">Quick Filters – <span id="packageQuickFilterEditorPlatform">Android</span></h2>
                    <p class="text-[11px] text-gray-500 mt-1">Mỗi ô giống ô lọc tương ứng phía trên: Tag tìm trong tag, Filter 1 / 2 / 3 tìm trên toàn bộ dòng log. Dòng log phải khớp tất cả ô đã điền, ô trống thì bỏ qua. Android và iOS có danh sách riêng, lưu trên máy này, không ảnh hưởng preset.</p>
                </div>
                <button id="closePackageQuickFilterEditor" type="button" class="text-gray-500 hover:text-gray-800 text-2xl font-bold">&times;</button>
            </div>
            <div class="overflow-auto flex-1 border rounded-md" style="min-height: 0;">
                <table class="min-w-full bg-white text-xs">
                    <thead class="bg-gray-50 sticky top-0 z-10">
                        <tr>
                            <th class="text-left font-semibold text-gray-600 py-2 px-2 border-b">Name</th>
                            <th class="qf-tag-col text-left font-semibold text-gray-600 py-2 px-2 border-b">Tag</th>
                            <th class="text-left font-semibold text-gray-600 py-2 px-2 border-b">Filter 1</th>
                            <th class="text-left font-semibold text-gray-600 py-2 px-2 border-b">Filter 2</th>
                            <th class="text-left font-semibold text-gray-600 py-2 px-2 border-b">Filter 3</th>
                            <th class="py-2 px-2 border-b"></th>
                        </tr>
                    </thead>
                    <tbody id="packageQuickFilterEditorBody"></tbody>
                </table>
            </div>
            <p id="packageQuickFilterEditorError" class="hidden mt-2 text-xs font-semibold text-red-600"></p>
            <div class="flex flex-wrap justify-between gap-2 mt-3">
                <div class="flex gap-2">
                    <button id="addPackageQuickFilterBtn" type="button" class="text-xs font-semibold py-2 px-3 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white">+ Add filter</button>
                    <button id="resetPackageQuickFiltersBtn" type="button" class="text-xs font-semibold py-2 px-3 rounded-lg bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-200">Reset to defaults</button>
                    <button id="exportPackageQuickFiltersBtn" type="button" class="text-xs font-semibold py-2 px-3 rounded-lg bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-200" title="Save this list to a .json file (backup or share)">Export</button>
                    <button id="importPackageQuickFiltersBtn" type="button" class="text-xs font-semibold py-2 px-3 rounded-lg bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-200" title="Load a list from an exported .json file (replaces this list after Save)">Import</button>
                    <input type="file" id="packageQuickFilterImportInput" accept=".json,application/json" class="hidden">
                </div>
                <div class="flex gap-2">
                    <button id="cancelPackageQuickFilterEditor" type="button" class="text-xs font-semibold py-2 px-3 rounded-lg bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-200">Cancel</button>
                    <button id="savePackageQuickFiltersBtn" type="button" class="text-xs font-semibold py-2 px-4 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white">Save</button>
                </div>
            </div>
        </div>
    </div>

    <div id="logDetailModal" class="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 hidden">
        <div class="bg-white rounded-xl shadow-lg p-6 w-3/4 h-2/3 flex flex-col">
            <div class="flex justify-between items-center border-b pb-3 mb-4">
                <h2 class="text-xl font-bold text-gray-800">Log Details</h2>
                <div class="flex items-center gap-3">
                    <button id="convertLogJsonBtn" class="text-sm bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 px-3 rounded">Convert to JSON</button>
                    <button id="closeLogDetailModal" class="text-gray-500 hover:text-gray-800 text-2xl font-bold">&times;</button>
                </div>
            </div>
            <div class="flex-grow overflow-auto bg-gray-50 p-4 rounded-md border">
                <pre id="logDetailContent" class="text-sm font-mono whitespace-pre-wrap text-gray-800 select-text"></pre>
            </div>
        </div>
    </div>

    <!-- SCRIPTS -->
    <script>
        const socket = typeof window.io === 'function'
            ? window.io()
            : { on() { return this; }, emit() { return this; } };
        let currentTab = 'LoadAdsExt';
        // --- UI zoom (text size) ---
        // Scales the root font size, so every rem-based size (Tailwind text,
        // padding, gaps) follows; the level is saved on the server because the
        // desktop webview does not keep localStorage between launches.
        var UI_ZOOM_LEVELS = [0.7, 0.75, 0.8, 0.85, 0.9, 1, 1.1, 1.25, 1.4];
        var UI_DEFAULT_ZOOM = {{ ui_default_zoom }};
        var currentUiZoom = Number(document.documentElement.dataset.uiZoom) || UI_DEFAULT_ZOOM;
        var uiZoomSaveTimer = null;
        function renderUiZoomLabel() {
            const label = document.getElementById('uiZoomReset');
            if (label) label.textContent = `${Math.round(currentUiZoom * 100)}%`;
        }
        function applyUiZoom(level) {
            currentUiZoom = level;
            document.documentElement.style.fontSize = `${16 * level}px`;
            document.documentElement.dataset.uiZoom = String(level);
            renderUiZoomLabel();
            if (typeof schedulePackageViewportRender === 'function') schedulePackageViewportRender();
            clearTimeout(uiZoomSaveTimer);
            uiZoomSaveTimer = setTimeout(() => {
                fetch('/api/ui-settings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ zoom: level }),
                }).catch(() => {});
            }, 300);
        }
        function stepUiZoom(direction) {
            let index = UI_ZOOM_LEVELS.indexOf(currentUiZoom);
            if (index < 0) index = UI_ZOOM_LEVELS.indexOf(UI_DEFAULT_ZOOM);
            index = Math.max(0, Math.min(UI_ZOOM_LEVELS.length - 1, index + direction));
            applyUiZoom(UI_ZOOM_LEVELS[index]);
        }
        document.getElementById('uiZoomOut')?.addEventListener('click', () => stepUiZoom(-1));
        document.getElementById('uiZoomIn')?.addEventListener('click', () => stepUiZoom(1));
        document.getElementById('uiZoomReset')?.addEventListener('click', () => applyUiZoom(UI_DEFAULT_ZOOM));
        document.addEventListener('keydown', (e) => {
            if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
            if (e.key === '-' || e.key === '_') { e.preventDefault(); stepUiZoom(-1); }
            else if (e.key === '=' || e.key === '+') { e.preventDefault(); stepUiZoom(1); }
            else if (e.key === '0') { e.preventDefault(); applyUiZoom(UI_DEFAULT_ZOOM); }
        });
        renderUiZoomLabel();

        // Hide Top / Show Top: collapses the header card to give the tables more
        // room; saved on the server like the zoom level.
        function applyTopHidden(hidden) {
            document.getElementById('appHeader')?.classList.toggle('hidden', hidden);
            const btn = document.getElementById('toggleTopBtn');
            if (btn) btn.textContent = hidden ? 'Show Top' : 'Hide Top';
            if (typeof schedulePackageViewportRender === 'function') schedulePackageViewportRender();
            fetch('/api/ui-settings', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ hide_top: hidden }),
            }).catch(() => {});
        }
        document.getElementById('toggleTopBtn')?.addEventListener('click', () => {
            applyTopHidden(!document.getElementById('appHeader')?.classList.contains('hidden'));
        });

        // --- Package Log quick filters ---
        // User-editable lists saved per platform on this machine
        // (package_quick_filters_<platform>.json, not preset data).  A filter
        // works like the boxes above the table: Tag (Android only) must be in
        // the tag; Message 1 and Message 2 must each be in the whole log line.
        var packageQuickFilters = {{ package_quick_filters|tojson }};
        var packageQuickFilterDefaults = {{ package_quick_filter_defaults|tojson }};
        var packageQuickFilterDraft = [];
        var packageQuickFilterEditorPlatform = 'android';
        function qfEsc(value) {
            return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
        }
        function quickFilterPlatform() {
            return activePlatform === 'ios' ? 'ios' : 'android';
        }
        function availableQuickFilters() {
            return (packageQuickFilters && packageQuickFilters[quickFilterPlatform()]) || [];
        }
        function resolveQuickFilter(id) {
            if (!id) return null;
            const item = availableQuickFilters().find(f => f.id === id);
            if (!item) return null;
            // Matching is never case sensitive.
            const norm = text => String(text || '').trim().toLowerCase();
            const tag = quickFilterPlatform() === 'ios' ? '' : norm(item.tag);
            const terms = [norm(item.message1), norm(item.message2), norm(item.message3)].filter(Boolean);
            if (!tag && !terms.length) return null;
            return { tag, terms };
        }
        function quickFilterTitle(item) {
            const parts = [];
            if (item.tag && quickFilterPlatform() !== 'ios') parts.push(`Tag: ${item.tag}`);
            if (item.message1) parts.push(`Filter 1: ${item.message1}`);
            if (item.message2) parts.push(`Filter 2: ${item.message2}`);
            if (item.message3) parts.push(`Filter 3: ${item.message3}`);
            return parts.join(' · ');
        }
        function renderPackageQuickFilters() {
            const list = document.getElementById('packageQuickFilterList');
            if (!list) return;
            const checked = document.querySelector('input[name="tagQuickFilter"]:checked')?.value || '';
            const items = availableQuickFilters();
            list.innerHTML = items.map(f => `<label class="inline-flex items-center gap-2 min-w-0" title="${qfEsc(quickFilterTitle(f))}"><input type="radio" name="tagQuickFilter" value="${qfEsc(f.id)}" class="h-4 w-4 shrink-0 text-indigo-600 border-gray-300 focus:ring-indigo-500"><span class="truncate">${qfEsc(f.name)}</span></label>`).join('');
            const keep = items.some(f => f.id === checked) ? checked : '';
            const radio = Array.from(document.querySelectorAll('input[name="tagQuickFilter"]')).find(r => r.value === keep);
            if (radio) radio.checked = true;
            renderHistoryQuickFilterOptions();
        }
        // Recorded Logs show one platform at a time; its quick filters come
        // from that platform's list.
        var packageHistoryPlatform = 'android';
        function historyQuickFilters() {
            return (packageQuickFilters && packageQuickFilters[packageHistoryPlatform]) || [];
        }
        function renderHistoryQuickFilterOptions() {
            const select = document.getElementById('packageHistoryQuickFilter');
            if (!select) return;
            const current = select.value;
            const items = historyQuickFilters();
            select.innerHTML = '<option value="">All</option>' + items.map(f => `<option value="${qfEsc(f.id)}">${qfEsc(f.name)}</option>`).join('');
            select.value = items.some(f => f.id === current) ? current : '';
        }

        function renderQuickFilterEditor() {
            const body = document.getElementById('packageQuickFilterEditorBody');
            if (!body) return;
            const isIos = packageQuickFilterEditorPlatform === 'ios';
            document.querySelectorAll('.qf-tag-col').forEach(el => el.classList.toggle('hidden', isIos));
            const title = document.getElementById('packageQuickFilterEditorPlatform');
            if (title) title.textContent = isIos ? 'iOS' : 'Android';
            if (!packageQuickFilterDraft.length) {
                body.innerHTML = '<tr><td colspan="7" class="py-6 text-center text-gray-400">No filters. Add one or reset to defaults.</td></tr>';
                return;
            }
            const input = 'w-full p-1.5 border rounded-md';
            const last = packageQuickFilterDraft.length - 1;
            body.innerHTML = packageQuickFilterDraft.map((f, i) => `<tr class="border-b align-middle">
                <td class="p-1.5 w-44"><input data-i="${i}" data-field="name" value="${qfEsc(f.name)}" maxlength="40" class="${input}"></td>
                <td class="p-1.5${isIos ? ' hidden' : ''}"><input data-i="${i}" data-field="tag" value="${qfEsc(f.tag)}" maxlength="200" placeholder="Tag..." class="${input} font-mono"></td>
                <td class="p-1.5"><input data-i="${i}" data-field="message1" value="${qfEsc(f.message1)}" maxlength="200" placeholder="Search text 1..." class="${input} font-mono"></td>
                <td class="p-1.5"><input data-i="${i}" data-field="message2" value="${qfEsc(f.message2)}" maxlength="200" placeholder="Search text 2..." class="${input} font-mono"></td>
                <td class="p-1.5"><input data-i="${i}" data-field="message3" value="${qfEsc(f.message3)}" maxlength="200" placeholder="Search text 3..." class="${input} font-mono"></td>
                <td class="p-1.5 whitespace-nowrap text-right">
                    <button type="button" data-i="${i}" data-action="up" class="px-2 py-1 rounded hover:bg-gray-100 disabled:opacity-30" title="Move up"${i === 0 ? ' disabled' : ''}>↑</button>
                    <button type="button" data-i="${i}" data-action="down" class="px-2 py-1 rounded hover:bg-gray-100 disabled:opacity-30" title="Move down"${i === last ? ' disabled' : ''}>↓</button>
                    <button type="button" data-i="${i}" data-action="delete" class="px-2 py-1 rounded text-red-600 hover:bg-red-50" title="Delete">✕</button>
                </td></tr>`).join('');
        }
        function showQuickFilterEditorError(message, isInfo = false) {
            const el = document.getElementById('packageQuickFilterEditorError');
            if (!el) return;
            el.textContent = message || '';
            el.classList.toggle('hidden', !message);
            el.classList.toggle('text-red-600', !isInfo);
            el.classList.toggle('text-emerald-700', isInfo);
        }
        function openQuickFilterEditor() {
            packageQuickFilterEditorPlatform = quickFilterPlatform();
            packageQuickFilterDraft = availableQuickFilters().map(f => ({ ...f }));
            showQuickFilterEditorError('');
            renderQuickFilterEditor();
            document.getElementById('packageQuickFilterEditor')?.classList.remove('hidden');
        }
        function closeQuickFilterEditor() {
            document.getElementById('packageQuickFilterEditor')?.classList.add('hidden');
        }
        function saveQuickFilterEditor() {
            const platform = packageQuickFilterEditorPlatform;
            const filters = packageQuickFilterDraft.map(f => ({ ...f, tag: platform === 'ios' ? '' : String(f.tag || '').trim() }));
            for (let i = 0; i < filters.length; i++) {
                const f = filters[i];
                if (!String(f.name || '').trim()) return showQuickFilterEditorError(`Row ${i + 1}: name is empty.`);
                if (![f.tag, f.message1, f.message2, f.message3].some(v => String(v || '').trim())) {
                    return showQuickFilterEditorError(`Row ${i + 1} (${f.name}): fill at least one filter box.`);
                }
            }
            showQuickFilterEditorError('');
            fetch('/api/package-quick-filters', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ platform, filters }),
            })
                .then(r => r.json())
                .then(data => {
                    if (!data.ok) throw new Error(data.error || 'save_failed');
                    packageQuickFilters[platform] = data.filters || [];
                    renderPackageQuickFilters();
                    if (typeof renderPackageLogTable === 'function') renderPackageLogTable(true);
                    closeQuickFilterEditor();
                })
                .catch(err => showQuickFilterEditorError(`Could not save: ${err.message || err}`));
        }
        document.getElementById('editPackageQuickFiltersBtn')?.addEventListener('click', openQuickFilterEditor);
        document.getElementById('closePackageQuickFilterEditor')?.addEventListener('click', closeQuickFilterEditor);
        document.getElementById('cancelPackageQuickFilterEditor')?.addEventListener('click', closeQuickFilterEditor);
        document.getElementById('savePackageQuickFiltersBtn')?.addEventListener('click', saveQuickFilterEditor);
        document.getElementById('addPackageQuickFilterBtn')?.addEventListener('click', () => {
            const id = 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
            packageQuickFilterDraft.push({ id, name: 'New filter', tag: '', message1: '', message2: '', message3: '' });
            renderQuickFilterEditor();
            const inputs = document.querySelectorAll('#packageQuickFilterEditorBody input[data-field="name"]');
            inputs[inputs.length - 1]?.select();
        });
        document.getElementById('resetPackageQuickFiltersBtn')?.addEventListener('click', () => {
            packageQuickFilterDraft = ((packageQuickFilterDefaults || {})[packageQuickFilterEditorPlatform] || []).map(f => ({ ...f }));
            showQuickFilterEditorError('');
            renderQuickFilterEditor();
        });
        // In the desktop app the server opens the native Save dialog; in a
        // browser the page asks for a location itself (save picker, or a
        // normal download where the picker is not available).
        var packageNativeFileDialog = {{ 'true' if native_file_dialog else 'false' }};
        async function exportQuickFilters() {
            const platform = packageQuickFilterEditorPlatform;
            const filters = packageQuickFilterDraft.map(f => ({ ...f, tag: platform === 'ios' ? '' : f.tag }));
            const post = (extra) => fetch('/api/package-quick-filters/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ platform, filters, ...extra }),
            }).then(r => r.json()).then(data => {
                if (!data.ok && !data.cancelled) throw new Error(data.error || 'export_failed');
                return data;
            });
            try {
                if (packageNativeFileDialog) {
                    const data = await post({});
                    if (!data.cancelled) showQuickFilterEditorError(`Exported ${data.count} filters to ${data.path}`, true);
                    return;
                }
                const stamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '_');
                const suggestedName = `quick_filters_${platform}_${stamp}.json`;
                let handle = null;
                if (typeof window.showSaveFilePicker === 'function') {
                    // Ask for the location first, while the click still counts as a user action.
                    try {
                        handle = await window.showSaveFilePicker({
                            suggestedName,
                            types: [{ description: 'JSON file', accept: { 'application/json': ['.json'] } }],
                        });
                    } catch (err) {
                        if (err && err.name === 'AbortError') return;
                        handle = null;
                    }
                }
                const data = await post({ download: true });
                if (handle) {
                    const writable = await handle.createWritable();
                    await writable.write(data.content);
                    await writable.close();
                    showQuickFilterEditorError(`Exported ${data.count} filters to ${handle.name}`, true);
                    return;
                }
                const url = URL.createObjectURL(new Blob([data.content], { type: 'application/json' }));
                const link = document.createElement('a');
                link.href = url;
                link.download = suggestedName;
                document.body.appendChild(link);
                link.click();
                link.remove();
                setTimeout(() => URL.revokeObjectURL(url), 1000);
                showQuickFilterEditorError(`Exported ${data.count} filters as ${suggestedName} (browser download).`, true);
            } catch (err) {
                showQuickFilterEditorError(`Could not export: ${err.message || err}`);
            }
        }
        function importQuickFilters(file) {
            if (!file) return;
            const platform = packageQuickFilterEditorPlatform;
            const reader = new FileReader();
            reader.onload = () => {
                let data;
                try {
                    data = JSON.parse(String(reader.result || ''));
                } catch (err) {
                    showQuickFilterEditorError('Could not import: the file is not valid JSON.');
                    return;
                }
                fetch('/api/package-quick-filters/import', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ platform, data }),
                })
                    .then(r => r.json())
                    .then(result => {
                        if (!result.ok) {
                            if (result.error === 'wrong_platform') {
                                throw new Error(`this file is for ${result.file_platform === 'ios' ? 'iOS' : 'Android'}; open the filters of that platform to import it`);
                            }
                            throw new Error(result.error || 'import_failed');
                        }
                        packageQuickFilterDraft = (result.filters || []).map(f => ({ ...f }));
                        renderQuickFilterEditor();
                        showQuickFilterEditorError(`Imported ${packageQuickFilterDraft.length} filters from ${file.name}. Press Save to replace the current list.`, true);
                    })
                    .catch(err => showQuickFilterEditorError(`Could not import: ${err.message || err}`));
            };
            reader.onerror = () => showQuickFilterEditorError('Could not read the file.');
            reader.readAsText(file);
        }
        document.getElementById('exportPackageQuickFiltersBtn')?.addEventListener('click', exportQuickFilters);
        document.getElementById('importPackageQuickFiltersBtn')?.addEventListener('click', () => {
            const input = document.getElementById('packageQuickFilterImportInput');
            if (!input) return;
            input.value = '';
            input.click();
        });
        document.getElementById('packageQuickFilterImportInput')?.addEventListener('change', (e) => {
            importQuickFilters(e.target.files && e.target.files[0]);
        });
        (function bindQuickFilterEditorBody() {
            const body = document.getElementById('packageQuickFilterEditorBody');
            if (!body) return;
            const update = (e) => {
                const el = e.target;
                const i = Number(el.dataset.i);
                const field = el.dataset.field;
                if (!field || !packageQuickFilterDraft[i]) return;
                packageQuickFilterDraft[i][field] = el.value;
            };
            body.addEventListener('input', update);
            body.addEventListener('change', update);
            body.addEventListener('click', (e) => {
                const btn = e.target.closest('button[data-action]');
                if (!btn) return;
                const i = Number(btn.dataset.i);
                const list = packageQuickFilterDraft;
                if (btn.dataset.action === 'delete') list.splice(i, 1);
                else if (btn.dataset.action === 'up' && i > 0) [list[i - 1], list[i]] = [list[i], list[i - 1]];
                else if (btn.dataset.action === 'down' && i < list.length - 1) [list[i + 1], list[i]] = [list[i], list[i + 1]];
                renderQuickFilterEditor();
            });
        })();

        let selectedDevice = 'all';

        // --- Common Elements ---
        const deviceListEl = document.getElementById('deviceList');
        const pauseBtn = document.getElementById('pauseBtn');
        const jsonModal = document.getElementById('jsonModal');
        const closeJsonModal = document.getElementById('closeJsonModal');
        const jsonContent = document.getElementById('jsonContent');
        const logDetailModal = document.getElementById('logDetailModal');
        const closeLogDetailModal = document.getElementById('closeLogDetailModal');
        const logDetailContent = document.getElementById('logDetailContent');
        const convertLogJsonBtn = document.getElementById('convertLogJsonBtn');
        const deviceFilter = document.getElementById('deviceFilter');
        const clearAllBtn = document.getElementById('clearAllBtn');
        const restartAppBtn = document.getElementById('restartAppBtn');
        const platformBtn = document.getElementById('platformBtn');
        const platformModal = document.getElementById('platformModal');
        const installationIdGameEl = document.getElementById('installationIdGame');
        const installationIdValueEl = document.getElementById('installationIdValue');
        const copyInstallationIdBtn = document.getElementById('copyInstallationIdBtn');
        let activePlatform = localStorage.getItem('eventInspectorPlatform') || '';
        renderPackageQuickFilters();
        let sdkCheckPresets = {{ sdk_check_presets | tojson }};
        let brightSkadLastPayload = null;
        // Keep SDK Check state declarations above the platform/bootstrap code.
        // setActivePlatform() runs during initial page setup, before the rest
        // of the SDK Check handlers are registered.  A later `let` declaration
        // leaves the binding in the temporal dead zone and breaks platform
        // switching with `Cannot access 'sdkCheckRunning' before
        // initialization`.
        let sdkCheckRunning = false;
        let sdkCheckClearPending = false;
        socket.on('connect', () => {
            if (!activePlatform) return;
            socket.emit('set_platform', {
                platform: activePlatform === 'ios' ? 'ios' : 'android',
                reset: false,
            });
        });
        let installationIdState = {
            device_id: '',
            device_name: '',
            package_id: '',
            game_code: '',
            installation_id: '',
            platform: 'android',
        };

        function platformLabel(platform) {
            return platform === 'ios' ? 'iOS' : 'Android';
        }

        async function copyTextToClipboard(text) {
            const value = String(text || '');
            if (!value) return false;

            if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
                try {
                    await navigator.clipboard.writeText(value);
                    return true;
                } catch (err) {
                    console.warn('Clipboard API copy failed, using fallback', err);
                }
            }

            const textarea = document.createElement('textarea');
            textarea.value = value;
            textarea.setAttribute('readonly', '');
            textarea.style.position = 'fixed';
            textarea.style.top = '0';
            textarea.style.left = '0';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.focus();
            textarea.select();
            textarea.setSelectionRange(0, textarea.value.length);
            let copied = false;
            try {
                copied = typeof document.execCommand === 'function' && document.execCommand('copy');
            } catch (err) {
                console.warn('Clipboard fallback failed', err);
            } finally {
                document.body.removeChild(textarea);
            }
            return copied;
        }

        function showPlatformModal() {
            platformModal?.classList.remove('hidden');
            platformModal?.classList.add('flex');
        }

        function hidePlatformModal() {
            platformModal?.classList.add('hidden');
            platformModal?.classList.remove('flex');
        }

        function resetSelectedDeviceFilter() {
            selectedDevice = 'all';
            if (deviceFilter) deviceFilter.value = 'all';
        }

        function resetDeviceStatusUi(platform = activePlatform) {
            resetSelectedDeviceFilter();
            if (deviceFilter) {
                deviceFilter.replaceChildren(new Option('All Devices', 'all'));
            }
            if (!deviceListEl) return;
            const message = platform === 'ios'
                ? 'Waiting... (iOS: connect a trusted USB device)'
                : 'Waiting... (ADB: connect an Android device)';
            deviceListEl.replaceChildren();
            const waiting = document.createElement('p');
            waiting.className = 'text-orange-500';
            waiting.textContent = message;
            deviceListEl.appendChild(waiting);
        }

        function syncPlatformUi() {
            const platform = activePlatform === 'ios' ? 'ios' : 'android';
            document.querySelectorAll('.sdk-check-panel').forEach(panel => {
                panel.classList.toggle('hidden', panel.getAttribute('data-sdk-platform') !== platform);
            });
            const sdkPresetPanel = document.getElementById('sdkCheckPresetPanel');
            if (sdkPresetPanel) sdkPresetPanel.classList.remove('hidden');
            document.querySelectorAll('[data-ios-hidden-source="appmetrica"]').forEach(el => {
                el.classList.toggle('hidden', platform === 'ios');
                const radio = el.querySelector('input[type="radio"]');
                if (platform === 'ios' && radio?.checked) {
                    const allRadio = document.querySelector('input[name="validatorSourceFilter"][value="all"]');
                    if (allRadio) allRadio.checked = true;
                }
            });
            const packageLabel = document.getElementById('packageIdInputLabel');
            const packageInput = document.getElementById('packageIdInput');
            if (packageLabel) packageLabel.textContent = platform === 'ios' ? 'Bundle Search:' : 'Package ID:';
            if (packageInput) packageInput.placeholder = platform === 'ios' ? 'PixelArt' : 'com.example.app';
            document.querySelectorAll('.package-id-checkbox').forEach(cb => {
                const label = cb.closest('label');
                const text = label?.querySelector('span');
                const platformValue = platform === 'ios' ? cb.dataset.iosValue : cb.dataset.androidValue;
                const quickLabel = platform === 'ios' ? cb.dataset.iosLabel : cb.dataset.androidLabel;
                if (label) label.classList.toggle('hidden', !platformValue);
                if (platformValue) cb.value = platformValue;
                if (text && quickLabel) text.textContent = quickLabel;
                if (!platformValue) cb.checked = false;
            });
            const sdkBadge = document.getElementById('sdkCheckPlatformBadge');
            if (sdkBadge) sdkBadge.textContent = `${platformLabel(platform)} Table`;

            const packageTagInput = document.getElementById('packageTagFilterInput');
            if (packageTagInput) {
                const isIos = platform === 'ios';
                packageTagInput.disabled = isIos;
                packageTagInput.placeholder = isIos ? 'Tag filter unavailable on iOS' : 'Tag...';
                packageTagInput.classList.toggle('bg-gray-100', isIos);
                packageTagInput.classList.toggle('text-gray-400', isIos);
                if (isIos) packageTagInput.value = '';
            }
            // Quick filters depend on the platform (keywords per platform; iOS
            // has no tag).
            renderPackageQuickFilters();
        }

        function resetSdkCheckUiState(clearInput = false) {
            sdkCheckRunning = false;
            if (clearInput) sdkCheckClearPending = false;
            const sdkBtn = document.getElementById('startSdkCheckBtn');
            if (sdkBtn) sdkBtn.textContent = 'Start Checking';
            document.getElementById('sdkCheckTableBody')?.replaceChildren();
            const iosBody = document.getElementById('sdkCheckIosTableBody');
            if (iosBody) iosBody.innerHTML = '<tr><td class="py-2 px-4 text-sm text-gray-500 italic">iOS table ready. Device connection will be added in the next step.</td></tr>';
            if (clearInput) {
                const sdkInput = document.getElementById('sdkCheckInput');
                if (sdkInput) sdkInput.value = '';
                document.querySelectorAll('input[name="sdkCheckPreset"]').forEach(input => input.checked = false);
                const presetStatus = document.getElementById('sdkCheckPresetStatus');
                if (presetStatus) presetStatus.textContent = 'Chọn preset để tự nạp danh sách.';
            }
        }

        function resetRuntimeUiForPlatformSwitch() {
            resetSdkCheckUiState(true);
            if (typeof setPackagePauseState === 'function') setPackagePauseState(false);
            if (typeof closePackageStreamModal === 'function') closePackageStreamModal();
            if (typeof setPackageRunningState === 'function') setPackageRunningState(false);
            if (typeof resetPackageLogUiState === 'function') resetPackageLogUiState();
            if (typeof resetPriceRotationUiState === 'function') resetPriceRotationUiState();
            resetSelectedDeviceFilter();
            syncPlatformUi();
            resetDeviceStatusUi(activePlatform);
        }

        function renderInstallationIdPanel() {
            if (!installationIdGameEl || !installationIdValueEl) return;
            const packageId = installationIdState.package_id || '';
            const gameCode = installationIdState.game_code || '';
            const installationId = installationIdState.installation_id || '';
            installationIdGameEl.textContent = packageId ? `${gameCode || 'Unknown'} | ${packageId}` : '-';
            installationIdValueEl.value = installationId;
            installationIdValueEl.placeholder = installationId ? '' : '-';
            if (copyInstallationIdBtn) {
                copyInstallationIdBtn.disabled = !installationId;
                copyInstallationIdBtn.classList.toggle('opacity-50', !installationId);
                copyInstallationIdBtn.classList.toggle('cursor-not-allowed', !installationId);
            }
        }

        installationIdValueEl?.addEventListener('focus', () => {
            if (installationIdValueEl.value) installationIdValueEl.select();
        });

        copyInstallationIdBtn?.addEventListener('click', async () => {
            const installationId = installationIdState.installation_id || '';
            if (!installationId) return;
            const originalText = copyInstallationIdBtn.textContent;
            try {
                if (!await copyTextToClipboard(installationId)) {
                    throw new Error('clipboard_unavailable');
                }
                copyInstallationIdBtn.textContent = 'Copied';
            } catch (err) {
                alert('Copy failed: ' + err);
            }
            setTimeout(() => {
                copyInstallationIdBtn.textContent = originalText;
            }, 1200);
        });

        function setActivePlatform(platform, persist = true) {
            activePlatform = platform === 'ios' ? 'ios' : 'android';
            if (persist) localStorage.setItem('eventInspectorPlatform', activePlatform);
            if (platformBtn) platformBtn.textContent = `Platform: ${platformLabel(activePlatform)}`;
            resetBrightSdkUiState();
            resetRuntimeUiForPlatformSwitch();
            socket.emit('set_platform', { platform: activePlatform, reset: !!persist });
        }

        document.querySelectorAll('[data-platform-choice]').forEach(btn => {
            btn.addEventListener('click', () => {
                setActivePlatform(btn.getAttribute('data-platform-choice') || 'android');
                hidePlatformModal();
            });
        });

        platformBtn?.addEventListener('click', showPlatformModal);
        // The platform bootstrap calls several reset helpers whose state is
        // declared later in this script (package log, price rotation, and
        // Bright/SKAds).  Running it immediately leaves those `let` bindings
        // in the temporal dead zone and aborts the rest of the script.  Defer
        // bootstrap until the complete script has initialized so all platform
        // and device handlers are registered before the first reset.
        window.setTimeout(() => {
            if (activePlatform) setActivePlatform(activePlatform, false);
            else showPlatformModal();
        }, 0);

        // --- Tab Logic ---
        function setServicesCheckerStatus(message, state = 'idle') {
            const status = document.getElementById('servicesCheckerStatus');
            if (!status) return;
            status.textContent = message;
            status.className = state === 'error'
                ? 'text-sm text-red-700'
                : state === 'busy'
                    ? 'text-sm text-amber-700'
                    : 'text-sm text-slate-500';
        }

        function setServicesCheckerError(message = '') {
            const errorBox = document.getElementById('servicesCheckerError');
            if (!errorBox) return;
            errorBox.textContent = message;
            errorBox.classList.toggle('hidden', !message);
        }

        function showServicesCheckerFrame(url) {
            const frame = document.getElementById('servicesCheckerFrame');
            if (!frame) return;
            frame.src = `${url}/?eventinspector=${Date.now()}`;
            frame.dataset.started = 'true';
            frame.classList.remove('hidden');
        }

        async function openServicesChecker() {
            const frame = document.getElementById('servicesCheckerFrame');
            if (!frame) return;
            if (frame.dataset.started === 'true') return;

            setServicesCheckerError();
            setServicesCheckerStatus('Starting...', 'busy');
            try {
                const response = await fetch('/api/services-checker/start', { method: 'POST' });
                const data = await response.json();
                if (!response.ok || !data.ok || !data.url) {
                    throw new Error(data.error || 'services_checker_start_failed');
                }
                showServicesCheckerFrame(data.url);
                setServicesCheckerStatus('Ready.');
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                setServicesCheckerStatus('Services Checker unavailable.', 'error');
                setServicesCheckerError(`Unable to start Services Checker: ${message}`);
                frame.dataset.started = 'false';
            }
        }

        function switchTab(tabName) {
            currentTab = tabName;
            // Hide all contents with Safety Check
            ['tabContentLoadAds', 'tabContentLoadAdsExt', 'tabContentValidator', 'tabContentDefaultAdEvents', 'tabContentBrightSDK', 'tabContentSpecific', 'tabContentAdRevenue', 'tabContentCallbackAd', 'tabContentPriceRotation', 'tabContentSdkCheck', 'tabContentServicesChecker', 'tabContentPackage'].forEach(id => {
                const el = document.getElementById(id);
                if (el) el.classList.add('hidden');
                else console.warn('Missing element ID:', id);
            });
            // Deactivate all buttons
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            
            // Activate selected with Safety Check
            const contentEl = document.getElementById('tabContent' + tabName);
            if (contentEl) contentEl.classList.remove('hidden');
            
            const btnEl = document.getElementById('tabBtn' + tabName);
            if (btnEl) btnEl.classList.add('active');

            if (tabName === 'SdkCheck' || tabName === 'BrightSDK') loadSdkCheckPresetsFromGit();
            if (tabName === 'Specific' && specificLive) specificLive.onTabShown();
            if (tabName === 'Validator' && validatorLive) validatorLive.onTabShown();
            if (tabName === 'AdRevenue' && adRevenueLive) adRevenueLive.onTabShown();
            if (tabName === 'CallbackAd' && callbackLive) callbackLive.onTabShown();
            if (tabName === 'PriceRotation' && priceRotationLive) priceRotationLive.onTabShown();
            if (tabName === 'LoadAds' && loadAdsLive) loadAdsLive.onTabShown();
            if (tabName === 'LoadAdsExt' && loadAdsExtLive) loadAdsExtLive.onTabShown();
            if (tabName === 'ServicesChecker') openServicesChecker();
            if (['Validator', 'DefaultAdEvents', 'Specific', 'AdRevenue'].includes(tabName)) {
                // The profile is shared by all validation/revenue tabs. Sync
                // it on tab entry so a tab cannot keep a stale selector after
                // another tab changed the Default File.
                refreshProfiles().catch(() => {});
            }
            
            socket.emit('change_tab', { tab_name: tabName });
        }

        // --- Global Control Logic ---
        pauseBtn.addEventListener('click', () => {
            // Optimistic toggle for resizer usability
            const willPause = pauseBtn.textContent === 'Pause';
            setResizerEnabled(willPause);
            isPausedClient = willPause;
            socket.emit('toggle_pause');
        });
        clearAllBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to clear ALL logs?')) {
                resetSdkCheckUiState(true);
                sdkCheckClearPending = true;
                renderSdkCheckPresetOptions();
                resetBrightSdkUiState();
                resetPriceRotationUiState();
                socket.emit('clear_all_logs');
            }
        });

        restartAppBtn?.addEventListener('click', () => {
            if (restartAppBtn.disabled) return;
            const originalText = restartAppBtn.textContent;
            restartAppBtn.disabled = true;
            restartAppBtn.textContent = 'Checking...';
            fetch('/check_update', { method: 'POST' })
                .then(r => r.json())
                .then(data => {
                    if (data.status === 'updated') {
                        restartAppBtn.textContent = 'Restarting...';
                        return fetch('/restart_app', { method: 'POST' });
                    }
                    if (data.status === 'up_to_date') {
                        alert('Already up to date.');
                    } else if (data.error) {
                        alert('Update check failed: ' + data.error);
                    } else {
                        alert('Update check failed.');
                    }
                })
                .catch(err => {
                    alert('Update check failed: ' + err);
                })
                .finally(() => {
                    restartAppBtn.disabled = false;
                    restartAppBtn.textContent = originalText;
                });
        });

        const manualRestartBtn = document.getElementById('manualRestartBtn');
        manualRestartBtn?.addEventListener('click', () => {
            if (manualRestartBtn.disabled) return;
            const originalText = manualRestartBtn.textContent;
            manualRestartBtn.disabled = true;
            manualRestartBtn.textContent = 'Restarting...';
            fetch('/restart_app', { method: 'POST' })
                .then(async response => {
                    const data = await response.json().catch(() => ({}));
                    if (!response.ok || data.ok === false) {
                        throw new Error(data.error || `HTTP ${response.status}`);
                    }
                    if (data.mode === 'local_runtime_reset') {
                        window.location.reload();
                    }
                    return data;
                })
                .catch(err => {
                    alert('Restart failed: ' + err);
                })
                .finally(() => {
                    manualRestartBtn.disabled = false;
                    manualRestartBtn.textContent = originalText;
                });
        });
        
        // --- Recording Logic (Updated for separate tabs) ---
        function toggleRecord(tabName) {
            const btn = document.getElementById('btnRecord_' + tabName);
            const input = document.getElementById('sheetName_' + tabName);
            
            if (!btn.classList.contains('bg-red-600')) {
                const name = input.value.trim();
                socket.emit('toggle_record', { sheet_name: name, tab_name: tabName });
            } else { 
                socket.emit('toggle_record', { tab_name: tabName }); 
            }
        }

        function safeRecordSheetName(value) {
            const text = String(value || 'Running').trim();
            if (!text || text.length > 200 || /<!doctype|<html|<head/i.test(text)) {
                return 'Running';
            }
            return text;
        }

        socket.on('record_status', (s) => {
            // Update UI based on which tab triggered the status
            const tabName = s.tab_name;
            const btn = document.getElementById('btnRecord_' + tabName);
            if (btn) {
                if (s.is_recording) {
                    btn.textContent = 'Stop Recording: ' + safeRecordSheetName(s.current_sheet);
                    btn.className = 'bg-red-600 text-white font-bold py-2 px-4 rounded animate-record shadow-lg text-sm';
                } else {
                    btn.textContent = 'Start Record';
                    btn.className = 'bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded shadow text-sm';
                }
            }
        });

        let isPausedClient = false;
        function setResizerEnabled(enabled) {
            document.querySelectorAll('.resizer').forEach(r => {
                r.classList.toggle('disabled', !enabled);
                r.style.pointerEvents = enabled ? 'auto' : 'none';
            });
        }

        socket.on('pause_status', (data) => {
            isPausedClient = !!data.is_paused;
            if (data.is_paused) {
                pauseBtn.textContent = 'Resume';
                pauseBtn.classList.replace('bg-yellow-500', 'bg-blue-500');
            } else {
                pauseBtn.textContent = 'Pause';
                pauseBtn.classList.replace('bg-blue-500', 'bg-yellow-500');
            }
            setResizerEnabled(isPausedClient);
        });

        // --- Rendering Helpers ---
        function escapeHTML(str) {
            if (!str) return '';
            const p = document.createElement("p");
            p.textContent = str;
            return p.innerHTML.replace(/  /g, ' &nbsp;');
        }
        
        // FIX: Better attribute escaping for JSON buttons
        function escapeAttribute(str) {
            if (!str) return '{}';
            // Replace single quotes with HTML entity to prevent breaking the attribute
            return str.replace(/'/g, '&#39;');
        }

        function logRecordKey(record) {
            const device = String(record?.device_id || '').trim();
            const raw = String(record?.raw_log ?? record?.log ?? '').trim();
            if (raw) return `${device}||${raw}`;
            return `${device}||${JSON.stringify(record || {})}`;
        }

        function dedupeLogRecords(records) {
            const seen = new Set();
            return (Array.isArray(records) ? records : []).filter(record => {
                const key = logRecordKey(record);
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }

        function renderSimpleTable(id, data) {
            const tbody = document.getElementById(id);
            if (!tbody) return;
            const uniqueData = dedupeLogRecords(data);
            const filtered = (selectedDevice === 'all') ? uniqueData : uniqueData.filter(e => e.device_id === selectedDevice);
            
            const useProvider = id === 'loadAdsTableBody' || id === 'loadAdsExtTableBody';
            const useAdNetwork = id === 'loadAdsExtTableBody';
            const columnCount = useProvider ? 5 : 4;
            if (filtered.length === 0) {
                 tbody.innerHTML = `<tr><td colspan="${columnCount}" class="text-center py-4 text-gray-400">Waiting for recording...</td></tr>`;
                 return;
            }

            tbody.innerHTML = filtered.map(e => `
                <tr class="hover:bg-gray-50 border-b text-sm">
                    <td class="py-2 px-2 text-purple-700 text-sm font-medium whitespace-nowrap">${escapeHTML(e.device_name || '')}</td>
                    ${useProvider ? `<td class="py-2 px-2 text-orange-600 text-sm font-medium whitespace-nowrap">${escapeHTML(e.provider || (useAdNetwork ? 'Unknown' : ''))}</td>` : ''}
                    <td class="py-2 px-2 text-blue-600 text-sm font-medium whitespace-nowrap">${escapeHTML(useAdNetwork ? (e.ad_network || e.ad_source || '') : (e.ad_source || e.ad_network || ''))}</td>
                    <td class="py-2 px-2 text-green-600 text-sm font-medium whitespace-nowrap">${escapeHTML(e.ad_format || '')}</td>
                    <td class="py-2 px-3 log-cell text-xs font-normal text-gray-600">${escapeHTML(e.raw_log || '')}</td>
                </tr>
            `).join('');
        }

        // Load Ads / Load Ads Ext as live tables.  The row and placeholder
        // templates below are the ones renderSimpleTable uses, so the HTML of a
        // full render and of appended rows is identical.
        let lastLoadAdsData = [];
        let lastLoadAdsExtData = [];
        function simpleTableRowHtml(id) {
            const useProvider = id === 'loadAdsTableBody' || id === 'loadAdsExtTableBody';
            const useAdNetwork = id === 'loadAdsExtTableBody';
            return e => `
                <tr class="hover:bg-gray-50 border-b text-sm">
                    <td class="py-2 px-2 text-purple-700 text-sm font-medium whitespace-nowrap">${escapeHTML(e.device_name || '')}</td>
                    ${useProvider ? `<td class="py-2 px-2 text-orange-600 text-sm font-medium whitespace-nowrap">${escapeHTML(e.provider || (useAdNetwork ? 'Unknown' : ''))}</td>` : ''}
                    <td class="py-2 px-2 text-blue-600 text-sm font-medium whitespace-nowrap">${escapeHTML(useAdNetwork ? (e.ad_network || e.ad_source || '') : (e.ad_source || e.ad_network || ''))}</td>
                    <td class="py-2 px-2 text-green-600 text-sm font-medium whitespace-nowrap">${escapeHTML(e.ad_format || '')}</td>
                    <td class="py-2 px-3 log-cell text-xs font-normal text-gray-600">${escapeHTML(e.raw_log || '')}</td>
                </tr>
            `;
        }
        function simpleTableEmptyHtml(id) {
            const useProvider = id === 'loadAdsTableBody' || id === 'loadAdsExtTableBody';
            const columnCount = useProvider ? 5 : 4;
            return `<tr><td colspan="${columnCount}" class="text-center py-4 text-gray-400">Waiting for recording...</td></tr>`;
        }
        function simpleTableLive(id, tabContentId, getRows, setRows) {
            return createLiveTable({
                tbodyId: id,
                tabContentId,
                emptyHtml: () => simpleTableEmptyHtml(id),
                getRows,
                setRows,
                rowHtml: simpleTableRowHtml(id),
                makeFilter: () => (e => selectedDevice === 'all' || e.device_id === selectedDevice),
            });
        }
        var loadAdsLive = simpleTableLive('loadAdsTableBody', 'tabContentLoadAds', () => lastLoadAdsData, rows => { lastLoadAdsData = rows; });
        var loadAdsExtLive = simpleTableLive('loadAdsExtTableBody', 'tabContentLoadAdsExt', () => lastLoadAdsExtData, rows => { lastLoadAdsExtData = rows; });

        let validator_results_cache = [];
        let defaultEventNames = {{ default_event_names | tojson }};
        let defaultEventStatusEls = {};
        let currentProfileName = {{ current_profile_name | tojson }};
        let currentProfileRevision = {{ profile_revision | tojson }};
        let profileRefreshRequestId = 0;
        let defaultAdEventData = {{ default_ad_event_data | tojson }};

        function defaultAdEventHitKey(adFormat, provider, eventName) {
            return [adFormat, provider, String(eventName || '').trim().toLowerCase()].join('|');
        }

        function defaultAdEventSimpleHitKey(category, provider, eventName) {
            return [category, provider, String(eventName || '').trim().toLowerCase()].join('|');
        }

        function renderDefaultAdEventMatrix() {
            const container = document.getElementById('defaultAdEventMatrix');
            if (!container) return;
            const data = defaultAdEventData || {};
            const providers = Array.isArray(data.providers) ? data.providers : [];
            const formats = Array.isArray(data.formats) ? data.formats : [];
            const simpleCategories = Array.isArray(data.simple_categories) ? data.simple_categories : [];
            const hits = new Set(
                (Array.isArray(data.hits) ? data.hits : [])
                    .filter(hit => selectedDevice === 'all' || hit.device_id === selectedDevice)
                    .map(hit => defaultAdEventHitKey(hit.ad_format, hit.provider, hit.event_name))
            );
            const simpleHits = new Set(
                (Array.isArray(data.simple_hits) ? data.simple_hits : [])
                    .filter(hit => selectedDevice === 'all' || hit.device_id === selectedDevice)
                    .map(hit => defaultAdEventSimpleHitKey(hit.category, hit.provider, hit.event_name))
            );

            const hasFormats = formats.some(section => Array.isArray(section.events) && section.events.length);
            const hasSimpleCategories = simpleCategories.some(section => Array.isArray(section.events) && section.events.length);
            if (!hasFormats && !hasSimpleCategories) {
                container.innerHTML = '<div class="bg-white rounded-xl shadow-md p-4 text-sm text-gray-400">No default ad events configured</div>';
                return;
            }

            const coverageFor = (section, providerKey) => {
                const events = Array.isArray(section.events) ? section.events : [];
                const configured = events.filter(event => (event.providers || []).includes(providerKey));
                const checked = configured.filter(event => hits.has(defaultAdEventHitKey(section.key, providerKey, event.name)));
                return {checked: checked.length, total: configured.length};
            };
            const totalCoverageFor = (providerKey) => formats.reduce((summary, section) => {
                const current = coverageFor(section, providerKey);
                summary.checked += current.checked;
                summary.total += current.total;
                return summary;
            }, {checked: 0, total: 0});
            const summaryCards = providers.map(provider => {
                const summary = totalCoverageFor(provider.key);
                return `<div class="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 min-w-[150px]">
                    <div class="text-xs font-semibold text-slate-600">${escapeHTML(provider.label || provider.key || '')}</div>
                    <div class="text-lg font-bold ${summary.checked ? 'text-green-600' : 'text-slate-500'}">${summary.checked}/${summary.total}</div>
                    <div class="text-[11px] text-slate-500">matched events</div>
                </div>`;
            }).join('');
            const sections = formats.map(section => {
                const events = Array.isArray(section.events) ? section.events : [];
                if (!events.length) return '';
                const rows = events.map(event => {
                    const configuredProviders = new Set(event.providers || []);
                    return `<tr class="border-b last:border-b-0 hover:bg-gray-50">
                        <td class="py-2 px-3 text-sm font-semibold text-gray-800 whitespace-nowrap">${escapeHTML(event.name || '')}</td>
                        ${providers.map(provider => {
                            const configured = configuredProviders.has(provider.key);
                            const checked = configured && hits.has(defaultAdEventHitKey(section.key, provider.key, event.name));
                            return `<td class="py-2 px-3 text-center text-lg ${checked ? 'text-green-600 font-bold' : 'text-gray-300'}">${checked ? '✓' : ''}</td>`;
                        }).join('')}
                    </tr>`;
                }).join('');
                const sectionCoverage = providers.map(provider => {
                    const summary = coverageFor(section, provider.key);
                    return `<span class="text-[10px] font-normal text-slate-500">${summary.checked}/${summary.total}</span>`;
                });
                return `<section class="bg-white rounded-xl shadow-md p-4">
                    <h3 class="text-base font-bold text-gray-800 mb-3">${escapeHTML(section.label || section.key || '')}</h3>
                    <div class="overflow-x-auto">
                        <table class="min-w-full bg-white">
                            <thead class="bg-gray-50">
                                <tr>
                                    <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Event</th>
                                    ${providers.map((provider, index) => `<th class="text-center text-sm font-semibold text-gray-600 py-2 px-3 border-b">${escapeHTML(provider.label || provider.key || '')}<br>${sectionCoverage[index]}</th>`).join('')}
                                </tr>
                            </thead>
                            <tbody>${rows}</tbody>
                        </table>
                    </div>
                </section>`;
            }).join('');
            const simpleSummaryCards = simpleCategories.reduce((html, category) => {
                const providersForCategory = Array.isArray(category.providers) ? category.providers : [];
                const events = Array.isArray(category.events) ? category.events : [];
                const cards = providersForCategory.map(provider => {
                    const configured = events.filter(event => (event.providers || []).includes(provider.key));
                    const checked = configured.filter(event => simpleHits.has(defaultAdEventSimpleHitKey(category.key, provider.key, event.name)));
                    return `<div class="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 min-w-[150px]">
                        <div class="text-xs font-semibold text-slate-600">${escapeHTML(provider.label || provider.key || '')}</div>
                        <div class="text-lg font-bold ${checked.length ? 'text-green-600' : 'text-slate-500'}">${checked.length}/${configured.length}</div>
                        <div class="text-[11px] text-slate-500">matched events</div>
                    </div>`;
                }).join('');
                return html + cards;
            }, '');
            // Simple providers are intentionally rendered as separate blocks.
            // Their workbook sheets contain one event list per platform, so a
            // shared category table would incorrectly mix unrelated events.
            const simpleSections = simpleCategories.reduce((html, category) => {
                const categoryProviders = Array.isArray(category.providers) ? category.providers : [];
                const events = Array.isArray(category.events) ? category.events : [];
                const providerBlocks = categoryProviders.map(provider => {
                    const configuredEvents = events.filter(event => (event.providers || []).includes(provider.key));
                    if (!configuredEvents.length) return '';
                    const checkedCount = configuredEvents.filter(event => simpleHits.has(defaultAdEventSimpleHitKey(category.key, provider.key, event.name))).length;
                    const rows = configuredEvents.map(event => {
                        const checked = simpleHits.has(defaultAdEventSimpleHitKey(category.key, provider.key, event.name));
                        return `<tr class="border-b last:border-b-0 hover:bg-gray-50">
                            <td class="py-2 px-3 text-sm font-semibold text-gray-800 whitespace-nowrap">${escapeHTML(event.name || '')}</td>
                            <td class="py-2 px-3 text-center text-lg ${checked ? 'text-green-600 font-bold' : 'text-gray-300'}">${checked ? '✓' : ''}</td>
                        </tr>`;
                    }).join('');
                    return `<section class="bg-white rounded-xl shadow-md p-4">
                        <h3 class="text-base font-bold text-gray-800 mb-3">${escapeHTML(provider.label || provider.key || '')}</h3>
                        <div class="overflow-x-auto">
                            <table class="min-w-full bg-white">
                                <thead class="bg-gray-50">
                                    <tr>
                                        <th class="text-left text-sm font-semibold text-gray-600 py-2 px-3 border-b">Event</th>
                                        <th class="text-center text-sm font-semibold text-gray-600 py-2 px-3 border-b">${escapeHTML(provider.label || provider.key || '')}<br><span class="text-[10px] font-normal">${checkedCount}/${configuredEvents.length}</span></th>
                                    </tr>
                                </thead>
                                <tbody>${rows}</tbody>
                            </table>
                        </div>
                    </section>`;
                }).join('');
                return html + providerBlocks;
            }, '');
            container.innerHTML = `<div class="flex flex-wrap gap-3">${summaryCards}${simpleSummaryCards}</div>${sections}${simpleSections}`;
        }

        function updateProfileStatus(payload) {
            const gameEl = document.getElementById('profileGameText');
            if (gameEl) {
                gameEl.textContent = payload.game_name || 'Unknown';
            }
        }

        function brightSkadPresetEntries() {
            return Object.entries(sdkCheckPresets || {})
                .filter(([, preset]) => {
                    if (!preset || !preset.bright) return false;
                    return preset.platform === 'ios' || preset.platform === 'all';
                });
        }

        function updateBrightSkadPresetStatus() {
            const select = document.getElementById('brightSkadPresetSelect');
            const status = document.getElementById('brightSkadPresetStatus');
            if (!select || !status) return;
            const name = select.value || '';
            const bright = sdkCheckPresets?.[name]?.bright;
            if (!name || !bright) {
                status.textContent = 'Chưa có preset Bright/SKAds.';
                return;
            }
            const identifiers = Array.isArray(bright.skadnetwork_identifiers) ? bright.skadnetwork_identifiers : [];
            status.textContent = `${name}: BrightData ${bright.version || 'Not set'} · ${identifiers.length} SKAdNetwork entries`;
        }

        function renderBrightSkadPresetOptions() {
            const select = document.getElementById('brightSkadPresetSelect');
            if (!select) return;
            const entries = brightSkadPresetEntries();
            const current = select.value || '';
            const defaultName = entries.some(([name]) => name === 'C-192-iOS')
                ? 'C-192-iOS'
                : (entries[0]?.[0] || '');
            select.innerHTML = entries.length
                ? entries.map(([name, preset]) => {
                    const count = Array.isArray(preset?.bright?.skadnetwork_identifiers)
                        ? preset.bright.skadnetwork_identifiers.length
                        : 0;
                    return `<option value="${escapeAttribute(name)}">${escapeHTML(name)} · ${count} SKAdNetwork</option>`;
                }).join('')
                : '<option value="">No Bright/SKAds preset</option>';
            const available = entries.map(([name]) => name);
            select.value = available.includes(current) ? current : defaultName;
            select.disabled = entries.length === 0;
            updateBrightSkadPresetStatus();
            if (brightSkadLastPayload && select.value) {
                renderBrightSkadComparison(buildBrightSkadComparison(brightSkadLastPayload, select.value));
            }
        }

        function buildBrightSkadComparison(payload, presetName) {
            const bright = sdkCheckPresets?.[presetName]?.bright;
            if (!bright) {
                return {preset_name: presetName || '', rows: [], extra_identifiers: [], passed_count: 0, total_count: 0};
            }
            const expectedVersion = String(bright.version || '').trim();
            const expectedIdentifiers = [];
            const expectedKeys = new Set();
            (Array.isArray(bright.skadnetwork_identifiers) ? bright.skadnetwork_identifiers : []).forEach(value => {
                const identifier = String(value || '').trim();
                const key = identifier.toLowerCase();
                if (!identifier || expectedKeys.has(key)) return;
                expectedKeys.add(key);
                expectedIdentifiers.push(identifier);
            });
            const rawVersion = String(payload?.brightsdk_version || '').trim();
            const actualVersion = rawVersion.toLowerCase() === 'not found' ? '' : rawVersion;
            const actualIdentifiers = [];
            const actualKeys = new Set();
            (Array.isArray(payload?.skadnetwork_identifiers) ? payload.skadnetwork_identifiers : []).forEach(value => {
                const identifier = String(value || '').trim();
                const key = identifier.toLowerCase();
                if (!identifier || actualKeys.has(key)) return;
                actualKeys.add(key);
                actualIdentifiers.push(identifier);
            });
            const rows = [];
            if (expectedVersion) {
                rows.push({
                    kind: 'BrightData',
                    label: 'BrightData',
                    actual: actualVersion || 'NOT FOUND',
                    expected: expectedVersion,
                    status: actualVersion === expectedVersion ? 'PASSED' : (actualVersion ? 'FAILED' : 'MISSING'),
                });
            }
            expectedIdentifiers.forEach(identifier => {
                const found = actualKeys.has(identifier.toLowerCase());
                rows.push({
                    kind: 'SKAdNetwork',
                    label: identifier,
                    actual: found ? identifier : 'NOT FOUND',
                    expected: identifier,
                    status: found ? 'PASSED' : 'MISSING',
                });
            });
            const extraIdentifiers = actualIdentifiers.filter(identifier => !expectedKeys.has(identifier.toLowerCase()));
            return {
                preset_name: presetName || '',
                rows,
                extra_identifiers: extraIdentifiers,
                expected_identifier_count: expectedIdentifiers.length,
                actual_identifier_count: actualIdentifiers.length,
                passed_count: rows.filter(row => row.status === 'PASSED').length,
                total_count: rows.length,
            };
        }

        function renderBrightSkadComparison(compare) {
            const body = document.getElementById('brightSkadCompareBody');
            const summary = document.getElementById('brightSkadCompareSummary');
            const warning = document.getElementById('brightSkadExtraWarning');
            if (!body || !summary || !warning) return;
            const rows = Array.isArray(compare?.rows) ? compare.rows : [];
            if (!rows.length) {
                body.innerHTML = '<tr><td colspan="4" class="py-3 px-3 text-center text-slate-500 italic">Chưa có preset/IPA để compare.</td></tr>';
                summary.textContent = 'Upload IPA để bắt đầu compare.';
                warning.classList.add('hidden');
                warning.innerHTML = '';
                return;
            }
            const passedCount = Number.isFinite(compare.passed_count)
                ? compare.passed_count
                : rows.filter(row => row.status === 'PASSED').length;
            summary.textContent = `${passedCount}/${rows.length} PASSED · ${compare.preset_name || 'Bright/SKAds preset'}`;
            body.innerHTML = rows.map(row => {
                const status = String(row.status || 'MISSING').toUpperCase();
                const passed = status === 'PASSED';
                const failed = status === 'FAILED';
                const statusClass = passed ? 'text-green-600' : (failed ? 'text-red-600' : 'text-amber-700');
                const statusIcon = passed ? '✓' : (failed ? '✕' : '—');
                return `<tr class="border-b last:border-b-0 hover:bg-gray-50">
                    <td class="py-2 px-3 text-center font-semibold ${statusClass}" title="${escapeHTML(status)}">${statusIcon} <span class="text-[10px]">${escapeHTML(status)}</span></td>
                    <td class="py-2 px-3 font-semibold text-slate-800 break-all">${escapeHTML(row.label || row.kind || '')}</td>
                    <td class="py-2 px-3 font-mono text-slate-700 break-all">${escapeHTML(row.actual || '')}</td>
                    <td class="py-2 px-3 font-mono text-slate-700 break-all">${escapeHTML(row.expected || '')}</td>
                </tr>`;
            }).join('');
            const extras = Array.isArray(compare.extra_identifiers) ? compare.extra_identifiers : [];
            if (!extras.length) {
                warning.classList.add('hidden');
                warning.innerHTML = '';
            } else {
                warning.classList.remove('hidden');
                warning.innerHTML = `<div class="font-semibold mb-1">Warning: ${extras.length} identifier(s) có trong IPA nhưng không có trong preset.</div><div class="font-mono break-all">${extras.map(value => escapeHTML(value)).join('<br>')}</div>`;
            }
        }

        function resetBrightSdkUiState() {
            brightSkadLastPayload = null;
            const appNameEl = document.getElementById('brightSdkAppNameValue');
            if (appNameEl) appNameEl.textContent = 'Not Found';
            const versionEl = document.getElementById('brightSdkVersionValue');
            if (versionEl) versionEl.textContent = 'Not Found';
            renderBrightSkadComparison({rows: [], extra_identifiers: []});
            const inputEl = document.getElementById('brightSdkFileInput');
            if (inputEl) inputEl.value = '';
        }

        function renderProfileOptions(payload) {
            if (!payload || typeof payload !== 'object') return;
            const incomingRevision = Number(payload.profile_revision);
            if (Number.isFinite(incomingRevision) && incomingRevision < currentProfileRevision) {
                return;
            }
            if (Number.isFinite(incomingRevision)) {
                currentProfileRevision = incomingRevision;
            }
            const profiles = Array.isArray(payload.profiles) ? payload.profiles : [];
            const selectedProfile = payload.current_profile || '';
            [document.getElementById('profileSelect'), document.getElementById('adRevenueProfileSelect'), document.getElementById('defaultAdEventProfileSelect')].forEach(select => {
                if (!select) return;
                select.innerHTML = profiles.length
                    ? profiles.map(name => `<option value="${escapeAttribute(name)}">${escapeHTML(name)}</option>`).join('')
                    : '<option value="">No profiles</option>';
                // Keep one canonical selection across every tab. Setting the
                // property (rather than relying only on the selected HTML
                // attribute) also handles selectors populated after the user
                // switched tabs.
                select.value = selectedProfile;
                select.disabled = profiles.length === 0;
            });
            currentProfileName = selectedProfile;
            defaultEventNames = payload.default_event_names || [];
            defaultAdEventData = payload.default_ad_event_data || {providers: [], formats: [], hits: [], simple_categories: [], simple_hits: []};
            renderDefaultEventStatusList();
            updateDefaultEventStatus(validator_results_cache);
            renderDefaultAdEventMatrix();
            updateProfileStatus(payload);
            const adRevenueGameEl = document.getElementById('adRevenueProfileGameText');
            if (adRevenueGameEl) {
                adRevenueGameEl.textContent = payload.game_name || 'Unknown';
            }
            const defaultAdEventGameEl = document.getElementById('defaultAdEventProfileGameText');
            if (defaultAdEventGameEl) {
                defaultAdEventGameEl.textContent = payload.game_name || 'Unknown';
            }
        }

        async function refreshProfiles() {
            const requestId = ++profileRefreshRequestId;
            const res = await fetch(`/api/profiles?ts=${Date.now()}-${requestId}`, { cache: 'no-store' });
            const payload = await res.json();
            if (requestId !== profileRefreshRequestId) return payload;
            renderProfileOptions(payload);
            return payload;
        }

        function renderDefaultEventStatusList() {
            const container = document.getElementById('defaultEventStatusList');
            if (!container) return;
            if (!defaultEventNames || defaultEventNames.length === 0) {
                container.innerHTML = '<div class="text-gray-400">No default events</div>';
                return;
            }
            container.innerHTML = defaultEventNames.map(name => `
                <div class="default-event-item inline-flex items-center gap-1.5 px-2 py-1 rounded-md border border-gray-200 bg-white cursor-pointer hover:bg-gray-50"
                     data-event-name="${escapeAttribute(name)}" title="Click to filter by this event">
                    <span class="event-status-icon text-gray-400 font-bold w-4 text-center" data-event="${name}" title="checking">...</span>
                    <span class="truncate max-w-[180px] text-xs font-medium text-gray-700" title="${escapeHTML(name)}">${escapeHTML(name)}</span>
                </div>
            `).join('');
            defaultEventStatusEls = {};
            container.querySelectorAll('.event-status-icon').forEach(el => {
                defaultEventStatusEls[el.dataset.event] = el;
            });
        }

        function updateDefaultEventStatus(results) {
            if (!defaultEventNames || defaultEventNames.length === 0) return;
            const statusMap = {};
            defaultEventNames.forEach(n => { statusMap[n] = 'pending'; });
            results.forEach(r => {
                if (!statusMap.hasOwnProperty(r.event_name)) return;
                if (r.status === 'FAILED') statusMap[r.event_name] = 'failed';
                else if (r.status === 'PASSED' && statusMap[r.event_name] === 'pending') statusMap[r.event_name] = 'passed';
            });
            Object.entries(statusMap).forEach(([name, status]) => {
                const el = defaultEventStatusEls[name];
                if (!el) return;
                // Write only real changes: rewriting an unchanged icon still
                // forces the browser to lay out the whole page again.
                const setIcon = (text, className, title) => {
                    if (el.textContent !== text) el.textContent = text;
                    if (el.className !== className) el.className = className;
                    if (el.title !== title) el.title = title;
                };
                if (status === 'passed') {
                    setIcon('✓', 'event-status-icon text-green-600 font-bold', 'PASSED');
                } else if (status === 'failed') {
                    setIcon('✕', 'event-status-icon text-red-600 font-bold', 'FAILED');
                } else {
                    setIcon('...', 'event-status-icon text-gray-400 font-bold', 'checking');
                }
            });
        }

        // One row of the Validator table.  Full renders and appended rows use
        // this same template, so both paths produce identical HTML.
        function validatorRowHtml(res) {
             return `
                    <tr class="hover:bg-gray-50 border-b text-sm">
                        <td class="py-2 px-3 text-purple-700 text-sm">${res.device_name}</td>
                        <td class="py-2 px-3 text-sm font-semibold ${res.status === 'PASSED' ? 'text-green-600' : 'text-red-600'}">${res.status}</td>
                        <td class="py-2 px-3"><span class="event-name-link cursor-pointer text-sm font-medium text-indigo-700 hover:underline" data-event-name="${escapeAttribute(res.event_name)}">${res.event_name}</span></td>
                        <td class="py-2 px-3 details-cell text-sm">${res.details}</td>
                        <td class="py-2 px-3 log-cell text-xs font-normal text-gray-600">${escapeHTML(res.raw_log || '')}</td>
                        <td class="py-2 px-3"><button class="view-json-btn text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 font-medium py-1 px-2 rounded" data-json='${escapeAttribute(res.json_data)}'>View JSON</button></td>
                    </tr>
                 `;
        }

        function validatorRowVisible(r) {
             const filterText = (document.getElementById('validatorEventFilterInput')?.value || '').toLowerCase().trim();
             const rawFilterText = (document.getElementById('validatorRawFilterInput')?.value || '').toLowerCase().trim();
             const sourceFilter = document.querySelector('input[name="validatorSourceFilter"]:checked')?.value || 'all';
                              if (selectedDevice !== 'all' && r.device_id !== selectedDevice) return false;
                 if (filterText && !(r.event_name || '').toLowerCase().includes(filterText)) return false;
                 if (rawFilterText && !(r.raw_log || '').toLowerCase().includes(rawFilterText)) return false;
                 if (sourceFilter !== 'all' && (r.source || 'firebase') !== sourceFilter) return false;
                 return true;
        }

        var validatorLive = createLiveTable({
             tbodyId: 'validatorTableBody',
             tabContentId: 'tabContentValidator',
             emptyHtml: '<tr><td colspan="6" class="text-center py-4">Waiting...</td></tr>',
             getRows: () => validator_results_cache,
             setRows: rows => { validator_results_cache = rows; },
             rowHtml: validatorRowHtml,
             visible: validatorRowVisible,
        });

        function renderValidatorTable(results) {
             validatorLive.renderFull(results);
        }

        renderDefaultAdEventMatrix();
        renderDefaultEventStatusList();
        refreshProfiles().catch(() => {});

        // Click default event to fill filter input
        document.getElementById('defaultEventStatusList')?.addEventListener('click', (e) => {
            const item = e.target.closest('.default-event-item');
            if (!item) return;
            const name = item.getAttribute('data-event-name') || '';
            const input = document.getElementById('validatorEventFilterInput');
            if (input) {
                input.value = name;
                input.focus();
                renderValidatorTable(validator_results_cache);
            }
        });


        // Click event name in results table to fill filter input
        document.getElementById('validatorTableBody')?.addEventListener('click', (e) => {
            const el = e.target.closest('.event-name-link');
            if (!el) return;
            const name = el.getAttribute('data-event-name') || '';
            const input = document.getElementById('validatorEventFilterInput');
            if (input) {
                input.value = name;
                input.focus();
                renderValidatorTable(validator_results_cache);
            }
        });

        document.getElementById('specificEventTableBody')?.addEventListener('click', (e) => {
            const el = e.target.closest('.event-name-link');
            if (!el) return;
            const name = el.getAttribute('data-event-name') || '';
            const input = document.getElementById('specificEventInput');
            if (input) {
                input.value = name;
                input.focus();
                socket.emit('update_specific_filter', {
                    eventNames: [name],
                    params: []
                });
                renderSpecificEventTable();
            }
        });

        async function switchSharedProfile(profileName) {
            if (!profileName) return;
            if (profileName === currentProfileName) {
                await refreshProfiles();
                return;
            }
            const res = await fetch('/api/profiles/select', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ profile_name: profileName })
            });
            const payload = await res.json();
            if (!payload.ok) {
                alert(payload.error || 'Failed to switch profile');
                await refreshProfiles();
                return;
            }
            validator_results_cache = [];
            renderValidatorTable(validator_results_cache);
            renderProfileOptions(payload);
        }

        document.getElementById('profileSelect')?.addEventListener('change', async (e) => {
            await switchSharedProfile(e.target.value);
        });

        document.getElementById('adRevenueProfileSelect')?.addEventListener('change', async (e) => {
            await switchSharedProfile(e.target.value);
        });

        document.getElementById('defaultAdEventProfileSelect')?.addEventListener('change', async (e) => {
            await switchSharedProfile(e.target.value);
        });

        document.getElementById('reloadProfileBtn')?.addEventListener('click', async () => {
            const res = await fetch('/api/profiles/reload', { method: 'POST' });
            const payload = await res.json();
            if (!payload.ok) {
                alert(payload.error || 'Failed to reload profile');
                return;
            }
            renderProfileOptions(payload);
        });

        document.getElementById('openProfileFolderBtn')?.addEventListener('click', async () => {
            const res = await fetch('/api/profiles/open-folder', { method: 'POST' });
            const payload = await res.json();
            if (!payload.ok) {
                alert(payload.error || 'Failed to open profile folder');
            }
        });

        document.getElementById('importProfileBtn')?.addEventListener('click', () => {
            document.getElementById('profileFileInput')?.click();
        });

        document.getElementById('profileFileInput')?.addEventListener('change', async (e) => {
            const file = e.target.files && e.target.files[0];
            if (!file) return;
            const formData = new FormData();
            formData.append('profile_file', file);
            const res = await fetch('/api/profiles/import', {
                method: 'POST',
                body: formData
            });
            const payload = await res.json();
            e.target.value = '';
            if (!payload.ok) {
                alert(payload.error || 'Failed to import profile');
                return;
            }
            validator_results_cache = [];
            renderValidatorTable(validator_results_cache);
            renderProfileOptions(payload);
        });

        document.getElementById('importBrightSdkBtn')?.addEventListener('click', () => {
            document.getElementById('brightSdkFileInput')?.click();
        });

        document.getElementById('clearBrightSdkBtn')?.addEventListener('click', () => {
            resetBrightSdkUiState();
        });

        document.getElementById('brightSkadPresetSelect')?.addEventListener('change', () => {
            updateBrightSkadPresetStatus();
            const presetName = document.getElementById('brightSkadPresetSelect')?.value || '';
            if (brightSkadLastPayload && presetName) {
                renderBrightSkadComparison(buildBrightSkadComparison(brightSkadLastPayload, presetName));
            }
        });

        document.getElementById('brightSdkFileInput')?.addEventListener('change', async (e) => {
            const file = e.target.files && e.target.files[0];
            if (!file) return;
            const formData = new FormData();
            formData.append('ipa_file', file);
            formData.append('preset_name', document.getElementById('brightSkadPresetSelect')?.value || 'C-192-iOS');
            const res = await fetch('/api/ipa/inspect', {
                method: 'POST',
                body: formData
            });
            const payload = await res.json();
            e.target.value = '';
            const appNameEl = document.getElementById('brightSdkAppNameValue');
            const valueEl = document.getElementById('brightSdkVersionValue');
            if (!payload.ok) {
                if (appNameEl) appNameEl.textContent = 'Not Found';
                if (valueEl) valueEl.textContent = 'Not Found';
                brightSkadLastPayload = null;
                renderBrightSkadComparison({rows: [], extra_identifiers: []});
                alert(payload.error || 'Failed to inspect IPA');
                return;
            }
            if (appNameEl) appNameEl.textContent = payload.app_name || 'Not Found';
            if (valueEl) valueEl.textContent = payload.brightsdk_version || 'Not Found';
            brightSkadLastPayload = payload;
            const presetName = document.getElementById('brightSkadPresetSelect')?.value || payload.bright_compare?.preset_name || '';
            renderBrightSkadComparison(buildBrightSkadComparison(payload, presetName));
        });

        // --- Socket Listeners (Renderers) ---
        socket.on('update_load_ads', (d) => { lastLoadAdsData = dedupeLogRecords(d); loadAdsLive.setAll(lastLoadAdsData); });
        socket.on('update_load_ads_ext', (d) => { lastLoadAdsExtData = dedupeLogRecords(d); loadAdsExtLive.setAll(lastLoadAdsExtData); });
        socket.on('load_ads_delta', (payload) => loadAdsLive.applyDelta(payload));
        socket.on('load_ads_ext_delta', (payload) => loadAdsExtLive.applyDelta(payload));
        socket.on('validator_delta', (payload) => {
            if (validatorLive.applyDelta(payload)) updateDefaultEventStatus(validator_results_cache);
        });
        socket.on('update_validator_table', (d) => {
            validator_results_cache = d || [];
            validatorLive.setAll(validator_results_cache);
            updateDefaultEventStatus(validator_results_cache);
        });

        socket.on('update_default_ad_events', (d) => {
            defaultAdEventData = d || {providers: [], formats: [], hits: [], simple_categories: [], simple_hits: []};
            renderDefaultAdEventMatrix();
        });

        socket.on('profile_updated', (payload) => {
            if (payload && typeof payload === 'object') renderProfileOptions(payload);
        });

        let lastSpecificEventData = [];

        // --- Shared live-table updater (Specific Validator, Validator) ---
        // New rows arrive as server deltas.  The table appends them with the
        // same row template a full render uses, so the HTML is identical to a
        // full rebuild.  A hidden tab only records that it must render later.
        // A large visible table is patched at most once per second: with
        // auto table layout every patch re-lays out the whole table.
        var LIVE_TABLE_LARGE_ROWS = 300;
        var LIVE_TABLE_LARGE_INTERVAL_MS = 1000;
        function createLiveTable(opts) {
            const state = {
                lastSeq: 0,
                renderedCount: 0,
                needsRender: false,
                pendingAppend: [],
                pendingRemove: 0,
                flushTimer: null,
                lastFlushAt: 0,
                leadWs: null,
            };
            const seqOf = r => Number(r && r.seq) || 0;
            function tabVisible() {
                const el = document.getElementById(opts.tabContentId);
                return !!el && !el.classList.contains('hidden');
            }
            // Leading whitespace of one rendered row (multi-line templates
            // start with a newline).  Kept so a patched table serialises
            // byte-for-byte like a fresh full render.
            function leadingWhitespace(row) {
                const html = opts.rowHtml(row);
                let i = 0;
                while (i < html.length && html.charCodeAt(i) <= 32) i++;
                return html.slice(0, i);
            }
            function removeLeadingRows(tbody, count) {
                let removed = 0;
                while (removed < count && tbody.firstChild) {
                    const node = tbody.firstChild;
                    if (node.nodeType === 1) removed++;
                    node.remove();
                }
                // The text between the removed row and the next one holds the
                // removed row's trailing whitespace plus the next row's
                // leading whitespace; a full render starts with the latter only.
                const first = tbody.firstChild;
                if (first && first.nodeType === 3) first.data = state.leadWs || '';
            }
            function clearPending() {
                state.pendingAppend = [];
                state.pendingRemove = 0;
                if (state.flushTimer) { clearTimeout(state.flushTimer); state.flushTimer = null; }
            }
            function renderFull(rows) {
                const tbody = document.getElementById(opts.tbodyId);
                if (!tbody) return;
                clearPending();
                state.needsRender = false;
                const visible = opts.makeFilter ? opts.makeFilter() : opts.visible;
                const filtered = dedupeLogRecords(rows || opts.getRows()).filter(visible);
                state.renderedCount = filtered.length;
                state.lastFlushAt = performance.now();
                if (filtered.length && state.leadWs === null) state.leadWs = leadingWhitespace(filtered[0]);
                if (filtered.length === 0) { tbody.innerHTML = typeof opts.emptyHtml === 'function' ? opts.emptyHtml() : opts.emptyHtml; }
                else {
                    tbody.innerHTML = filtered.map(opts.rowHtml).join('');
                }
            }
            function flush() {
                state.flushTimer = null;
                if (!tabVisible() || state.needsRender) { state.needsRender = true; clearPending(); return; }
                const tbody = document.getElementById(opts.tbodyId);
                if (!tbody) return;
                // An empty table shows the placeholder row; rebuild in that case.
                if (state.renderedCount === 0 || state.pendingRemove >= state.renderedCount) { renderFull(); return; }
                const appendRows = state.pendingAppend;
                const removeCount = state.pendingRemove;
                state.pendingAppend = [];
                state.pendingRemove = 0;
                if (appendRows.length && state.leadWs === null) state.leadWs = leadingWhitespace(appendRows[0]);
                if (removeCount) removeLeadingRows(tbody, removeCount);
                if (appendRows.length) tbody.insertAdjacentHTML('beforeend', appendRows.map(opts.rowHtml).join(''));
                state.renderedCount += appendRows.length - removeCount;
                state.lastFlushAt = performance.now();
            }
            function scheduleFlush() {
                if (state.flushTimer) return;
                if (state.renderedCount < LIVE_TABLE_LARGE_ROWS) { flush(); return; }
                const wait = Math.max(0, state.lastFlushAt + LIVE_TABLE_LARGE_INTERVAL_MS - performance.now());
                state.flushTimer = setTimeout(flush, wait);
            }
            return {
                state,
                tabVisible,
                renderFull,
                // Replace all rows (full snapshot from the server).
                setAll(rows) {
                    state.lastSeq = (rows || []).reduce((max, r) => Math.max(max, seqOf(r)), 0);
                    if (tabVisible()) renderFull();
                    else { clearPending(); state.needsRender = true; }
                },
                // Apply a delta: new rows plus `min_seq` (rows that left the server cache).
                applyDelta(payload) {
                    const minSeq = Number(payload?.min_seq) || 0;
                    const rows = opts.getRows();
                    const fresh = (payload?.rows || []).filter(r => seqOf(r) > state.lastSeq);
                    if (fresh.length) state.lastSeq = Math.max(state.lastSeq, ...fresh.map(seqOf));
                    const knownKeys = new Set(rows.map(logRecordKey));
                    const added = fresh.filter(r => !knownKeys.has(logRecordKey(r)));
                    const evicted = minSeq ? rows.filter(r => seqOf(r) < minSeq) : [];
                    let next = evicted.length ? rows.filter(r => seqOf(r) >= minSeq) : rows;
                    for (const r of added) next.push(r);
                    opts.setRows(next);
                    if (!added.length && !evicted.length) return false;
                    if (!tabVisible() || state.needsRender) { state.needsRender = true; clearPending(); return true; }
                    const visible = opts.makeFilter ? opts.makeFilter() : opts.visible;
                    for (const r of added) if (visible(r)) state.pendingAppend.push(r);
                    state.pendingRemove += evicted.filter(visible).length;
                    scheduleFlush();
                    return true;
                },
                onTabShown() {
                    if (state.needsRender) renderFull();
                },
            };
        }

        socket.on('update_specific_event_table', (d) => {
             lastSpecificEventData = dedupeLogRecords(d);
             specificLive.setAll(lastSpecificEventData);
        });

        // One row of the Specific Validator table.  Full renders and appended
        // rows use this same template, so both paths produce identical HTML.
        function specificRowHtml(res) {
             return `<tr class="hover:bg-gray-50 border-b text-sm"><td class="py-2 px-3 text-purple-700 text-sm">${res.device_name}</td><td class="py-2 px-3 text-sm font-semibold ${res.status === 'PASSED'?'text-green-600':'text-red-600'}">${res.status}</td><td class="py-2 px-3"><span class="event-name-link cursor-pointer text-sm font-medium text-indigo-700 hover:underline" data-event-name="${escapeAttribute(res.event_name)}">${res.event_name}</span></td><td class="py-2 px-3 details-cell text-sm">${res.details}</td><td class="py-2 px-3 log-cell text-xs font-normal text-gray-600">${escapeHTML(res.raw_log || '')}</td><td class="py-2 px-3"><button class="view-json-btn text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 font-medium py-1 px-2 rounded" data-json='${escapeAttribute(res.json_data)}'>View JSON</button></td></tr>`;
        }

        function specificRowVisible(r) {
             const sourceFilter = document.querySelector('input[name="specificSourceFilter"]:checked')?.value || 'all';
             const textFilter = (document.getElementById('specificParamInput')?.value || '').trim().toLowerCase();
             if (selectedDevice !== 'all' && r.device_id !== selectedDevice) return false;
             if (sourceFilter !== 'all' && (r.source || 'firebase') !== sourceFilter) return false;
             if (textFilter && !(r.raw_log || '').toLowerCase().includes(textFilter)) return false;
             return true;
        }

        var specificLive = createLiveTable({
             tbodyId: 'specificEventTableBody',
             tabContentId: 'tabContentSpecific',
             emptyHtml: '<tr><td colspan="6" class="text-center py-4">Waiting...</td></tr>',
             getRows: () => lastSpecificEventData,
             setRows: rows => { lastSpecificEventData = rows; },
             rowHtml: specificRowHtml,
             visible: specificRowVisible,
        });

        function renderSpecificEventTable() {
             specificLive.renderFull();
        }

        socket.on('specific_event_delta', (payload) => specificLive.applyDelta(payload));

        let lastAdRevenueData = [];

         socket.on('update_adrevenue_table', (d) => {
             lastAdRevenueData = dedupeLogRecords(d);
             adRevenueLive.setAll(lastAdRevenueData);
        });
        socket.on('adrevenue_delta', (payload) => adRevenueLive.applyDelta(payload));

        function adRevenueRowHtml(res) {
            return `<tr class="hover:bg-gray-50 border-b text-sm"><td class="py-2 px-3 text-purple-700 text-sm align-top whitespace-nowrap">${res.device_name}</td><td class="py-2 px-3 text-sm font-semibold ${res.status === 'PASSED'?'text-green-600':(res.status === 'FAILED'?'text-red-600':'text-orange-500')} align-top whitespace-nowrap">${res.status}</td><td class="py-2 px-3 align-top"><span class="event-name-link cursor-pointer text-sm font-medium text-indigo-700 hover:underline" data-event-name="${escapeAttribute(res.event_name)}">${res.event_name}</span></td><td class="py-2 px-3 details-cell text-sm align-top"><div class="adrevenue-panel adrevenue-details-panel">${res.details}</div></td><td class="py-2 px-3 log-cell text-xs font-normal text-gray-600 align-top"><div class="adrevenue-panel adrevenue-raw-panel">${escapeHTML(res.raw_log || '')}</div></td><td class="py-2 px-3 align-top whitespace-nowrap"><button class="view-json-btn text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 font-medium py-1 px-2 rounded" data-json='${escapeAttribute(res.json_data)}'>View JSON</button></td></tr>`;
        }

        var adRevenueLive = createLiveTable({
            tbodyId: 'adRevenueTableBody',
            tabContentId: 'tabContentAdRevenue',
            emptyHtml: '<tr><td colspan="6" class="text-center py-4">Waiting...</td></tr>',
            getRows: () => lastAdRevenueData,
            setRows: rows => { lastAdRevenueData = rows; },
            rowHtml: adRevenueRowHtml,
            makeFilter: () => {
            const filterText = document.getElementById('adRevenueFilterInput').value.toLowerCase();
            const sourceFilter = document.querySelector('input[name="adRevenueSourceFilter"]:checked')?.value || 'all';
                return r => {
                if (selectedDevice !== 'all' && r.device_id !== selectedDevice) return false;
                if (sourceFilter !== 'all' && (r.source || '').toLowerCase() !== sourceFilter) return false;
                if (filterText && !(r.raw_log || '').toLowerCase().includes(filterText)) return false;
                return true;
                };
            },
        });

        function renderAdRevenueTable() {
            adRevenueLive.renderFull();
        }

        // ============================================
        // === FIXED: Callback Table with Client Filter ===
        // ============================================
        
        let lastCallbackData = [];

        socket.on('callback_ad_delta', (payload) => callbackLive.applyDelta(payload));
        socket.on('update_callback_ad_table', (d) => {
            lastCallbackData = dedupeLogRecords(d);
            callbackLive.setAll(lastCallbackData);
        });

        function callbackRowHtml(res) {
                     const nameLower = (res.event_name || '').toLowerCase();
                     const isFailed = nameLower.includes('failed');
                     const isImpression = nameLower.includes('onimpression') || nameLower.includes('_onimpression');
                     const eventClass = isFailed ? 'text-red-600' : (isImpression ? 'text-blue-600' : '');
                     return `<tr class="hover:bg-gray-50 border-b text-sm"><td class="py-2 px-3 text-purple-700 text-sm align-top">${res.device_name}</td><td class="py-2 px-3 text-sm font-semibold align-top ${res.type==='Ad Event'?'text-orange-600':'text-cyan-600'}">${res.type}</td><td class="py-2 px-3 text-sm font-medium align-top ${eventClass}">${res.event_name}</td><td class="py-2 px-3 details-cell text-sm align-top"><div class="max-h-64 overflow-auto">${res.details}</div></td><td class="py-2 px-3 log-cell text-xs font-normal text-gray-600 align-top"><div class="max-h-64 overflow-auto whitespace-pre-wrap break-all">${escapeHTML(res.raw_log || '')}</div></td><td class="py-2 px-3 align-top"><button class="view-json-btn text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 font-medium py-1 px-2 rounded" data-json='${escapeAttribute(res.json_data)}'>View JSON</button></td></tr>`;
        }

        var callbackLive = createLiveTable({
            tbodyId: 'callbackAdTableBody',
            tabContentId: 'tabContentCallbackAd',
            emptyHtml: '<tr><td colspan="6" class="text-center py-4">Waiting...</td></tr>',
            getRows: () => lastCallbackData,
            setRows: rows => { lastCallbackData = rows; },
            rowHtml: callbackRowHtml,
            makeFilter: () => {
            // Get filter values. Multiple callback groups can be selected together.
            const selectedTypeFilters = Array.from(document.querySelectorAll('input[name="callbackType"]:checked'))
                .map(input => input.value)
                .filter(value => value !== 'all');
            const textFilter = document.getElementById('callbackAdFilterInput').value.toLowerCase();

                return r => {
                // Device filter
                if (selectedDevice !== 'all' && r.device_id !== selectedDevice) return false;
                
                // Type filter
                const callbackTypeMap = {
                    callback: 'Callback Levelplay',
                    gadsme_callback: 'Callback Gadsme',
                    adverty5_callback: 'Callback Adverty5',
                    ascendx_callback: 'Callback Ascendx',
                    cloudx_callback: 'Callback Cloudx',
                    ad_event: 'Ad Event',
                };
                if (selectedTypeFilters.length > 0 && !selectedTypeFilters.some(value => callbackTypeMap[value] === r.type)) return false;
                
                // Text filter
                if (textFilter && !r.raw_log.toLowerCase().includes(textFilter)) return false;

                return true;
                };
            },
        });

        function renderCallbackTable() {
            callbackLive.renderFull();
        }

        // Add listeners
        function bindCallbackCheckboxGroup() {
            document.querySelectorAll('input[name="callbackType"]').forEach(input => input.addEventListener('change', (event) => {
                const allInput = document.getElementById('callbackTypeAll');
                const specificInputs = Array.from(document.querySelectorAll('input[name="callbackType"]'))
                    .filter(item => item !== allInput);
                if (event.target === allInput) {
                    if (allInput.checked) specificInputs.forEach(item => item.checked = false);
                } else if (event.target.checked) {
                    allInput.checked = false;
                } else if (!specificInputs.some(item => item.checked)) {
                    allInput.checked = true;
                }
                renderCallbackTable();
            }));
        }
        bindCallbackCheckboxGroup();
        document.getElementById('callbackAdFilterInput').addEventListener('input', renderCallbackTable);

        let lastPriceRotationData = [];

        function resetPriceRotationUiState() {
            lastPriceRotationData = [];
            const filterInput = document.getElementById('priceRotationFilterInput');
            const allTypeFilter = document.getElementById('priceRotationTypeAll');
            const allAdTypeFilter = document.getElementById('priceRotationAdTypeAll');
            const tbody = document.getElementById('priceRotationTableBody');
            if (filterInput) filterInput.value = '';
            document.querySelectorAll('input[name="priceRotationTypeFilter"]').forEach(input => {
                input.checked = input === allTypeFilter;
            });
            document.querySelectorAll('input[name="priceRotationAdTypeFilter"]').forEach(input => {
                input.checked = input === allAdTypeFilter;
            });
            // Same placeholder as before, rendered through the live table so it
            // knows the table is empty again.
            if (tbody && typeof priceRotationLive !== 'undefined' && priceRotationLive) priceRotationLive.renderFull();
        }

        socket.on('price_rotation_delta', (payload) => priceRotationLive.applyDelta(payload));
        socket.on('update_price_rotation_table', (data) => {
            lastPriceRotationData = dedupeLogRecords(data);
            priceRotationLive.setAll(lastPriceRotationData);
        });

        function priceRotationRowHtml(res) {
                const jsonData = res.json_data && res.json_data !== '{}' ? res.json_data : '';
                const action = jsonData
                    ? `<button class="view-json-btn text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 font-medium py-1 px-2 rounded" data-json='${escapeAttribute(jsonData)}'>View JSON</button>`
                    : '<span class="text-xs text-gray-400">-</span>';
                return `<tr class="hover:bg-gray-50 border-b text-sm">
                    <td class="py-2 px-3 text-purple-700 text-sm align-top truncate" title="${escapeAttribute(res.device_name || res.device_id || '')}">${escapeHTML(res.device_name || res.device_id || '')}</td>
                    <td class="py-2 px-3 text-sm font-semibold text-cyan-600 align-top break-words">${escapeHTML(res.type || 'Unknown')}</td>
                    <td class="py-2 px-3 align-top price-rotation-details"><pre class="text-xs font-mono whitespace-pre-wrap break-all max-h-64 overflow-auto">${escapeHTML(res.details || '')}</pre></td>
                    <td class="py-2 px-3 log-cell text-xs font-normal text-gray-600 align-top"><div class="max-h-64 overflow-auto whitespace-pre-wrap break-all">${escapeHTML(res.raw_log || '')}</div></td>
                    <td class="py-2 px-3 align-top whitespace-nowrap">${action}</td>
                </tr>`;
        }

        var priceRotationLive = createLiveTable({
            tbodyId: 'priceRotationTableBody',
            tabContentId: 'tabContentPriceRotation',
            emptyHtml: '<tr><td colspan="5" class="text-center py-4">Waiting...</td></tr>',
            getRows: () => lastPriceRotationData,
            setRows: rows => { lastPriceRotationData = rows; },
            rowHtml: priceRotationRowHtml,
            makeFilter: () => {
            const filterInput = document.getElementById('priceRotationFilterInput');
            const selectedTypeFilters = Array.from(document.querySelectorAll('input[name="priceRotationTypeFilter"]:checked'))
                .map(input => input.value)
                .filter(value => value !== 'all');
            const selectedAdTypeFilters = Array.from(document.querySelectorAll('input[name="priceRotationAdTypeFilter"]:checked'))
                .map(input => input.value)
                .filter(value => value !== 'all');
            const textFilter = (filterInput?.value || '').toLowerCase();
            const priceRotationTypeAliases = {
                max: new Set(['max', 'applovin', 'applovinmax'])
            };
            const normalizePriceRotationType = value => String(value || '')
                .trim()
                .toLowerCase()
                .replace(/[^a-z0-9]+/g, '');
                return res => {
                if (selectedDevice !== 'all' && res.device_id !== selectedDevice) return false;
                const rotationType = normalizePriceRotationType(res.type);
                if (rotationType === 'ascendx' || rotationType === 'cloudx') return false;
                if (selectedTypeFilters.length > 0 && !selectedTypeFilters.some(value => {
                    const aliases = priceRotationTypeAliases[value] || new Set([value]);
                    return aliases.has(rotationType);
                })) return false;
                const rawLog = (res.raw_log || '').toLowerCase();
                if (selectedAdTypeFilters.length > 0 && !selectedAdTypeFilters.some(value => rawLog.includes(value))) return false;
                if (textFilter && !rawLog.includes(textFilter)) return false;
                return true;
                };
            },
        });

        function renderPriceRotationTable() {
            priceRotationLive.renderFull();
        }

        document.getElementById('priceRotationFilterInput')?.addEventListener('input', renderPriceRotationTable);
        function bindPriceRotationCheckboxGroup(inputName, allInputId) {
            document.querySelectorAll(`input[name="${inputName}"]`).forEach(input => input.addEventListener('change', (event) => {
                const allInput = document.getElementById(allInputId);
                const specificInputs = Array.from(document.querySelectorAll(`input[name="${inputName}"]`))
                    .filter(item => item !== allInput);
                if (event.target === allInput) {
                    if (allInput.checked) specificInputs.forEach(item => item.checked = false);
                } else if (event.target.checked) {
                    allInput.checked = false;
                } else if (!specificInputs.some(item => item.checked)) {
                    allInput.checked = true;
                }
                renderPriceRotationTable();
            }));
        }

        bindPriceRotationCheckboxGroup('priceRotationTypeFilter', 'priceRotationTypeAll');
        bindPriceRotationCheckboxGroup('priceRotationAdTypeFilter', 'priceRotationAdTypeAll');

        socket.on('update_sdk_check_table', (data) => {
            const tbody = document.getElementById(activePlatform === 'ios' ? 'sdkCheckIosTableBody' : 'sdkCheckTableBody');
            if (!tbody) return;
            tbody.innerHTML = data.map(res => {
                 let rowClass = (selectedDevice !== 'all' && res.status !== 'HEADER') ? 'pl-8' : '';
                 let statusText = '';
                 if (res.status === 'PASSED') statusText = '<span class="font-semibold text-green-600"> - PASSED</span>';
                 else if (res.status === 'FAILED') statusText = '<span class="font-semibold text-red-600"> - FAILED</span>';
                 else if (res.status === 'NOT_FOUND') statusText = '<span class="font-semibold text-red-600"> - Not Found</span>';
                 else if (res.status === 'FOUND') statusText = '<span class="font-semibold text-amber-600"> - FOUND</span>';
                 else if (res.status === 'HEADER') rowClass += ' font-semibold text-indigo-600 bg-gray-50';
                 else if (res.status === 'LABEL') rowClass += ' font-bold text-indigo-700 pt-3';
                 else if (res.status === 'SECTION') rowClass += ' font-bold text-slate-500 pt-4 border-t border-gray-200';
                 else if (res.status === 'WAITING') rowClass += ' text-gray-500 italic';
                 // Same text size as the other result tables (and it follows the zoom):
                 // an unsized <pre> fell back to the browser's larger monospace size.
                 const sizeClass = res.status === 'LABEL' ? 'text-sm' : 'text-xs';
                 
                 // Filter
                 if (selectedDevice !== 'all' && res.device_id !== selectedDevice && res.status !== 'LABEL' && res.status !== 'SECTION' && res.status !== 'WAITING') return '';
                 
                 return `<tr><td class="py-1 px-4 ${rowClass}"><pre class="font-mono ${sizeClass} leading-5" style="margin: 0; white-space: pre-wrap;">${escapeHTML(res.display_text)}${statusText}</pre></td></tr>`;
            }).join('');
        });

        socket.on('runtime_reset', () => {
            resetRuntimeUiForPlatformSwitch();
        });

        let lastPackageLogs = [];
        let selectedPackageRowKeys = new Set();
        let lastPackageFilterSignature = '';
        let lastPackageRenderedCount = 0;
        let lastPackageFirstRowKey = '';
        let lastPackageLastRowKey = '';
        let packageHistorySessions = [];
        let packageUiPaused = false;
        let pausedPackageSnapshot = [];

        function getPackageRowKey(l) {
            // Live rows are unique by seq; history rows keep the record key.
            if (l && l.seq) return 'seq:' + l.seq;
            return logRecordKey(l);
        }

        function getPackageFilterState() {
            const selectedQuickFilter = document.querySelector('input[name="tagQuickFilter"]:checked');
            return {
                selectedDevice,
                filterText: document.getElementById('packageFilterInput').value.toLowerCase(),
                filterText2: document.getElementById('packageFilterInput2').value.toLowerCase(),
                filterText3: document.getElementById('packageFilterInput3').value.toLowerCase(),
                tagFilter: document.getElementById('packageTagFilterInput').value.toLowerCase(),
                quick: resolveQuickFilter(selectedQuickFilter?.value || ''),
                errorsOnly: document.getElementById('showErrorsOnly').checked,
                warningsOnly: document.getElementById('showWarningsOnly').checked,
            };
        }

        function filterPackageLogs(logs, state) {
            return logs.filter(l => {
                if (state.selectedDevice !== 'all' && l.device_id !== state.selectedDevice) return false;
                const isWarningLevel = l.level === 'W';
                if (state.errorsOnly && state.warningsOnly) {
                    if (!(l.is_error || isWarningLevel)) return false;
                } else if (state.errorsOnly) {
                    if (!l.is_error) return false;
                } else if (state.warningsOnly) {
                    if (!isWarningLevel) return false;
                }
                // Message filters search the whole log line (time, level, tag,
                // message, device), like the saved-session search does.
                const logText = `${l.time_display || l.time || ''} ${l.level || ''} ${l.tag || ''}: ${l.message || ''} ${l.log || ''} ${l.device_name || ''}`;
                const logHaystack = logText.toLowerCase();
                const tagHaystack = `${l.tag || ''}`.toLowerCase();
                if (state.quick) {
                    const quick = state.quick;
                    if (quick.tag && !tagHaystack.includes(quick.tag)) return false;
                    if (quick.terms.some(term => !logHaystack.includes(term))) return false;
                }
                if (state.tagFilter && !tagHaystack.includes(state.tagFilter)) return false;
                if (state.filterText && !logHaystack.includes(state.filterText)) return false;
                if (state.filterText2 && !logHaystack.includes(state.filterText2)) return false;
                if (state.filterText3 && !logHaystack.includes(state.filterText3)) return false;
                return true;
            });
        }

        function packageRowHtml(l, idx) {
            const msgText = (l.message || l.log || '');
            const isErrorLevel = (l.level === 'E' || l.level === 'F');
            const isWarningLevel = l.level === 'W';
            const rowClass = isErrorLevel ? 'text-red-500' : (isWarningLevel ? 'text-amber-500' : '');
            const msgClass = isErrorLevel ? 'text-red-500' : (isWarningLevel ? 'text-amber-500' : '');
            const rowKey = getPackageRowKey(l);
            const selectedClass = selectedPackageRowKeys.has(rowKey) ? 'selected' : '';
            return `<tr class="package-log-row hover:bg-gray-50 ${rowClass} ${selectedClass}" data-row-key="${encodeURIComponent(rowKey)}" data-row-index="${idx}"><td class="py-1.5 px-2 font-mono text-[11px] leading-4 time-cell col-time">${escapeHTML(l.time_display || l.time || '')}</td><td class="py-1.5 pr-1 pl-2 font-mono text-[11px] leading-4 tag-cell col-tag" title="${escapeHTML(l.tag || '')}">${escapeHTML(l.tag || '')}</td><td class="py-1.5 pl-1 pr-3 font-mono text-[11px] leading-4 log-cell message-cell col-message ${msgClass}">${escapeHTML(msgText)}</td></tr>`;
        }

        // --- Live Package Log view (virtualized) ---
        // lastPackageLogs is the client store, oldest first, holding up to
        // PACKAGE_STORE_MAX_ROWS rows.  packageView holds the rows that pass
        // the current filters.  Only the rows inside the viewport (plus a
        // small overscan) exist in the DOM; spacer rows keep the scrollbar
        // the size of the whole view.
        const PACKAGE_STORE_MAX_ROWS = 200000;
        const PACKAGE_STORE_TRIM_SLACK = 20000;
        const PACKAGE_VIEW_OVERSCAN = 40;
        let packageView = [];
        let packageViewSignature = '';
        let packageRowHeight = 18;
        let packageStickToBottom = true;
        let packageNewRowsWhileAway = 0;
        let packagePausedAtSeq = 0;
        let packagePausedNewRows = 0;
        let packageViewportFirst = -1;
        let packageViewportLast = -1;
        let packageViewportDirty = true;
        let packageViewportFramePending = false;
        // Rows older than the live store, loaded from SQLite while the user
        // scrolls up.  While they are shown the store is not trimmed, so the
        // view never gets a hole between history and live rows.
        const PACKAGE_STORE_HARD_MAX_ROWS = 400000;
        const PACKAGE_OLDER_PAGE_ROWS = 2000;
        let packageHistoryPrefix = [];
        let packageOlderLoading = false;
        let packageOlderExhausted = false;
        let packageSessionId = null;

        function rebuildPackageView(state, signature) {
            const live = packageUiPaused
                ? lastPackageLogs.filter(row => row.seq <= packagePausedAtSeq)
                : lastPackageLogs;
            const source = packageHistoryPrefix.length ? packageHistoryPrefix.concat(live) : live;
            packageView = filterPackageLogs(source, state);
            packageViewSignature = signature;
            packageViewportDirty = true;
        }

        function appendToPackageStore(rows) {
            for (const row of rows) lastPackageLogs.push(row);
            let keptFromSeq = 0;
            let prefixDropped = false;
            if (packageHistoryPrefix.length && lastPackageLogs.length > PACKAGE_STORE_HARD_MAX_ROWS) {
                // Memory guard while browsing history for a long time.
                packageHistoryPrefix = [];
                packageOlderExhausted = false;
                prefixDropped = true;
            }
            if (!packageHistoryPrefix.length && lastPackageLogs.length > PACKAGE_STORE_MAX_ROWS + PACKAGE_STORE_TRIM_SLACK) {
                lastPackageLogs.splice(0, lastPackageLogs.length - PACKAGE_STORE_MAX_ROWS);
                keptFromSeq = lastPackageLogs[0].seq;
            }
            if (packageUiPaused) {
                packagePausedNewRows += rows.length;
                updatePackagePauseLabel();
                return;
            }
            const state = getPackageFilterState();
            const signature = JSON.stringify(state);
            if (prefixDropped || signature !== packageViewSignature) {
                rebuildPackageView(state, signature);
                return;
            }
            const matched = filterPackageLogs(rows, state);
            for (const row of matched) packageView.push(row);
            if (keptFromSeq) {
                let dropped = 0;
                while (dropped < packageView.length && packageView[dropped].seq < keptFromSeq) dropped++;
                if (dropped) {
                    packageView.splice(0, dropped);
                    if (!packageStickToBottom) {
                        // Keep the rows the user is reading in place.
                        const container = document.getElementById('packageLogContainer');
                        if (container) container.scrollTop = Math.max(0, container.scrollTop - dropped * packageRowHeight);
                    }
                }
            }
            if (matched.length) {
                packageViewportDirty = true;
                if (!packageStickToBottom) {
                    packageNewRowsWhileAway += matched.length;
                    updatePackageNewRowsButton();
                }
            }
        }

        function packageAutoScrollEnabled() {
            const autoScroll = document.getElementById('autoScroll');
            return !autoScroll || autoScroll.checked;
        }

        function updatePackageNewRowsButton() {
            const btn = document.getElementById('packageNewRowsBtn');
            if (!btn) return;
            const show = !packageStickToBottom && packageNewRowsWhileAway > 0;
            btn.classList.toggle('hidden', !show);
            if (show) btn.textContent = `${packageNewRowsWhileAway.toLocaleString()} dòng mới ↓`;
        }

        function renderPackageViewport(force = false) {
            const tbody = document.getElementById('packageLogTableBody');
            const container = document.getElementById('packageLogContainer');
            if (!tbody || !container) return;
            const head = container.querySelector('thead');
            const headHeight = head ? head.offsetHeight : 0;
            const viewHeight = container.clientHeight || 600;
            const rowHeight = packageRowHeight;
            const total = packageView.length;
            const pinned = packageStickToBottom && packageAutoScrollEnabled() && !packageUiPaused;
            const scrollTop = pinned
                ? Math.max(0, headHeight + total * rowHeight - viewHeight)
                : container.scrollTop;
            const bodyTop = Math.max(0, scrollTop - headHeight);
            const first = Math.max(0, Math.min(total, Math.floor(bodyTop / rowHeight) - PACKAGE_VIEW_OVERSCAN));
            const last = Math.min(total, Math.ceil((bodyTop + viewHeight) / rowHeight) + PACKAGE_VIEW_OVERSCAN);
            if (!force && !packageViewportDirty && first === packageViewportFirst && last === packageViewportLast) return;
            const html = [];
            if (first > 0) html.push(`<tr class="package-log-spacer"><td colspan="3" style="height:${first * rowHeight}px"></td></tr>`);
            for (let i = first; i < last; i++) html.push(packageRowHtml(packageView[i], i));
            const bottom = (total - last) * rowHeight;
            if (bottom > 0) html.push(`<tr class="package-log-spacer"><td colspan="3" style="height:${bottom}px"></td></tr>`);
            tbody.innerHTML = html.join('');
            packageViewportFirst = first;
            packageViewportLast = last;
            packageViewportDirty = false;
            if (pinned) container.scrollTop = container.scrollHeight;
            const sample = tbody.querySelector('tr.package-log-row');
            const measured = sample ? sample.getBoundingClientRect().height : 0;
            if (measured > 0 && Math.abs(measured - packageRowHeight) > 0.5) {
                // Spacer math must use the real row height or the scrollbar drifts.
                packageRowHeight = measured;
                packageViewportDirty = true;
                schedulePackageViewportRender();
            }
        }

        function schedulePackageViewportRender() {
            if (packageViewportFramePending) return;
            packageViewportFramePending = true;
            setTimeout(() => {
                packageViewportFramePending = false;
                renderPackageViewport();
            }, 16);
        }

        // Status badge for loading older rows.  Informational messages hide
        // themselves so the badge never sits over the first log row.
        var packageOlderStatusTimer = null;
        function setPackageOlderStatus(text, autoHideMs = 0) {
            const el = document.getElementById('packageOlderStatus');
            if (!el) return;
            clearTimeout(packageOlderStatusTimer);
            el.textContent = text || '';
            el.classList.toggle('hidden', !text);
            if (text && autoHideMs > 0) packageOlderStatusTimer = setTimeout(() => setPackageOlderStatus(''), autoHideMs);
        }

        function releasePackageHistoryPrefix() {
            // Back at the newest rows: drop history rows so the store can be
            // trimmed again.
            if (!packageHistoryPrefix.length) return;
            packageHistoryPrefix = [];
            packageOlderExhausted = false;
            packageSelectionAnchor = null;
            setPackageOlderStatus('');
            renderPackageLogTable(true);
        }

        function loadOlderPackageRows() {
            if (packageOlderLoading || packageOlderExhausted) return;
            const oldest = packageHistoryPrefix[0] || lastPackageLogs[0];
            if (!oldest || !oldest.seq || oldest.seq <= 1) return;
            const requestStream = packageStreamId;
            const beforeSeq = oldest.seq;
            const params = new URLSearchParams({ before_seq: String(beforeSeq), limit: String(PACKAGE_OLDER_PAGE_ROWS) });
            if (packageSessionId) params.set('session_id', String(packageSessionId));
            packageOlderLoading = true;
            setPackageOlderStatus('Đang tải dòng cũ từ lịch sử…');
            fetch(`/api/package-log/older?${params.toString()}`)
                .then(r => r.json())
                .then(data => {
                    if (requestStream !== packageStreamId) return;
                    if (!data.ok) throw new Error(data.error || 'failed_to_load_older_rows');
                    const rows = (data.rows || []).filter(row => row.seq < beforeSeq);
                    if (!data.has_more) packageOlderExhausted = true;
                    if (!rows.length) {
                        // Nothing older: only say so when older rows were
                        // already loaded; otherwise the first row is in view.
                        setPackageOlderStatus(packageOlderExhausted && packageHistoryPrefix.length ? 'Đã tới đầu session' : '', 2500);
                        return;
                    }
                    packageHistoryPrefix = rows.concat(packageHistoryPrefix);
                    const state = getPackageFilterState();
                    const signature = JSON.stringify(state);
                    const matched = filterPackageLogs(rows, state);
                    if (signature === packageViewSignature) {
                        packageView = matched.concat(packageView);
                        if (packageSelectionAnchor !== null) packageSelectionAnchor += matched.length;
                        const container = document.getElementById('packageLogContainer');
                        // Keep the row under the cursor in place.
                        if (container) container.scrollTop += matched.length * packageRowHeight;
                        packageViewportDirty = true;
                        renderPackageViewport(true);
                    } else {
                        renderPackageLogTable(true);
                    }
                    setPackageOlderStatus(`+${packageHistoryPrefix.length.toLocaleString()} dòng cũ từ lịch sử` + (packageOlderExhausted ? ' (đầu session)' : ''), 3000);
                })
                .catch(err => setPackageOlderStatus(`Không tải được dòng cũ: ${err}`, 6000))
                .finally(() => { packageOlderLoading = false; });
        }

        function renderPackageLogTable(forceFull = false) {
            const state = getPackageFilterState();
            const signature = JSON.stringify(state);
            const filterChanged = signature !== packageViewSignature;
            if (forceFull || filterChanged) rebuildPackageView(state, signature);
            if (filterChanged && selectedPackageRowKeys.size && revealSelectedPackageRow()) return;
            renderPackageViewport(forceFull);
        }

        // After a filter change (Clear filter, All, typing), keep the selected
        // row and scroll the new view so it sits in the middle.
        function revealSelectedPackageRow() {
            const container = document.getElementById('packageLogContainer');
            if (!container) return false;
            const idx = packageView.findIndex(row => selectedPackageRowKeys.has(getPackageRowKey(row)));
            if (idx < 0) return false;
            packageSelectionAnchor = idx;
            packageStickToBottom = false;
            renderPackageViewport(true);
            const head = container.querySelector('thead');
            const headHeight = head ? head.offsetHeight : 0;
            container.scrollTop = Math.max(0, headHeight + idx * packageRowHeight - (container.clientHeight - headHeight) / 2);
            packageLastScrollTop = container.scrollTop;
            packageViewportDirty = true;
            renderPackageViewport(true);
            updatePackageNewRowsButton();
            return true;
        }

        var packageLastScrollTop = 0;
        document.getElementById('packageLogContainer')?.addEventListener('scroll', () => {
            const container = document.getElementById('packageLogContainer');
            const scrolledUp = container.scrollTop < packageLastScrollTop;
            packageLastScrollTop = container.scrollTop;
            const atBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - packageRowHeight * 2;
            packageStickToBottom = atBottom;
            if (atBottom) packageNewRowsWhileAway = 0;
            updatePackageNewRowsButton();
            if (atBottom && packageAutoScrollEnabled()) releasePackageHistoryPrefix();
            // Fetch older history only while the user scrolls up near the top.
            if (scrolledUp && container.scrollTop < packageRowHeight * 30) loadOlderPackageRows();
            schedulePackageViewportRender();
        }, { passive: true });

        document.getElementById('packageNewRowsBtn')?.addEventListener('click', () => {
            packageStickToBottom = true;
            packageNewRowsWhileAway = 0;
            updatePackageNewRowsButton();
            if (packageHistoryPrefix.length) releasePackageHistoryPrefix();
            renderPackageViewport(true);
        });

        document.getElementById('autoScroll')?.addEventListener('change', (e) => {
            if (e.target.checked) {
                packageStickToBottom = true;
                packageNewRowsWhileAway = 0;
                updatePackageNewRowsButton();
                releasePackageHistoryPrefix();
            }
            renderPackageViewport(true);
        });

        window.addEventListener('resize', () => schedulePackageViewportRender());

        socket.on('package_log_db_status', (status) => {
            const banner = document.getElementById('packageDbWarning');
            if (!banner) return;
            const dropped = Number(status?.dropped_rows || 0);
            const error = String(status?.error || '');
            const parts = [];
            if (error) parts.push(`Không ghi được lịch sử Package Log (đang thử lại): ${error}`);
            if (dropped > 0) parts.push(`${dropped.toLocaleString()} dòng không lưu được vào lịch sử (vẫn còn trên màn hình nếu chưa bị cắt).`);
            banner.textContent = parts.join(' ');
            banner.classList.toggle('hidden', parts.length === 0);
        });

        function searchWholePackageSession() {
            // Live filters cover the in-memory rows; this runs the same
            // Filter 1-3 / Tag / quick filters on every saved row of the
            // session through the History view.
            const message1 = document.getElementById('packageFilterInput')?.value || '';
            const message2 = document.getElementById('packageFilterInput2')?.value || '';
            const message3 = document.getElementById('packageFilterInput3')?.value || '';
            const tag = document.getElementById('packageTagFilterInput')?.value || '';
            const quickId = document.querySelector('input[name="tagQuickFilter"]:checked')?.value || '';
            const modal = document.getElementById('packageHistoryModal');
            if (!modal) return;
            modal.style.zIndex = '60';
            modal.classList.remove('hidden');
            setPackageHistoryPlatform(activePlatform, false);
            fetch(`/api/package-log/sessions?limit=500&platform=${packageHistoryPlatform}`)
                .then(r => r.json())
                .then(data => {
                    if (!data.ok) throw new Error(data.error || 'failed_to_load_sessions');
                    renderPackageHistorySessions(data);
                    loadPackageHistoryMaintenance();
                    const select = document.getElementById('packageHistorySessionSelect');
                    if (select && packageSessionId && Array.from(select.options).some(o => o.value === String(packageSessionId))) {
                        select.value = String(packageSessionId);
                    }
                    document.getElementById('packageHistoryFilterInput').value = message1;
                    document.getElementById('packageHistoryFilterInput2').value = message2;
                    document.getElementById('packageHistoryFilterInput4').value = message3;
                    document.getElementById('packageHistoryFilterInput3').value = tag;
                    renderHistoryQuickFilterOptions();
                    document.getElementById('packageHistoryQuickFilter').value = quickId;
                    loadSelectedPackageHistory(false);
                })
                .catch(err => {
                    document.getElementById('packageHistoryMeta').textContent = `Failed to load sessions: ${err}`;
                });
        }
        document.getElementById('searchPackageSessionBtn')?.addEventListener('click', searchWholePackageSession);

        // --- Package Log history maintenance (size, retention, compact) ---
        let packageHistoryMaintenance = null;
        function formatPackageBytes(bytes) {
            const value = Number(bytes || 0);
            if (value >= 1024 * 1024 * 1024) return `${(value / 1024 / 1024 / 1024).toFixed(2)} GB`;
            return `${Math.round(value / 1024 / 1024)} MB`;
        }
        function renderPackageHistoryMaintenance(info) {
            packageHistoryMaintenance = info;
            const infoEl = document.getElementById('packageHistoryStorageInfo');
            const retentionBtn = document.getElementById('applyPackageRetentionBtn');
            const compactBtn = document.getElementById('compactPackageHistoryBtn');
            if (!infoEl || !info || !info.ok) return;
            const policy = info.policy || {};
            const cleanup = info.cleanup || {};
            const policyText = `giữ ${policy.days} ngày / ${formatPackageBytes(policy.max_bytes)}`;
            const autoText = policy.confirmed ? 'tự dọn mỗi giờ' : 'chưa bật tự dọn';
            const cleanupText = cleanup.sessions > 0
                ? `Có thể dọn ${cleanup.sessions} session (${Number(cleanup.rows).toLocaleString()} dòng, khoảng ${formatPackageBytes(cleanup.estimated_bytes)}) từ ${cleanup.oldest_label}.`
                : 'Không có session nào cần dọn.';
            infoEl.textContent = `Lịch sử: ${formatPackageBytes(info.usage.file_bytes)} · ${info.total_sessions} session · ${Number(info.total_rows).toLocaleString()} dòng · Chính sách: ${policyText} (${autoText}). ${cleanupText}`;
            if (retentionBtn) {
                retentionBtn.classList.toggle('hidden', cleanup.sessions === 0 && policy.confirmed);
                retentionBtn.textContent = cleanup.sessions > 0 ? `Dọn ${cleanup.sessions} session cũ` : 'Bật tự dọn';
            }
            if (compactBtn) {
                compactBtn.disabled = !!info.capturing;
                compactBtn.title = info.capturing
                    ? 'Dừng Package Log trước khi Compact'
                    : 'Thu nhỏ file lịch sử và tạo index cho cuộn xem log cũ';
            }
        }
        function loadPackageHistoryMaintenance() {
            fetch('/api/package-log/maintenance')
                .then(r => r.json())
                .then(renderPackageHistoryMaintenance)
                .catch(err => {
                    const infoEl = document.getElementById('packageHistoryStorageInfo');
                    if (infoEl) infoEl.textContent = `Không đọc được dung lượng lịch sử: ${err}`;
                });
        }
        document.getElementById('applyPackageRetentionBtn')?.addEventListener('click', () => {
            const info = packageHistoryMaintenance || {};
            const cleanup = info.cleanup || {};
            const policy = info.policy || {};
            const lines = [
                cleanup.sessions > 0
                    ? `Xóa vĩnh viễn ${cleanup.sessions} session (${Number(cleanup.rows).toLocaleString()} dòng, khoảng ${formatPackageBytes(cleanup.estimated_bytes)})?`
                    : 'Bật tự dọn lịch sử Package Log?',
                `Chính sách: giữ ${policy.days} ngày, tối đa ${formatPackageBytes(policy.max_bytes)}. Session đang ghi không bao giờ bị xóa.`,
                'Sau khi đồng ý, app tự dọn theo chính sách mỗi giờ. Dữ liệu đã xóa không khôi phục được.',
            ];
            if (!confirm(lines.join(String.fromCharCode(10)))) return;
            const infoEl = document.getElementById('packageHistoryStorageInfo');
            if (infoEl) infoEl.textContent = 'Đang dọn lịch sử…';
            fetch('/api/package-log/retention/apply', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ confirm: true }),
            })
                .then(r => r.json())
                .then(result => {
                    if (!result.ok) throw new Error(result.error || 'retention_failed');
                    loadPackageHistorySessions();
                })
                .catch(err => { if (infoEl) infoEl.textContent = `Dọn lịch sử thất bại: ${err}`; });
        });
        document.getElementById('compactPackageHistoryBtn')?.addEventListener('click', () => {
            const message = [
                'Compact file lịch sử Package Log?',
                'Với file lớn có thể mất vài chục giây; app không ghi lịch sử trong lúc đó.',
            ].join(String.fromCharCode(10));
            if (!confirm(message)) return;
            const infoEl = document.getElementById('packageHistoryStorageInfo');
            if (infoEl) infoEl.textContent = 'Đang compact lịch sử…';
            fetch('/api/package-log/compact', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ confirm: true }),
            })
                .then(r => r.json())
                .then(result => {
                    if (!result.ok) {
                        throw new Error(result.error === 'capture_running' ? 'đang chạy Package Log, hãy Stop trước' : (result.error || 'compact_failed'));
                    }
                    if (infoEl) infoEl.textContent = `Compact xong: ${formatPackageBytes(result.before.file_bytes)} → ${formatPackageBytes(result.after.file_bytes)} (${result.seconds}s).`;
                    setTimeout(loadPackageHistoryMaintenance, 1500);
                })
                .catch(err => { if (infoEl) infoEl.textContent = `Compact thất bại: ${err}`; });
        });



        // Live Package Log rows arrive as ordered deltas.  Each row has a
        // server-assigned `seq`; `packageStreamId` changes on Start/Stop/Clear.
        let packageStreamId = null;
        let lastPackageSeq = 0;
        let packageSyncPending = false;
        let packageLiveRenderPending = false;
        function requestPackageLogSync() {
            packageSyncPending = true;
            socket.emit('package_log_sync', {
                stream: packageStreamId,
                since_seq: lastPackageSeq,
                // The server replay window; the client store keeps up to
                // PACKAGE_STORE_MAX_ROWS rows as live deltas arrive.
                limit: 50000,
            });
        }
        function resetPackageLogStore(streamId, sessionId) {
            packageStreamId = (streamId === undefined) ? null : streamId;
            if (sessionId !== undefined) packageSessionId = sessionId || null;
            packageHistoryPrefix = [];
            packageOlderExhausted = false;
            packageOlderLoading = false;
            setPackageOlderStatus('');
            lastPackageSeq = 0;
            lastPackageLogs = [];
            packageView = [];
            packageViewSignature = '';
            packageViewportDirty = true;
            packageStickToBottom = true;
            packageNewRowsWhileAway = 0;
            packagePausedNewRows = 0;
            selectedPackageRowKeys.clear();
            updatePackageNewRowsButton();
            updatePackagePauseLabel();
        }
        function scheduleLivePackageRender() {
            // Coalesce bursts of deltas into one render.  A timer (not
            // requestAnimationFrame) keeps rendering while the window is
            // hidden or minimized, as the previous snapshot renderer did.
            if (packageLiveRenderPending) return;
            packageLiveRenderPending = true;
            setTimeout(() => {
                packageLiveRenderPending = false;
                if (!packageUiPaused) renderPackageLogTable();
            }, 50);
        }
        socket.on('connect', () => requestPackageLogSync());
        socket.on('package_log_reset', (payload) => {
            resetPackageLogStore(payload ? payload.stream : null, payload ? payload.session_id : null);
            if (payload && 'package_id' in payload) {
                packageStreamTarget = payload.package_id || '';
                renderPackageStreamProjects();
            }
            renderPackageLogTable(true);
        });
        socket.on('package_log_delta', (payload) => {
            if (!payload || !Array.isArray(payload.rows)) return;
            if (payload.sync) {
                packageSyncPending = false;
            } else if (packageSyncPending) {
                // The socket is ordered: the pending sync reply already covers
                // every row emitted before it.
                return;
            }
            if (packageStreamId === null) packageStreamId = payload.stream;
            if (payload.stream !== packageStreamId) {
                requestPackageLogSync();
                return;
            }
            const fresh = payload.rows.filter(row => row.seq > lastPackageSeq);
            if (!fresh.length) return;
            if (!payload.sync && lastPackageSeq > 0 && fresh[0].seq > lastPackageSeq + 1) {
                // A delta was missed (for example during a reconnect).
                requestPackageLogSync();
                return;
            }
            lastPackageSeq = fresh[fresh.length - 1].seq;
            appendToPackageStore(fresh);
            scheduleLivePackageRender();
        });

        function renderPackageHistorySessions(payload) {
            const select = document.getElementById('packageHistorySessionSelect');
            const meta = document.getElementById('packageHistoryMeta');
            if (!select || !meta) return;
            const prev = select.value;
            packageHistorySessions = payload.sessions || [];
            select.innerHTML = '';
            if (packageHistorySessions.length === 0) {
                packageHistoryLoadedRows = [];
                packageHistoryOffset = 0;
                packageHistoryHasMore = false;
                updatePackageHistoryLoadMoreButton();
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = 'No saved sessions';
                select.appendChild(opt);
                meta.textContent = 'No saved package-log sessions yet.';
                document.getElementById('packageHistoryTableBody').innerHTML = '';
                return;
            }
            packageHistorySessions.forEach((s, idx) => {
                const opt = document.createElement('option');
                opt.value = String(s.id);
                opt.textContent = `#${s.id} - ${s.package_id} - ${s.started_label}`;
                if ((!prev && idx === 0) || prev === String(s.id)) opt.selected = true;
                select.appendChild(opt);
            });
            const current = packageHistorySessions.find(s => String(s.id) === select.value) || packageHistorySessions[0];
            if (current) {
                meta.textContent = `Selected: ${current.package_id} | Started: ${current.started_label} | Status: ${current.status} | Rows: ${current.row_count}`;
            }
        }

        function getPackageHistoryRowKey(row) {
            return [row.time_display || '', row.device_name || row.device_id || '', row.tag || '', row.raw_log || row.message || ''].join('||');
        }

        function renderPackageHistoryRows(rows, append = false) {
            const tbody = document.getElementById('packageHistoryTableBody');
            if (!tbody) return;
            const uniqueRows = dedupeLogRecords(rows);
            if ((!uniqueRows || uniqueRows.length === 0) && !append) {
                tbody.innerHTML = '<tr><td colspan="4" class="py-3 px-2 text-xs text-gray-500">No saved logs found for this filter.</td></tr>';
                return;
            }
            if (!uniqueRows || uniqueRows.length === 0) return;
            if (!append) tbody.innerHTML = '';
            const startIdx = append ? tbody.querySelectorAll('tr.package-history-row').length : 0;
            const html = uniqueRows.map((row, idx) => {
                const rowKey = getPackageHistoryRowKey(row);
                const selectedClass = selectedPackageHistoryRowKeys.has(rowKey) ? 'selected' : '';
                return `
                <tr class="package-history-row hover:bg-gray-50 ${selectedClass}" data-row-key="${encodeURIComponent(rowKey)}" data-row-index="${startIdx + idx}">
                    <td class="py-2 px-2 font-mono text-[11px] text-gray-700 align-top whitespace-nowrap">${escapeHTML(row.time_display || '')}</td>
                    <td class="py-2 px-2 text-[11px] text-gray-700 align-top whitespace-nowrap">${escapeHTML(row.device_name || row.device_id || '')}</td>
                    <td class="py-2 px-2 font-mono text-[11px] text-gray-700 align-top whitespace-nowrap max-w-[88px] w-[88px] overflow-hidden text-ellipsis" title="${escapeHTML(row.tag || '')}">${escapeHTML(row.tag || '')}</td>
                    <td class="py-2 px-2 font-mono text-[11px] text-gray-700 align-top whitespace-pre-wrap break-all">${escapeHTML(row.raw_log || row.message || '')}</td>
                </tr>
            `}).join('');
            if (append) tbody.insertAdjacentHTML('beforeend', html);
            else tbody.innerHTML = html;
        }

        function renderPackageHistoryPlatformSwitch() {
            document.querySelectorAll('#packageHistoryPlatformSwitch button').forEach(btn => {
                const on = btn.dataset.platform === packageHistoryPlatform;
                btn.classList.toggle('bg-indigo-600', on);
                btn.classList.toggle('text-white', on);
                btn.classList.toggle('bg-white', !on);
                btn.classList.toggle('text-slate-700', !on);
            });
        }
        function setPackageHistoryPlatform(platform, reload = true) {
            packageHistoryPlatform = platform === 'ios' ? 'ios' : 'android';
            renderPackageHistoryPlatformSwitch();
            renderHistoryQuickFilterOptions();
            if (!reload) return;
            const select = document.getElementById('packageHistorySessionSelect');
            if (select) select.value = '';
            packageHistoryLoadedRows = [];
            packageHistoryOffset = 0;
            packageHistoryHasMore = false;
            updatePackageHistoryLoadMoreButton();
            const tbody = document.getElementById('packageHistoryTableBody');
            if (tbody) tbody.innerHTML = '';
            loadPackageHistorySessions();
        }
        document.getElementById('packageHistoryPlatformSwitch')?.addEventListener('click', (e) => {
            const btn = e.target.closest('button[data-platform]');
            if (btn && btn.dataset.platform !== packageHistoryPlatform) setPackageHistoryPlatform(btn.dataset.platform);
        });

        function loadPackageHistorySessions() {
            fetch(`/api/package-log/sessions?limit=500&platform=${packageHistoryPlatform}`)
                .then(r => r.json())
                .then(data => {
                    if (!data.ok) throw new Error(data.error || 'failed_to_load_sessions');
                    renderPackageHistorySessions(data);
                    loadPackageHistoryMaintenance();
                })
                .catch(err => {
                    document.getElementById('packageHistoryMeta').textContent = `Failed to load sessions: ${err}`;
                });
        }

        function loadSelectedPackageHistory(append = false) {
            if (packageHistoryLoading) return;
            const select = document.getElementById('packageHistorySessionSelect');
            const keyword1 = document.getElementById('packageHistoryFilterInput')?.value || '';
            const keyword2 = document.getElementById('packageHistoryFilterInput2')?.value || '';
            const keyword3 = document.getElementById('packageHistoryFilterInput4')?.value || '';
            const tagKeyword = document.getElementById('packageHistoryFilterInput3')?.value || '';
            const quickId = document.getElementById('packageHistoryQuickFilter')?.value || '';
            if (!select || !select.value) {
                packageHistoryLoadedRows = [];
                packageHistoryOffset = 0;
                packageHistoryHasMore = false;
                updatePackageHistoryLoadMoreButton();
                renderPackageHistoryRows([]);
                return;
            }
            const offset = append ? packageHistoryOffset : 0;
            const query = new URLSearchParams({
                session_id: select.value,
                q1: keyword1,
                q2: keyword2,
                q3: tagKeyword,
                q4: keyword3,
                qf: quickId,
                qp: packageHistoryPlatform,
                offset: String(offset),
                limit: String(PACKAGE_HISTORY_PAGE_SIZE),
            });
            packageHistoryLoading = true;
            updatePackageHistoryLoadMoreButton();
            fetch(`/api/package-log/logs?${query.toString()}`)
                .then(r => r.json())
                .then(data => {
                    if (!data.ok) throw new Error(data.error || 'failed_to_load_logs');
                    const current = packageHistorySessions.find(s => String(s.id) === String(data.session_id));
                    const meta = document.getElementById('packageHistoryMeta');
                    if (!append) {
                        packageHistoryLoadedRows = [];
                        packageHistoryOffset = 0;
                        clearPackageHistorySelection();
                    }
                    packageHistoryLoadedRows = append ? packageHistoryLoadedRows.concat(data.rows || []) : (data.rows || []);
                    packageHistoryOffset = packageHistoryLoadedRows.length;
                    packageHistoryHasMore = !!data.has_more;
                    updatePackageHistoryLoadMoreButton();
                    if (meta && current) {
                        const filters = [
                            keyword1 ? `text1:${keyword1}` : '',
                            keyword2 ? `text2:${keyword2}` : '',
                            keyword3 ? `text3:${keyword3}` : '',
                            tagKeyword ? `tag:${tagKeyword}` : '',
                            quickId ? `quick:${(historyQuickFilters().find(f => f.id === quickId) || {}).name || quickId}` : '',
                        ].filter(Boolean).join(' | ');
                        meta.textContent = `Selected: ${current.package_id} | Started: ${current.started_label} | Status: ${current.status} | Showing: ${packageHistoryLoadedRows.length} / ${data.total_rows}${filters ? ` | Filter: ${filters}` : ''}`;
                    }
                    renderPackageHistoryRows(data.rows || [], append);
                })
                .catch(err => {
                    document.getElementById('packageHistoryMeta').textContent = `Failed to load saved logs: ${err}`;
                    packageHistoryHasMore = false;
                    updatePackageHistoryLoadMoreButton();
                })
                .finally(() => {
                    packageHistoryLoading = false;
                    updatePackageHistoryLoadMoreButton();
                });
        }

        function exportSelectedPackageHistory(filteredOnly = true) {
            const select = document.getElementById('packageHistorySessionSelect');
            const keyword1 = document.getElementById('packageHistoryFilterInput')?.value || '';
            const keyword2 = document.getElementById('packageHistoryFilterInput2')?.value || '';
            const keyword3 = document.getElementById('packageHistoryFilterInput4')?.value || '';
            const tagKeyword = document.getElementById('packageHistoryFilterInput3')?.value || '';
            const quickId = document.getElementById('packageHistoryQuickFilter')?.value || '';
            if (!select || !select.value) {
                alert('Please select a recorded session first.');
                return;
            }
            fetch('/api/package-log/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    session_id: Number(select.value),
                    q1: filteredOnly ? keyword1 : '',
                    q2: filteredOnly ? keyword2 : '',
                    q3: filteredOnly ? tagKeyword : '',
                    q4: filteredOnly ? keyword3 : '',
                    qf: filteredOnly ? quickId : '',
                    qp: packageHistoryPlatform,
                    filtered_only: filteredOnly,
                }),
            })
                .then(r => r.json())
                .then(data => {
                    if (data.cancelled) return;
                    if (!data.ok) throw new Error(data.error || 'failed_to_export_logs');
                    alert(`Exported ${data.row_count} rows to:\n${data.path}`);
                })
                .catch(err => {
                    alert(`Failed to export recorded log: ${err}`);
                });
        }

        function clearAllRecordedPackageHistory() {
            const confirmed = confirm('Clear all recorded package logs from the database? This cannot be undone.');
            if (!confirmed) return;
            fetch('/api/package-log/clear', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ clear_all: true }),
            })
                .then(r => r.json())
                .then(data => {
                    if (!data.ok) throw new Error(data.error || 'failed_to_clear_package_logs');
                    packageHistorySessions = [];
                    const select = document.getElementById('packageHistorySessionSelect');
                    if (select) {
                        select.innerHTML = '<option value="">No saved sessions</option>';
                    }
                    packageHistoryLoadedRows = [];
                    packageHistoryOffset = 0;
                    packageHistoryHasMore = false;
                    updatePackageHistoryLoadMoreButton();
                    document.getElementById('packageHistoryFilterInput').value = '';
                    document.getElementById('packageHistoryFilterInput2').value = '';
                    document.getElementById('packageHistoryFilterInput3').value = '';
                    document.getElementById('packageHistoryFilterInput4').value = '';
                    document.getElementById('packageHistoryQuickFilter').value = '';
                    document.getElementById('packageHistoryMeta').textContent = 'No saved package-log sessions yet.';
                    renderPackageHistoryRows([]);
                    alert(`Cleared ${data.deleted_sessions} session(s) and ${data.deleted_rows} log row(s).`);
                })
                .catch(err => {
                    alert(`Failed to clear recorded logs: ${err}`);
                });
        }

        function clearRecordedPackageHistoryFilters(reload = true) {
            const filter1 = document.getElementById('packageHistoryFilterInput');
            const filter2 = document.getElementById('packageHistoryFilterInput2');
            const filter3 = document.getElementById('packageHistoryFilterInput3');
            const filter4 = document.getElementById('packageHistoryFilterInput4');
            if (filter1) filter1.value = '';
            if (filter2) filter2.value = '';
            if (filter3) filter3.value = '';
            if (filter4) filter4.value = '';
            const quickSelect = document.getElementById('packageHistoryQuickFilter');
            if (quickSelect) quickSelect.value = '';
            if (reload) loadSelectedPackageHistory(false);
        }

        // --- JSON Modal Handler ---
        document.body.addEventListener('click', (e) => {
            if (e.target && e.target.classList.contains('view-json-btn')) {
                const jsonData = e.target.getAttribute('data-json'); // Use getAttribute for safety
                try {
                    const parsed = JSON.parse(jsonData);
                    jsonContent.textContent = JSON.stringify(parsed, null, 2);
                    jsonModal.classList.remove('hidden');
                } catch(err) { 
                    console.error(err);
                    alert('Invalid JSON data. Check console for details.'); 
                }
            }
        });
        closeJsonModal.addEventListener('click', () => jsonModal.classList.add('hidden'));
        closeLogDetailModal.addEventListener('click', () => logDetailModal.classList.add('hidden'));

        // --- Log Detail Modal (Package Log) ---
        let selectedPackageHistoryRowKeys = new Set();
        let isSelectingHistoryRows = false;
        let historyDragStartIndex = null;
        const PACKAGE_HISTORY_PAGE_SIZE = 2000;
        let packageHistoryLoadedRows = [];
        let packageHistoryOffset = 0;
        let packageHistoryHasMore = false;
        let packageHistoryLoading = false;

        function updatePackageHistoryLoadMoreButton() {
            const btn = document.getElementById('loadMorePackageHistoryBtn');
            if (!btn) return;
            btn.classList.toggle('hidden', !packageHistoryHasMore);
            btn.disabled = !packageHistoryHasMore || packageHistoryLoading;
            btn.textContent = packageHistoryLoading ? 'Loading...' : 'Load More';
        }

        function getSelectedPackageRows() {
            // Live rows are data objects (most selected rows are not in the
            // DOM with a virtualized table).  Keep the on-screen order.
            if (selectedPackageRowKeys.size === 0) return [];
            const fromView = packageView.filter(row => selectedPackageRowKeys.has(getPackageRowKey(row)));
            if (fromView.length > 0) return fromView;
            return lastPackageLogs.filter(row => selectedPackageRowKeys.has(getPackageRowKey(row)));
        }

        function packageRowCellTexts(row) {
            if (row && typeof row.querySelectorAll === 'function') {
                return Array.from(row.querySelectorAll('td')).map(td => (td.textContent || '').trim());
            }
            return [row?.time_display || row?.time || '', row?.tag || '', row?.message || row?.log || '']
                .map(value => String(value).trim());
        }

        function getSelectedPackageHistoryRows() {
            const selected = Array.from(document.querySelectorAll('#packageHistoryTableBody tr.package-history-row.selected'));
            if (selected.length > 0) return selected;
            const allRows = Array.from(document.querySelectorAll('#packageHistoryTableBody tr.package-history-row'));
            const byKey = new Map();
            allRows.forEach(r => {
                const rawKey = r.getAttribute('data-row-key') || '';
                try { byKey.set(decodeURIComponent(rawKey), r); } catch {}
            });
            const fromKeys = [];
            selectedPackageHistoryRowKeys.forEach(k => {
                if (byKey.has(k)) fromKeys.push(byKey.get(k));
            });
            return fromKeys;
        }

        function clearPackageHistorySelection() {
            selectedPackageHistoryRowKeys.clear();
            document.querySelectorAll('#packageHistoryTableBody tr.package-history-row.selected').forEach(r => r.classList.remove('selected'));
            isSelectingHistoryRows = false;
            historyDragStartIndex = null;
        }

        function getActiveSelectedLogRows() {
            if (packageHistoryModal && !packageHistoryModal.classList.contains('hidden')) {
                const historyRows = getSelectedPackageHistoryRows();
                if (historyRows.length > 0) return historyRows;
            }
            return getSelectedPackageRows();
        }

        function getRowText(row) {
            return packageRowCellTexts(row).join(' ');
        }

        function getRowMessageText(row) {
            if (row && typeof row.querySelector === 'function') {
                const msgCell = row.querySelector('td.col-message');
                return (msgCell ? msgCell.textContent : '').trim();
            }
            return String(row?.message || row?.log || '').trim();
        }

        let logDetailRawText = '';
        let logDetailIsJsonView = false;

        function openLogDetailModal(rows) {
            const lines = rows.map(getRowText);
            logDetailRawText = lines.join('\\n');
            logDetailIsJsonView = false;
            logDetailContent.textContent = logDetailRawText;
            // Always above Recorded Logs / session search (z 50-60) and the
            // quick filter editor (70); equal z-index stacking differs between
            // the Windows and macOS webviews.
            logDetailModal.style.zIndex = '80';
            logDetailModal.classList.remove('hidden');
        }

        function extractJsonFromText(text) {
            if (!text) return null;
            const starts = [];
            for (let i = 0; i < text.length; i++) {
                const ch = text[i];
                if (ch === '{' || ch === '[') starts.push(i);
            }
            for (const start of starts) {
                let openBraces = 0;
                let openBrackets = 0;
                for (let i = start; i < text.length; i++) {
                    const ch = text[i];
                    if (ch === '{') openBraces++;
                    if (ch === '}') openBraces--;
                    if (ch === '[') openBrackets++;
                    if (ch === ']') openBrackets--;
                    if (openBraces === 0 && openBrackets === 0 && i > start) {
                        const candidate = text.slice(start, i + 1);
                        try {
                            JSON.parse(candidate);
                            return candidate;
                        } catch (e) {}
                        break;
                    }
                }
            }
            return null;
        }

        function extractEventName(text) {
            if (!text) return '';
            const m = text.match(/with\\s+name\\s+([\\w\\.:-]+)/i);
            if (m && m[1]) return m[1];
            return '';
        }

        function extractHeaderBeforeJson(text) {
            if (!text) return '';
            const idx = text.indexOf('{');
            if (idx === -1) return '';
            let header = text.slice(0, idx).trim();
            // Remove leading time/tag/prefix if present
            const i1 = header.indexOf('] : ');
            if (i1 !== -1) header = header.slice(i1 + 4).trim();
            const i2 = header.indexOf('] ');
            if (i2 !== -1) header = header.slice(i2 + 2).trim();
            return header;
        }

        function convertSelectedLogsToJson() {
            if (logDetailIsJsonView) {
                logDetailContent.textContent = logDetailRawText;
                logDetailIsJsonView = false;
                return;
            }
            const rows = getActiveSelectedLogRows();
            if (rows.length === 0) return;
            const outputs = rows.map((row, idx) => {
                const line = getRowText(row);
                const msg = getRowMessageText(row);
                const jsonStr = extractJsonFromText(msg || line);
                const evt = extractEventName(line);
                const headerLine = extractHeaderBeforeJson(line);
                if (jsonStr) {
                    try {
                        const parsed = JSON.parse(jsonStr);
                        const pretty = JSON.stringify(parsed, null, 2);
                        const metaLines = [];
                        if (headerLine) metaLines.push(`context: ${headerLine}`);
                        else if (evt) metaLines.push(`event_name: ${evt}`);
                        metaLines.push('');
                        metaLines.push('raw_log:');
                        metaLines.push(line);
                        metaLines.push('');
                        metaLines.push('extracted_json:');
                        metaLines.push(pretty);
                        return `--- #${idx + 1} (JSON) ---\\n${metaLines.join('\\n')}`;
                    } catch (e) {
                        return `--- #${idx + 1} ---\\n${line}`;
                    }
                }
                return `--- #${idx + 1} ---\\n${line}`;
            });
            logDetailContent.textContent = outputs.join('\\n\\n');
            logDetailIsJsonView = true;
        }

        convertLogJsonBtn.addEventListener('click', convertSelectedLogsToJson);

        document.getElementById('packageLogTableBody').addEventListener('dblclick', (e) => {
            const row = e.target.closest('tr.package-log-row');
            if (!row) return;
            isSelectingRows = false;
            dragStartIndex = null;
            // Ensure clicked row is included
            const rawKey = row.getAttribute('data-row-key') || '';
            try { selectedPackageRowKeys.add(decodeURIComponent(rawKey)); } catch {}
            const selected = getActiveSelectedLogRows();
            const rowsToShow = selected.length > 0 ? selected : [row];
            openLogDetailModal(rowsToShow);
        });

        document.addEventListener('keydown', (e) => {
            if (e.key !== 'Enter') return;
            const selected = getActiveSelectedLogRows();
            if (selected.length === 0) return;
            openLogDetailModal(selected);
        });
        
        // --- Device Status ---
        socket.on('device_status', (status) => {
            // A discovery pass started on the previous platform can finish
            // after a switch. Ignore that late status instead of overwriting
            // the new platform with an old Waiting/device list.
            const statusPlatform = status?.platform === 'ios'
                ? 'ios'
                : status?.platform === 'android'
                    ? 'android'
                    : '';
            // Every current backend status is platform-tagged. Reject legacy
            // or malformed payloads so an iOS list can never be rendered in
            // the Android view (or vice versa).
            if (!statusPlatform || statusPlatform !== activePlatform) return;
            const currentFilter = deviceFilter.value;
            deviceFilter.innerHTML = '<option value="all">All Devices</option>';
            const devices = Array.isArray(status.connected_devices)
                ? status.connected_devices.filter(d => !d.platform || d.platform === statusPlatform)
                : [];
            if (devices.length > 0) {
                devices.forEach(d => {
                    const opt = document.createElement('option');
                    opt.value = d.id; opt.textContent = d.display_name || d.name || d.id;
                    deviceFilter.appendChild(opt);
                });
            }
            // Restore selection if exists
            if ([...deviceFilter.options].some(o => o.value === currentFilter)) {
                deviceFilter.value = currentFilter;
                selectedDevice = currentFilter;
            } else {
                resetSelectedDeviceFilter();
            }

            if (devices.length > 0) {
                 deviceListEl.innerHTML = '<ul class="list-disc list-inside text-left">' + 
                    devices.map(d => `<li class="${escapeHTML(d.status_class || 'text-green-600')} font-semibold">${escapeHTML(d.display_name || d.name || d.id)}</li>`).join('') +
                    '</ul>';
            } else {
                deviceListEl.innerHTML = `<p class="text-orange-500">${status.message || 'Waiting...'}</p>`;
            }
        });

        socket.on('platform_status', (status) => {
            const platform = status?.platform === 'ios' ? 'ios' : 'android';
            const platformChanged = activePlatform !== platform;
            activePlatform = platform;
            if (platformBtn) platformBtn.textContent = `Platform: ${platformLabel(platform)}`;
            if (platformChanged) resetDeviceStatusUi(platform);
            syncPlatformUi();
            if (platform === 'ios') {
                installationIdState = {
                    device_id: '',
                    device_name: '',
                    package_id: '',
                    game_code: '',
                    installation_id: '',
                    platform,
                };
                renderInstallationIdPanel();
            }
        });

        socket.on('installation_id_status', (status) => {
            if (
                selectedDevice !== 'all' &&
                status?.device_id &&
                status.device_id !== selectedDevice
            ) {
                return;
            }
            installationIdState = {
                device_id: status?.device_id || '',
                device_name: status?.device_name || '',
                package_id: status?.package_id || '',
                game_code: status?.game_code || '',
                installation_id: status?.installation_id || '',
                platform: status?.platform || activePlatform || 'android',
            };
            renderInstallationIdPanel();
        });
        
        // --- FIXED: Trigger refresh on device change to update all tables including callback
        deviceFilter.addEventListener('change', (e) => { 
            selectedDevice = e.target.value; 
            installationIdState = {
                device_id: '',
                device_name: '',
                package_id: '',
                game_code: '',
                installation_id: '',
                platform: activePlatform || 'android',
            };
            renderInstallationIdPanel();
            socket.emit('refresh_request'); 
            renderCallbackTable(); // Trigger client-side re-render immediately
            renderAdRevenueTable();
            renderPriceRotationTable();
            renderDefaultAdEventMatrix();
        });

        // --- Specific Tab Logic ---
        function setValidationButtonState(isActive) {
            const btn = document.getElementById('startValidationBtn');
            if (!btn) return;
            if (isActive) {
                btn.textContent = 'Stop';
                btn.className = 'bg-red-500 hover:bg-red-600 text-white font-semibold text-xs px-4 rounded-lg h-9';
            } else {
                btn.textContent = 'Start Checking';
                btn.className = 'bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs px-4 rounded-lg h-9';
            }
        }

        document.getElementById('startValidationBtn').addEventListener('click', () => {
            const btn = document.getElementById('startValidationBtn');
            const isStarting = btn && btn.textContent === 'Start Checking';
            if (isStarting) socket.emit('start_validation', []);
            else socket.emit('stop_validation');
        });

        document.getElementById('clearValidatorFilterBtn')?.addEventListener('click', () => {
            const eventInput = document.getElementById('validatorEventFilterInput');
            const rawInput = document.getElementById('validatorRawFilterInput');
            if (eventInput) eventInput.value = '';
            if (rawInput) rawInput.value = '';
            const allRadio = document.querySelector('input[name="validatorSourceFilter"][value="all"]');
            if (allRadio) allRadio.checked = true;
            renderValidatorTable(validator_results_cache);
        });
        
        document.getElementById('validatorEventFilterInput').addEventListener('input', () => {
            renderValidatorTable(validator_results_cache);
        });
        document.getElementById('validatorRawFilterInput').addEventListener('input', () => {
            renderValidatorTable(validator_results_cache);
        });

        document.querySelectorAll('input[name="validatorSourceFilter"]').forEach(r => {
            r.addEventListener('change', () => renderValidatorTable(validator_results_cache));
        });

        socket.on('validator_status', (data) => {
            setValidationButtonState(!!(data && data.active));
        });

        const specificEventInput = document.getElementById('specificEventInput');
        const specificParamInput = document.getElementById('specificParamInput');
        const updateSpecific = () => {
            socket.emit('update_specific_filter', {
                eventNames: specificEventInput.value.split('\\n').filter(p => p.trim()),
                params: []
            });
            renderSpecificEventTable();
        };
        specificEventInput.addEventListener('input', updateSpecific);
        specificParamInput.addEventListener('input', renderSpecificEventTable);
        document.querySelectorAll('input[name="specificSourceFilter"]').forEach(r => {
            r.addEventListener('change', () => renderSpecificEventTable());
        });
        
        document.getElementById('adRevenueFilterInput').addEventListener('input', renderAdRevenueTable);
        document.querySelectorAll('input[name="adRevenueSourceFilter"]').forEach(r => r.addEventListener('change', renderAdRevenueTable));
        let packageFilterRenderTimer = null;
        const schedulePackageRender = () => {
            clearTimeout(packageFilterRenderTimer);
            packageFilterRenderTimer = setTimeout(() => renderPackageLogTable(true), 120);
        };
        document.getElementById('packageFilterInput').addEventListener('input', schedulePackageRender);
        document.getElementById('packageFilterInput2').addEventListener('input', schedulePackageRender);
        document.getElementById('packageFilterInput3').addEventListener('input', schedulePackageRender);
        // One click empties every typed filter and goes back to the All quick filter.
        document.getElementById('clearPackageFiltersBtn')?.addEventListener('click', () => {
            ['packageTagFilterInput', 'packageFilterInput', 'packageFilterInput2', 'packageFilterInput3'].forEach(id => {
                const input = document.getElementById(id);
                if (input) input.value = '';
            });
            const allOpt = document.querySelector('input[name="tagQuickFilter"][value=""]');
            if (allOpt) allOpt.checked = true;
            renderPackageLogTable(true);
        });
        document.getElementById('packageTagFilterInput').addEventListener('input', () => {
            const allOpt = document.querySelector('input[name="tagQuickFilter"][value=""]');
            if (allOpt) allOpt.checked = true;
            schedulePackageRender();
        });
        document.getElementById('packageQuickFilterPanel')?.addEventListener('change', (e) => {
            if (e.target?.name !== 'tagQuickFilter') return;
            // Choosing a quick filter starts a fresh search: clear the typed filters.
            ['packageTagFilterInput', 'packageFilterInput', 'packageFilterInput2', 'packageFilterInput3'].forEach(id => {
                const input = document.getElementById(id);
                if (input) input.value = '';
            });
            renderPackageLogTable(true);
        });
        document.getElementById('packageFilterInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); renderPackageLogTable(true); }
        });
        document.getElementById('packageFilterInput3').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); renderPackageLogTable(true); }
        });
        document.getElementById('packageFilterInput2').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); renderPackageLogTable(true); }
        });
        document.getElementById('packageTagFilterInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); renderPackageLogTable(true); }
        });
        document.getElementById('showErrorsOnly').addEventListener('change', () => renderPackageLogTable(true));
        document.getElementById('showWarningsOnly').addEventListener('change', () => renderPackageLogTable(true));
        document.getElementById('loadPackageHistoryBtn').addEventListener('click', () => loadSelectedPackageHistory(false));
        document.getElementById('loadMorePackageHistoryBtn').addEventListener('click', () => loadSelectedPackageHistory(true));
        document.getElementById('exportPackageHistoryFilteredBtn').addEventListener('click', () => exportSelectedPackageHistory(true));
        document.getElementById('exportPackageHistoryAllBtn').addEventListener('click', () => exportSelectedPackageHistory(false));
        document.getElementById('refreshPackageSessionsBtn').addEventListener('click', loadPackageHistorySessions);
        document.getElementById('clearPackageHistoryBtn').addEventListener('click', clearAllRecordedPackageHistory);
        document.getElementById('clearPackageHistoryFiltersBtn').addEventListener('click', () => clearRecordedPackageHistoryFilters(true));
        document.getElementById('packageHistorySessionSelect').addEventListener('change', () => {
            clearRecordedPackageHistoryFilters(false);
            loadSelectedPackageHistory(false);
        });
        document.getElementById('packageHistoryFilterInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                loadSelectedPackageHistory(false);
            }
        });
        document.getElementById('packageHistoryFilterInput4').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                loadSelectedPackageHistory(false);
            }
        });
        document.getElementById('packageHistoryFilterInput2').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                loadSelectedPackageHistory(false);
            }
        });
        document.getElementById('packageHistoryQuickFilter')?.addEventListener('change', () => {
            ['packageHistoryFilterInput', 'packageHistoryFilterInput2', 'packageHistoryFilterInput3', 'packageHistoryFilterInput4'].forEach(id => {
                const input = document.getElementById(id);
                if (input) input.value = '';
            });
            loadSelectedPackageHistory(false);
        });
        document.getElementById('packageHistoryFilterInput3').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                loadSelectedPackageHistory(false);
            }
        });
        const packageHistoryModal = document.getElementById('packageHistoryModal');
        const openPackageHistoryBtn = document.getElementById('openPackageHistoryBtn');
        const closePackageHistoryModal = document.getElementById('closePackageHistoryModal');
        openPackageHistoryBtn?.addEventListener('click', () => {
            packageHistoryModal?.classList.remove('hidden');
            clearRecordedPackageHistoryFilters(false);
            setPackageHistoryPlatform(activePlatform, false);
            loadPackageHistorySessions();
        });
        closePackageHistoryModal?.addEventListener('click', () => {
            packageHistoryModal?.classList.add('hidden');
        });
        packageHistoryModal?.addEventListener('click', (e) => {
            if (e.target === packageHistoryModal) {
                clearPackageHistorySelection();
                packageHistoryModal.classList.add('hidden');
                return;
            }
            if (!e.target.closest('#packageHistoryTableBody tr.package-history-row') && !e.target.closest('#logDetailModal')) {
                clearPackageHistorySelection();
            }
        });

        const packageHistoryTableWrap = document.getElementById('packageHistoryTableWrap');
        packageHistoryTableWrap?.addEventListener('scroll', () => {
            if (!packageHistoryHasMore || packageHistoryLoading) return;
            const remaining = packageHistoryTableWrap.scrollHeight - packageHistoryTableWrap.scrollTop - packageHistoryTableWrap.clientHeight;
            if (remaining < 120) {
                loadSelectedPackageHistory(true);
            }
        });

        ['packageHistorySessionSelect','loadPackageHistoryBtn','loadMorePackageHistoryBtn','refreshPackageSessionsBtn','exportPackageHistoryAllBtn','exportPackageHistoryFilteredBtn','clearPackageHistoryBtn','clearPackageHistoryFiltersBtn','packageHistoryFilterInput','packageHistoryFilterInput2','packageHistoryFilterInput3','packageHistoryFilterInput4','packageHistoryQuickFilter','closePackageHistoryModal'].forEach(id => {
            document.getElementById(id)?.addEventListener('focus', clearPackageHistorySelection);
            document.getElementById(id)?.addEventListener('click', () => { clearPackageHistorySelection(); });
        });

        // --- Row Selection for Package Log (click + drag) ---
        let isSelectingRows = false;
        let dragStartIndex = null;

        let packageSelectionAnchor = null;

        function updateSelectionRange(startIdx, endIdx) {
            // Indices address packageView, so a range can span rows that are
            // scrolled out of the DOM.
            if (packageView.length === 0) return;
            const minIdx = Math.max(0, Math.min(startIdx, endIdx));
            const maxIdx = Math.min(packageView.length - 1, Math.max(startIdx, endIdx));
            selectedPackageRowKeys.clear();
            for (let i = minIdx; i <= maxIdx; i++) selectedPackageRowKeys.add(getPackageRowKey(packageView[i]));
            document.querySelectorAll('#packageLogTableBody tr.package-log-row').forEach(r => {
                let key = '';
                try { key = decodeURIComponent(r.getAttribute('data-row-key') || ''); } catch (err) {}
                r.classList.toggle('selected', selectedPackageRowKeys.has(key));
            });
        }

        document.getElementById('packageLogTableBody').addEventListener('mousedown', (e) => {
            const row = e.target.closest('tr.package-log-row');
            if (!row) return;
            if (e.detail && e.detail >= 2) return;
            if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA')) {
                document.activeElement.blur();
            }
            // Drop a leftover browser text selection (e.g. from a double-click drag).
            window.getSelection()?.removeAllRanges();
            isSelectingRows = true;
            const idx = parseInt(row.getAttribute('data-row-index') || '0', 10);
            if (e.shiftKey && packageSelectionAnchor !== null) {
                dragStartIndex = packageSelectionAnchor;
                updateSelectionRange(packageSelectionAnchor, idx);
            } else {
                packageSelectionAnchor = idx;
                dragStartIndex = idx;
                updateSelectionRange(idx, idx);
            }
            e.preventDefault();
        });

        document.getElementById('packageLogTableBody').addEventListener('mouseover', (e) => {
            if (!isSelectingRows || dragStartIndex === null) return;
            if (e.buttons === 0) {
                // The button was released outside the window, so no mouseup
                // reached us: stop the drag instead of following the pointer.
                isSelectingRows = false;
                dragStartIndex = null;
                return;
            }
            const row = e.target.closest('tr.package-log-row');
            if (!row) return;
            const idx = parseInt(row.getAttribute('data-row-index') || '0', 10);
            updateSelectionRange(dragStartIndex, idx);
        });

        function updateHistorySelectionRange(startIdx, endIdx) {
            const rows = Array.from(document.querySelectorAll('#packageHistoryTableBody tr.package-history-row'));
            if (rows.length === 0) return;
            const [minIdx, maxIdx] = startIdx <= endIdx ? [startIdx, endIdx] : [endIdx, startIdx];
            selectedPackageHistoryRowKeys.clear();
            rows.forEach(r => {
                const idx = parseInt(r.getAttribute('data-row-index') || '-1', 10);
                if (idx >= minIdx && idx <= maxIdx) {
                    const rawKey = r.getAttribute('data-row-key') || '';
                    try { selectedPackageHistoryRowKeys.add(decodeURIComponent(rawKey)); } catch (err) {}
                    r.classList.add('selected');
                } else {
                    r.classList.remove('selected');
                }
            });
        }

        document.getElementById('packageHistoryTableBody').addEventListener('mousedown', (e) => {
            const row = e.target.closest('tr.package-history-row');
            if (!row) return;
            if (e.detail && e.detail >= 2) return;
            if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA')) {
                document.activeElement.blur();
            }
            isSelectingHistoryRows = true;
            const idx = parseInt(row.getAttribute('data-row-index') || '0', 10);
            historyDragStartIndex = idx;
            updateHistorySelectionRange(idx, idx);
            e.preventDefault();
        });

        document.getElementById('packageHistoryTableBody').addEventListener('mouseover', (e) => {
            if (!isSelectingHistoryRows || historyDragStartIndex === null) return;
            if (e.buttons === 0) {
                isSelectingHistoryRows = false;
                historyDragStartIndex = null;
                return;
            }
            const row = e.target.closest('tr.package-history-row');
            if (!row) return;
            const idx = parseInt(row.getAttribute('data-row-index') || '0', 10);
            updateHistorySelectionRange(historyDragStartIndex, idx);
        });

        document.getElementById('packageHistoryTableBody').addEventListener('dblclick', (e) => {
            const row = e.target.closest('tr.package-history-row');
            if (!row) return;
            isSelectingHistoryRows = false;
            historyDragStartIndex = null;
            const rawKey = row.getAttribute('data-row-key') || '';
            try { selectedPackageHistoryRowKeys.add(decodeURIComponent(rawKey)); } catch {}
            const selected = getSelectedPackageHistoryRows();
            openLogDetailModal(selected.length > 0 ? selected : [row]);
        });

        function endPackageRowDrags() {
            isSelectingRows = false;
            dragStartIndex = null;
            isSelectingHistoryRows = false;
            historyDragStartIndex = null;
        }
        document.addEventListener('mouseup', endPackageRowDrags);
        window.addEventListener('blur', endPackageRowDrags);

        // Only a click on empty space clears the selected rows (and any
        // browser text selection in the log).  Buttons, inputs and quick
        // filters keep the selection, so a row picked in a filtered view can
        // be followed into the full log.
        function clearPackageLiveSelection() {
            endPackageRowDrags();
            packageSelectionAnchor = null;
            const sel = window.getSelection();
            const body = document.getElementById('packageLogTableBody');
            if (sel && sel.rangeCount && body && body.contains(sel.anchorNode)) sel.removeAllRanges();
            if (!selectedPackageRowKeys.size) return;
            selectedPackageRowKeys.clear();
            document.querySelectorAll('#packageLogTableBody tr.package-log-row.selected').forEach(r => r.classList.remove('selected'));
        }
        document.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            const target = e.target;
            if (target.closest('#logDetailModal, #packageQuickFilterEditor, #packageNewRowsBtn')) return;
            const container = document.getElementById('packageLogContainer');
            if (container && container.contains(target)) {
                if (target.closest('tr.package-log-row, thead')) return;
                // A click on the scrollbar lands outside the client area.
                const rect = container.getBoundingClientRect();
                if (e.clientX - rect.left >= container.clientWidth || e.clientY - rect.top >= container.clientHeight) return;
                clearPackageLiveSelection();
                return;
            }
            if (target.closest('button, input, select, textarea, label, a, [role="button"], .resizer')) return;
            clearPackageLiveSelection();
        });

        // --- Ctrl+C to copy selected package log rows ---
        document.addEventListener('keydown', async (e) => {
            if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 'c') return;
            const active = document.activeElement;
            if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')) return;
            const selectedRows = getActiveSelectedLogRows();
            if (selectedRows.length === 0) return;
            const lines = selectedRows.map(r => packageRowCellTexts(r).join('\\t'));
            const textToCopy = lines.join('\\n');
            try {
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    await navigator.clipboard.writeText(textToCopy);
                } else {
                    const ta = document.createElement('textarea');
                    ta.value = textToCopy;
                    ta.style.position = 'fixed';
                    ta.style.opacity = '0';
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                }
            } catch (err) {
                console.error('Copy failed', err);
            }
        });
        
        function renderSdkCheckPresetOptions() {
            const container = document.getElementById('sdkCheckPresetList');
            if (!container) return;
            const selectedPresetName = Array.from(container.querySelectorAll('input[name="sdkCheckPreset"]'))
                .find(input => input.checked)?.value || '';
            const presets = Object.entries(sdkCheckPresets || {})
                .filter(([, preset]) => preset && (preset.platform === 'all' || preset.platform === activePlatform));
            if (presets.length === 0) {
                container.innerHTML = '<span class="text-xs text-slate-500">Chưa có preset.</span>';
                return;
            }
            container.innerHTML = presets.map(([name, preset]) => `
                <label class="inline-flex items-center gap-2 rounded-md border border-indigo-200 bg-white px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm cursor-pointer hover:bg-indigo-50">
                    <input type="radio" name="sdkCheckPreset" value="${escapeAttribute(name)}" class="h-4 w-4 text-indigo-600 border-gray-300 focus:ring-indigo-500"${sdkCheckClearPending ? ' disabled' : ''}>
                    <span>${escapeHTML(name)}</span>
                    <span class="text-[11px] font-normal text-slate-500">${Array.isArray(preset.lines) ? preset.lines.length - 1 : 0} SDKs</span>
                </label>
            `).join('');
            if (selectedPresetName) {
                container.querySelectorAll('input[name="sdkCheckPreset"]').forEach(input => {
                    input.checked = input.value === selectedPresetName;
                });
            }
        }

        let sdkCheckPresetsLoadPromise = null;
        async function loadSdkCheckPresetsFromGit(force = false) {
            if (force) sdkCheckPresetsLoadPromise = null;
            if (sdkCheckPresetsLoadPromise) return sdkCheckPresetsLoadPromise;
            const status = document.getElementById('sdkCheckPresetStatus');
            if (status) status.textContent = 'Đang tải danh sách từ GitHub...';
            sdkCheckPresetsLoadPromise = (async () => {
                try {
                    const refreshQuery = force ? '&refresh=1' : '';
                    const response = await fetch(`/api/sdk-check-presets?ts=${Date.now()}${refreshQuery}`, {cache: 'no-store'});
                    const payload = await response.json();
                    if (!response.ok || !payload.ok || !payload.presets) throw new Error(payload.error || 'preset_load_failed');
                    sdkCheckPresets = payload.presets;
                    renderSdkCheckPresetOptions();
                    renderBrightSkadPresetOptions();
                    const sourceText = payload.source === 'github' ? 'GitHub' : 'Local fallback';
                    if (status) status.textContent = `${sourceText}: chọn preset để tự nạp danh sách.`;
                } catch (error) {
                    renderSdkCheckPresetOptions();
                    renderBrightSkadPresetOptions();
                    if (status) status.textContent = 'Không tải được GitHub, đang dùng danh sách local.';
                    console.warn('SDK preset load failed', error);
                } finally {
                    sdkCheckPresetsLoadPromise = null;
                }
            })();
            return sdkCheckPresetsLoadPromise;
        }

        function applySdkCheckPreset(name) {
            if (sdkCheckClearPending) return;
            const preset = sdkCheckPresets?.[name];
            if (!preset || (preset.platform !== 'all' && preset.platform !== activePlatform)) return;
            const lines = Array.isArray(preset?.lines) ? preset.lines.filter(line => String(line).trim()) : [];
            const input = document.getElementById('sdkCheckInput');
            if (!input || !lines.length) return;
            input.value = lines.join('\\n');
            // start_sdk_check atomically replaces the previous preset/state. Do
            // not send stop first: concurrent Socket.IO handlers could leave
            // the newly selected check inactive and make the table appear stuck.
            socket.emit('start_sdk_check', {text: input.value});
            sdkCheckRunning = true;
            const btn = document.getElementById('startSdkCheckBtn');
            if (btn) btn.textContent = 'Stop Checking';
            const status = document.getElementById('sdkCheckPresetStatus');
            if (status) status.textContent = `${name} loaded and checking (${lines.length - 1} SDKs).`;
        }

        renderSdkCheckPresetOptions();
        renderBrightSkadPresetOptions();
        document.getElementById('sdkCheckPresetList')?.addEventListener('change', (event) => {
            const input = event.target.closest('input[name="sdkCheckPreset"]');
            if (input?.checked) applySdkCheckPreset(input.value);
        });

        socket.on('clear_all_logs_complete', () => {
            sdkCheckClearPending = false;
            renderSdkCheckPresetOptions();
            lastPackageLogs = [];
            selectedPackageRowKeys.clear();
            pausedPackageSnapshot = [];
            renderPackageLogTable(true);
            const status = document.getElementById('sdkCheckPresetStatus');
            if (status && !sdkCheckRunning) status.textContent = 'Chọn preset để tự nạp danh sách.';
        });

        document.getElementById('sdkCheckInput')?.addEventListener('input', () => {
            document.querySelectorAll('input[name="sdkCheckPreset"]').forEach(input => input.checked = false);
            const status = document.getElementById('sdkCheckPresetStatus');
            if (status) status.textContent = 'Manual input đang được sử dụng.';
        });

        document.getElementById('reloadSdkCheckPresetsBtn')?.addEventListener('click', () => {
            loadSdkCheckPresetsFromGit(true);
        });

        document.getElementById('reloadBrightSkadPresetsBtn')?.addEventListener('click', async (event) => {
            const button = event.currentTarget;
            if (!button) return;
            const originalText = button.textContent || 'Reload';
            button.disabled = true;
            button.textContent = 'Reloading...';
            try {
                await loadSdkCheckPresetsFromGit(true);
            } finally {
                button.disabled = false;
                button.textContent = originalText;
            }
        });

        // Refresh remote presets on every app/page start so an edited GitHub list is available immediately.
        loadSdkCheckPresetsFromGit();

        document.getElementById('startSdkCheckBtn').addEventListener('click', () => {
             const btn = document.getElementById('startSdkCheckBtn');
             if (sdkCheckRunning) {
                 socket.emit('stop_sdk_check');
                 sdkCheckRunning = false;
                 btn.textContent = 'Start Checking';
                 return;
             }
             const text = document.getElementById('sdkCheckInput').value;
             if(text) {
                 socket.emit('start_sdk_check', {text: text});
                 sdkCheckRunning = true;
                 btn.textContent = 'Stop Checking';
             }
        });
        
        function setPackageControlsEnabled(enabled) {
            const packageIdInput = document.getElementById('packageIdInput');
            const packageCheckboxes = document.querySelectorAll('.package-id-checkbox');
            if (packageIdInput) packageIdInput.disabled = !enabled;
            packageCheckboxes.forEach(cb => cb.disabled = !enabled);
        }

        const packageStreamModal = document.getElementById('packageStreamModal');
        const startPackageLogBtn = document.getElementById('startPackageLogBtn');
        const pausePackageLogBtn = document.getElementById('pausePackageLogBtn');
        const stopPackageLogBtn = document.getElementById('stopPackageLogBtn');

        function setPackageRunningState(isRunning) {
            if (!startPackageLogBtn) return;
            startPackageLogBtn.textContent = isRunning ? 'Running...' : 'Start';
            startPackageLogBtn.disabled = isRunning;
            startPackageLogBtn.classList.toggle('opacity-60', isRunning);
            startPackageLogBtn.classList.toggle('cursor-not-allowed', isRunning);
            setPackageControlsEnabled(!isRunning);
        }

        function openPackageStreamModal() {
            packageStreamModal?.classList.remove('hidden');
        }

        function closePackageStreamModal() {
            packageStreamModal?.classList.add('hidden');
        }

        // Project switcher inside the Stream window.  Switching works like
        // Stop + Start outside: the server closes and keeps the current
        // session in history, then starts a new session for the new project.
        var packageStreamTarget = '';
        function renderPackageStreamProjects() {
            const select = document.getElementById('packageStreamProjectSelect');
            if (!select) return;
            const platform = activePlatform === 'ios' ? 'ios' : 'android';
            const items = (typeof packageLogPresets !== 'undefined' ? packageLogPresets : [])
                .filter(project => project && project[platform]?.id)
                .map(project => ({ id: project[platform].id, label: project[platform].label || project[platform].id }));
            if (packageStreamTarget && !items.some(item => item.id === packageStreamTarget)) {
                items.unshift({ id: packageStreamTarget, label: `${packageStreamTarget} (custom)` });
            }
            select.innerHTML = items.length
                ? items.map(item => `<option value="${qfEsc(item.id)}">${qfEsc(item.label)}</option>`).join('')
                : '<option value="">No projects</option>';
            select.value = packageStreamTarget || '';
        }
        function switchPackageStreamProject(pkg) {
            pkg = String(pkg || '').trim();
            if (!pkg || pkg === packageStreamTarget) return;
            setPackagePauseState(false);
            resetPackageLogUiState();
            const input = document.getElementById('packageIdInput');
            if (input) input.value = pkg;
            document.querySelectorAll('.package-id-checkbox').forEach(cb => { cb.checked = cb.value === pkg; });
            packageStreamTarget = pkg;
            socket.emit('start_package_log', {package_id: pkg});
            setPackageRunningState(true);
            renderPackageStreamProjects();
        }
        document.getElementById('packageStreamProjectSelect')?.addEventListener('change', (e) => {
            switchPackageStreamProject(e.target.value);
        });

        function resetPackageLogUiState() {
            const packageIdInput = document.getElementById('packageIdInput');
            const tagInput = document.getElementById('packageTagFilterInput');
            const filter1 = document.getElementById('packageFilterInput');
            const filter2 = document.getElementById('packageFilterInput2');
            const filter3 = document.getElementById('packageFilterInput3');
            const showErrorsOnly = document.getElementById('showErrorsOnly');
            const showWarningsOnly = document.getElementById('showWarningsOnly');
            const autoScroll = document.getElementById('autoScroll');
            const allQuickTag = document.querySelector('input[name="tagQuickFilter"][value=""]');
            const packageContainer = document.getElementById('packageLogContainer');

            if (packageIdInput) packageIdInput.value = '';
            document.querySelectorAll('.package-id-checkbox').forEach(cb => { cb.checked = false; });
            if (tagInput) tagInput.value = '';
            if (filter1) filter1.value = '';
            if (filter2) filter2.value = '';
            if (filter3) filter3.value = '';
            if (showErrorsOnly) showErrorsOnly.checked = false;
            if (showWarningsOnly) showWarningsOnly.checked = false;
            if (autoScroll) autoScroll.checked = true;
            if (allQuickTag) allQuickTag.checked = true;
            packageSelectedRowIndex = null;
            packageDragAnchor = null;
            packageLastClickedIndex = null;
            packageUiPaused = false;
            pausedPackageSnapshot = [];
            packageStickToBottom = true;
            packageNewRowsWhileAway = 0;
            updatePackageNewRowsButton();
            updatePackagePauseLabel();
            if (packageContainer) packageContainer.scrollTop = 0;
            renderPackageLogTable(true);
        }

        function updatePackagePauseLabel() {
            if (!pausePackageLogBtn) return;
            if (!packageUiPaused) {
                pausePackageLogBtn.textContent = 'Pause';
            } else {
                pausePackageLogBtn.textContent = packagePausedNewRows > 0
                    ? `Resume (+${packagePausedNewRows.toLocaleString()})`
                    : 'Resume';
            }
        }

        function setPackagePauseState(isPaused) {
            // Pause freezes the view only.  Rows keep arriving in the store,
            // so Resume shows everything received in the meantime.
            packageUiPaused = isPaused;
            packagePausedAtSeq = lastPackageSeq;
            packagePausedNewRows = 0;
            pausedPackageSnapshot = [];
            if (pausePackageLogBtn) {
                pausePackageLogBtn.classList.toggle('bg-amber-500', !isPaused);
                pausePackageLogBtn.classList.toggle('hover:bg-amber-600', !isPaused);
                pausePackageLogBtn.classList.toggle('bg-emerald-500', isPaused);
                pausePackageLogBtn.classList.toggle('hover:bg-emerald-600', isPaused);
            }
            updatePackagePauseLabel();
            renderPackageLogTable(true);
        }

        startPackageLogBtn?.addEventListener('click', () => {
             const packageInput = document.getElementById('packageIdInput');
             const pkg = String(packageInput?.value || '').trim();
             if (!pkg) {
                 alert(activePlatform === 'ios'
                     ? 'Chọn hoặc nhập Bundle Search trước khi start Package Log.'
                     : 'Chọn hoặc nhập Package ID trước khi start Package Log.');
                 packageInput?.focus();
                 return;
             }
             setPackagePauseState(false);
             packageStreamTarget = pkg;
             renderPackageStreamProjects();
             openPackageStreamModal();
             socket.emit('start_package_log', {package_id: pkg});
             setPackageRunningState(true);
        });

        pausePackageLogBtn?.addEventListener('click', () => {
            setPackagePauseState(!packageUiPaused);
        });

        stopPackageLogBtn?.addEventListener('click', () => {
            packageStreamTarget = '';
            socket.emit('start_package_log', {package_id: ''});
            setPackagePauseState(false);
            closePackageStreamModal();
            setPackageRunningState(false);
            resetPackageLogUiState();
        });

        // --- Column Resizer (Package Log Table) ---
        (function setupColumnResizers() {
            let active = null;
            let startX = 0;
            let startWidth = 0;

            function onMouseMove(e) {
                if (!active) return;
                const dx = e.clientX - startX;
                const newWidth = Math.max(50, startWidth + dx);
                const colClass = active.getAttribute('data-col');
                const colEl = document.querySelector(`.col-${colClass}`);
                if (colEl) colEl.style.width = newWidth + 'px';
            }

            function onMouseUp() {
                if (!active) return;
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                active = null;
            }

            document.addEventListener('mousedown', (e) => {
                const resizer = e.target.closest('.resizer');
                if (!resizer) return;
                if (!(isPausedClient || pauseBtn.textContent === 'Resume')) return;
                e.preventDefault();
                active = resizer;
                const colClass = resizer.getAttribute('data-col');
                const headerCell = document.querySelector(`th.${colClass}`);
                startX = e.clientX;
                startWidth = headerCell ? headerCell.offsetWidth : 0;
                    document.body.style.cursor = 'col-resize';
                    document.body.style.userSelect = 'none';
                    document.addEventListener('mousemove', onMouseMove);
                    document.addEventListener('mouseup', onMouseUp);
            });
            // Apply initial state
            setResizerEnabled(isPausedClient);
        })();

        // --- Package Log Quick Select ---
        const packageIdInput = document.getElementById('packageIdInput');
        const packageQuickSelectList = document.getElementById('packageQuickSelectList');
        const packageQuickSelectStatus = document.getElementById('packageQuickSelectStatus');
        const reloadPackagePresetsBtn = document.getElementById('reloadPackagePresetsBtn');
        let packageLogPresets = [];
        let packageLogPresetsRequestId = 0;

        function renderPackageLogPresets() {
            if (!packageQuickSelectList) return;
            const currentValue = String(packageIdInput?.value || '').trim();
            packageQuickSelectList.replaceChildren();
            const platform = activePlatform === 'ios' ? 'ios' : 'android';
            const visiblePresets = packageLogPresets.filter(item => item && item[platform]?.id);
            if (!visiblePresets.length) {
                const empty = document.createElement('span');
                empty.className = 'text-slate-500 italic';
                empty.textContent = 'No projects configured for this platform.';
                packageQuickSelectList.appendChild(empty);
                return;
            }

            visiblePresets.forEach((project) => {
                const target = project[platform];
                const label = document.createElement('label');
                label.className = 'inline-flex items-center gap-2';

                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.className = 'package-id-checkbox h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500';
                checkbox.value = target.id;
                checkbox.checked = currentValue === target.id;
                checkbox.dataset.androidValue = project.android?.id || '';
                checkbox.dataset.androidLabel = project.android?.label || '';
                checkbox.dataset.iosValue = project.ios?.id || '';
                checkbox.dataset.iosLabel = project.ios?.label || '';
                checkbox.dataset.projectCode = project.code || '';
                label.appendChild(checkbox);

                const text = document.createElement('span');
                text.textContent = target.label || target.id;
                label.appendChild(text);
                packageQuickSelectList.appendChild(label);
            });
            syncPlatformUi();
            renderPackageStreamProjects();
            const running = startPackageLogBtn && startPackageLogBtn.textContent !== 'Start';
            setPackageControlsEnabled(!running);
        }

        async function loadPackageLogPresets(forceRemote = false) {
            const requestId = ++packageLogPresetsRequestId;
            if (packageQuickSelectStatus) packageQuickSelectStatus.textContent = 'Loading projects...';
            if (reloadPackagePresetsBtn) reloadPackagePresetsBtn.disabled = true;
            try {
                const refreshQuery = forceRemote ? '&refresh=1' : '';
                const response = await fetch(`/api/package-log-presets?ts=${Date.now()}${refreshQuery}`, {cache: 'no-store'});
                const payload = await response.json();
                if (requestId !== packageLogPresetsRequestId) return;
                if (!response.ok || !payload.ok || !Array.isArray(payload.presets)) {
                    throw new Error(payload.error || 'package_log_presets_unavailable');
                }
                packageLogPresets = payload.presets;
                renderPackageLogPresets();
                const sourceText = payload.source === 'github' ? 'GitHub' : 'Local file';
                if (packageQuickSelectStatus) {
                    packageQuickSelectStatus.textContent = `${sourceText}: ${packageLogPresets.length} projects`;
                }
            } catch (error) {
                if (requestId !== packageLogPresetsRequestId) return;
                renderPackageLogPresets();
                if (packageQuickSelectStatus) packageQuickSelectStatus.textContent = 'Could not load project list.';
                console.warn('Package Log preset load failed', error);
            } finally {
                if (requestId === packageLogPresetsRequestId && reloadPackagePresetsBtn) {
                    reloadPackagePresetsBtn.disabled = false;
                }
            }
        }

        packageQuickSelectList?.addEventListener('change', (event) => {
            const selected = event.target.closest('input.package-id-checkbox');
            if (!selected) return;
            const checkboxes = packageQuickSelectList.querySelectorAll('.package-id-checkbox');
            if (selected.checked) {
                checkboxes.forEach(other => { if (other !== selected) other.checked = false; });
                packageIdInput.value = selected.value;
            } else if (packageIdInput.value === selected.value) {
                packageIdInput.value = '';
            }
        });

        // Manual input clears quick select.
        packageIdInput?.addEventListener('input', () => {
            packageQuickSelectList?.querySelectorAll('.package-id-checkbox').forEach(cb => { cb.checked = false; });
        });

        // Load the local file immediately so Package Log never waits on the
        // network.  Reload list explicitly checks the GitHub catalog for new
        // projects and falls back to the same local file when offline.
        reloadPackagePresetsBtn?.addEventListener('click', () => loadPackageLogPresets(true));
        loadPackageLogPresets(false);

        // Clear All should also stop package log and re-enable inputs
        clearAllBtn.addEventListener('click', () => {
            const btn = document.getElementById('startPackageLogBtn');
            if (btn && btn.textContent !== 'Start') {
                socket.emit('start_package_log', {package_id: ''});
                setPackagePauseState(false);
                closePackageStreamModal();
                setPackageRunningState(false);
                resetPackageLogUiState();
            }
            resetBrightSdkUiState();
            setPackageControlsEnabled(true);
        });

        loadPackageHistorySessions();
        
    </script>
</body>
</html>
"""

UI_SETTINGS_FILENAME = "ui_settings.json"
UI_ZOOM_LEVELS = (0.7, 0.75, 0.8, 0.85, 0.9, 1.0, 1.1, 1.25, 1.4)
# Default text size for users who never picked one (A-/A+ still go 70-140%).
UI_DEFAULT_ZOOM = 0.85


def _ui_settings_path():
    return os.path.join(_user_data_dir(), UI_SETTINGS_FILENAME)


def _read_stored_ui_settings():
    """Only the values the user chose; anything missing uses the default."""
    stored = {}
    try:
        with open(_ui_settings_path(), "r", encoding="utf-8") as handle:
            raw = json.load(handle)
        if isinstance(raw, dict):
            try:
                zoom = float(raw.get("zoom"))
                if zoom in UI_ZOOM_LEVELS:
                    stored["zoom"] = zoom
            except (TypeError, ValueError):
                pass
            if isinstance(raw.get("hide_top"), bool):
                stored["hide_top"] = raw["hide_top"]
    except (OSError, ValueError):
        pass
    return stored


def _load_ui_settings():
    """Per-machine UI preferences (the desktop webview drops localStorage)."""
    return {"zoom": UI_DEFAULT_ZOOM, "hide_top": False, **_read_stored_ui_settings()}


@app.post('/api/ui-settings')
def ui_settings_api():
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict) or not ({"zoom", "hide_top"} & payload.keys()):
        return jsonify({"ok": False, "error": "invalid_settings"}), 400
    stored = _read_stored_ui_settings()
    if "zoom" in payload:
        try:
            zoom = float(payload["zoom"])
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "invalid_zoom"}), 400
        if zoom not in UI_ZOOM_LEVELS:
            return jsonify({"ok": False, "error": "invalid_zoom"}), 400
        stored["zoom"] = zoom
    if "hide_top" in payload:
        if not isinstance(payload["hide_top"], bool):
            return jsonify({"ok": False, "error": "invalid_hide_top"}), 400
        stored["hide_top"] = payload["hide_top"]
    # Write only what the user changed, so a later default (e.g. a new
    # default zoom) still reaches users who never picked a zoom level.
    path = _ui_settings_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(stored, handle)
    os.replace(tmp_path, path)
    return jsonify({"ok": True, **_load_ui_settings()})


# --- Package Log quick filters (user-editable, saved per machine) ---
# These are the user's own UI settings, not preset data: they live in the
# data directory next to ui_settings.json and never ship with a release.
# Android and iOS keep separate lists in separate files.  A filter works like
# the filter boxes above the table: Tag (Android only) must appear in the
# tag; Message 1 and Message 2 must each appear somewhere in the log line.
PACKAGE_QUICK_FILTER_PLATFORMS = ("android", "ios")
PACKAGE_QUICK_FILTER_FIELDS = ("tag", "message1", "message2", "message3")
PACKAGE_QUICK_FILTER_MAX = 30
DEFAULT_PACKAGE_QUICK_FILTERS = {
    "android": (
        {"id": "integrationhelper", "name": "Integrationhelper", "tag": "integrationhelper", "message1": "", "message2": "", "message3": ""},
        {"id": "appmetrica", "name": "Appmetrica", "tag": "appmetrica", "message1": "", "message2": "", "message3": ""},
        {"id": "adjust", "name": "Adjust", "tag": "adjust", "message1": "", "message2": "", "message3": ""},
        {"id": "bidding", "name": "Bidding", "tag": "", "message1": "[ad,", "message2": "bidding", "message3": ""},
        {"id": "applovinsdk", "name": "Applovinsdk", "tag": "", "message1": "applovinsdk", "message2": "", "message3": ""},
        {"id": "event_tracking", "name": "Event tracking", "tag": "", "message1": "TrackingService->Track:", "message2": '"EventName":"', "message3": ""},
        {"id": "ads_event", "name": "Ads Event", "tag": "", "message1": "TrackingService->Track:", "message2": '"EventName":"ad', "message3": ""},
    ),
    "ios": (
        {"id": "integrationhelper", "name": "Integrationhelper", "tag": "", "message1": "integrationhelper", "message2": "", "message3": ""},
        {"id": "appmetrica", "name": "Appmetrica", "tag": "", "message1": "appmetrica", "message2": "", "message3": ""},
        {"id": "adjust", "name": "Adjust", "tag": "", "message1": "adjust]", "message2": "", "message3": ""},
        {"id": "bidding", "name": "Bidding", "tag": "", "message1": "[ad,", "message2": "bidding", "message3": ""},
        {"id": "applovinsdk", "name": "Applovinsdk", "tag": "", "message1": "applovinsdk", "message2": "", "message3": ""},
        {"id": "event_tracking", "name": "Event tracking", "tag": "", "message1": "TrackingService->Track:", "message2": '"EventName":"', "message3": ""},
        {"id": "ads_event", "name": "Ads Event", "tag": "", "message1": "TrackingService->Track:", "message2": '"EventName":"ad', "message3": ""},
    ),
}
_PACKAGE_QUICK_FILTER_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,40}$")


def _quick_filter_platform(value):
    return "ios" if value == "ios" else "android"


def _package_quick_filters_path(platform):
    return os.path.join(_user_data_dir(), f"package_quick_filters_{_quick_filter_platform(platform)}.json")


def _default_package_quick_filters(platform):
    return [dict(item) for item in DEFAULT_PACKAGE_QUICK_FILTERS[_quick_filter_platform(platform)]]


def _validate_package_quick_filters(items, platform):
    """Return (filters, error) with every filter normalized."""
    platform = _quick_filter_platform(platform)
    if not isinstance(items, list):
        return None, "filters_must_be_a_list"
    if len(items) > PACKAGE_QUICK_FILTER_MAX:
        return None, "too_many_filters"
    filters, seen_ids = [], set()
    for item in items:
        if not isinstance(item, dict):
            return None, "invalid_filter"
        filter_id = str(item.get("id") or "")
        name = str(item.get("name") or "").strip()
        if not _PACKAGE_QUICK_FILTER_ID_RE.match(filter_id) or filter_id in seen_ids:
            return None, "invalid_filter_id"
        if not name or len(name) > 40:
            return None, "invalid_filter_name"
        values = {}
        for field in PACKAGE_QUICK_FILTER_FIELDS:
            raw = item.get(field, "")
            if not isinstance(raw, str) or len(raw) > 200:
                return None, "invalid_filter_text"
            values[field] = raw.strip()
        if platform == "ios" and values["tag"]:
            # iOS logs have no tag.
            return None, "ios_filter_has_no_tag"
        if not any(values.values()):
            return None, "filter_needs_text"
        seen_ids.add(filter_id)
        filters.append({"id": filter_id, "name": name, **values})
    return filters, None


def _write_package_quick_filters(platform, filters, seen_defaults):
    path = _package_quick_filters_path(platform)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump({"version": 2, "platform": _quick_filter_platform(platform), "filters": filters,
                   "seen_defaults": sorted(set(seen_defaults))}, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _backup_broken_package_quick_filters(path):
    """Keep an unreadable filters file instead of letting a later Save overwrite it."""
    backup = f"{path}.broken-{time.strftime('%Y%m%d-%H%M%S')}"
    try:
        os.replace(path, backup)
        logging.warning("Quick filter file was unreadable; kept it as %s", backup)
    except OSError:
        logging.exception("Could not back up broken quick filter file %s", path)


def _load_package_quick_filters(platform):
    """Saved filters for one platform, or its defaults when nothing is saved.

    The user's list is never overwritten by an update.  Defaults added in a
    newer release are appended once (tracked in seen_defaults), so a default
    the user deleted does not come back.
    """
    path = _package_quick_filters_path(platform)
    defaults = _default_package_quick_filters(platform)
    try:
        with open(path, "r", encoding="utf-8") as handle:
            stored = json.load(handle)
    except FileNotFoundError:
        return defaults, True
    except (OSError, ValueError):
        _backup_broken_package_quick_filters(path)
        return defaults, True
    filters, error = _validate_package_quick_filters(stored.get("filters") if isinstance(stored, dict) else None, platform)
    if error:
        _backup_broken_package_quick_filters(path)
        return defaults, True
    seen = stored.get("seen_defaults")
    needs_write = not isinstance(seen, list)
    if needs_write:
        # Saved before seen_defaults existed: every current default counts as
        # seen, and that is recorded now so a later release can tell its new
        # defaults apart.
        seen = [item["id"] for item in defaults]
    seen = {str(item) for item in seen}
    existing_ids = {item["id"] for item in filters}
    new_defaults = [item for item in defaults if item["id"] not in seen and item["id"] not in existing_ids]
    if new_defaults and len(filters) + len(new_defaults) <= PACKAGE_QUICK_FILTER_MAX:
        filters = filters + new_defaults
        seen |= {item["id"] for item in new_defaults}
        needs_write = True
    if needs_write:
        try:
            _write_package_quick_filters(platform, filters, seen)
        except OSError:
            logging.exception("Could not save the quick filter list")
    return filters, False


def _package_quick_filter_sql(filter_id, platform):
    """SQL clause for one saved quick filter, or (None, None) when unknown.

    Same rules as the live view.  instr() keeps the text literal (LIKE would
    treat _ and % as wildcards).
    """
    filters, _ = _load_package_quick_filters(platform)
    item = next((f for f in filters if f["id"] == filter_id), None)
    if item is None:
        return None, None
    whole_log = ("raw_log", "message", "tag", "device_name")
    checks = [(item["tag"], ("tag",)), (item["message1"], whole_log), (item["message2"], whole_log), (item["message3"], whole_log)]
    clause, params = "", []
    for text, columns in checks:
        if not text:
            continue
        # Never case sensitive.
        parts = [f"instr(lower(coalesce({col}, '')), ?) > 0" for col in columns]
        params.extend([text.lower()] * len(columns))
        clause += " AND (" + " OR ".join(parts) + ")"
    return clause, params


@app.get('/api/package-quick-filters')
def package_quick_filters_get_api():
    platform = _quick_filter_platform(request.args.get("platform"))
    filters, is_default = _load_package_quick_filters(platform)
    return jsonify({"ok": True, "platform": platform, "filters": filters, "is_default": is_default,
                    "defaults": _default_package_quick_filters(platform)})


@app.post('/api/package-quick-filters')
def package_quick_filters_save_api():
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict) or payload.get("platform") not in PACKAGE_QUICK_FILTER_PLATFORMS:
        return jsonify({"ok": False, "error": "invalid_platform"}), 400
    platform = payload["platform"]
    path = _package_quick_filters_path(platform)
    if payload.get("reset") is True:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
        filters, is_default = _load_package_quick_filters(platform)
        return jsonify({"ok": True, "platform": platform, "filters": filters, "is_default": is_default})
    filters, error = _validate_package_quick_filters(payload.get("filters"), platform)
    if error:
        return jsonify({"ok": False, "error": error}), 400
    _write_package_quick_filters(platform, filters, [item["id"] for item in _default_package_quick_filters(platform)])
    return jsonify({"ok": True, "platform": platform, "filters": filters, "is_default": False})


PACKAGE_QUICK_FILTER_EXPORT_KIND = "eventinspector_package_quick_filters"


@app.post('/api/package-quick-filters/export')
def package_quick_filters_export_api():
    """Write a filter list to a JSON file the user picks (backup / share)."""
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict) or payload.get("platform") not in PACKAGE_QUICK_FILTER_PLATFORMS:
        return jsonify({"ok": False, "error": "invalid_platform"}), 400
    platform = payload["platform"]
    filters, error = _validate_package_quick_filters(payload.get("filters"), platform)
    if error:
        return jsonify({"ok": False, "error": error}), 400
    filename = f"quick_filters_{platform}_{time.strftime('%Y%m%d_%H%M%S')}.json"
    document = {"kind": PACKAGE_QUICK_FILTER_EXPORT_KIND, "version": 2, "platform": platform, "filters": filters}
    if payload.get("download") is True:
        # Opened in a browser (no app window): the page saves the file itself
        # so the browser can ask where to put it.
        return jsonify({"ok": True, "filename": filename, "count": len(filters),
                        "content": json.dumps(document, ensure_ascii=False, indent=2)})
    export_dir = os.path.join(_user_data_dir(), "exports")
    os.makedirs(export_dir, exist_ok=True)
    path = os.path.join(export_dir, filename)
    try:
        if webview and getattr(webview, "windows", None):
            selected = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG,
                directory=export_dir,
                save_filename=filename,
                file_types=("JSON files (*.json)", "All files (*.*)"),
            )
            if not selected:
                return jsonify({"ok": False, "cancelled": True})
            path = selected[0] if isinstance(selected, (list, tuple)) else selected
    except Exception:
        logging.exception("Quick filter save dialog failed")
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(document, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)
    return jsonify({"ok": True, "path": path, "count": len(filters)})


@app.post('/api/package-quick-filters/import')
def package_quick_filters_import_api():
    """Validate an exported filter file; the editor shows it and the user saves."""
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict) or payload.get("platform") not in PACKAGE_QUICK_FILTER_PLATFORMS:
        return jsonify({"ok": False, "error": "invalid_platform"}), 400
    data = payload.get("data")
    if not isinstance(data, dict) or not isinstance(data.get("filters"), list):
        return jsonify({"ok": False, "error": "not_a_quick_filter_file"}), 400
    file_platform = data.get("platform")
    if file_platform in PACKAGE_QUICK_FILTER_PLATFORMS and file_platform != payload["platform"]:
        return jsonify({"ok": False, "error": "wrong_platform", "file_platform": file_platform}), 400
    filters, error = _validate_package_quick_filters(data["filters"], payload["platform"])
    if error:
        return jsonify({"ok": False, "error": error}), 400
    return jsonify({"ok": True, "platform": payload["platform"], "filters": filters})


# INDIEZ logo shown in the footer, embedded so it ships inside the remote-update payload
INDIEZ_LOGO_DATA_URI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAUcAAABSCAYAAADHJRb3AAAdk0lEQVR42u2de1hU1733v2vPjRkYGK7DHUEFlCDKeIlGTYh3iBrr0ZijpzFtGm1jLPGtJk2Pnpya9iRpq4mJPo3v07w2jR6rTRMbxURJlHhHkKsGoyAiiKDMDLdhLnvv9f6BGqMDzJ4ZLpr1eR6fBJhZs/Zas7/7t9b6XQilFAwGg8H4PhwbAgaDwWDiyGAwGEwcGQwG44ETx8bGxqlsP5TBYDBxvItt27YdPHv27GE2RQwGg4njTWpra5/fsmUL/vznPz/KpojBYDBxvMmHH374/uXLl7F7925UVFTsYtPEYDD6GjJQ9vUopcjPz6d79+7F3/72N1y+fBkAMHPmTMyfPx9ZWVkLIyIidrMpYzAYPxhxtFgseP3112lubi4uXLgAs9n8nWnLcYiJiUFSUhJ++ctfIjMzk7BpYzAYvY18IHRCqVRi9uzZsFqtqK+v/544AkBsbCzmzJmD1NTUFWzKGAzGD2pZLYoirl+/PuvKlSvbs7OzA48dOwYAeOuttzBnzpztUVFRy/z8/NrZlDEYjB+UON7Jp59+SufNm4ekpCTk5+fr/f39G9lUMRiM+2ZZTR3XqWg5B4BAFjDZa3uBM2fOJGPGjKHPPPMMmDAyGIz7xnKkjiYDX/9Bgdh2ElS0gdOkQJnwplcPSnJzc+m4ceOCtFqtiU0Tg8EY8JajaL1M7dW/AbVfAyh/65de71hGRgaRyWRshhgMxsAXRypYYL/0Kqi9rtc7xoSRwWDgfomQEW7spn0hjLecwhkMBuM+EEcKwbi/zzp28uRJarPZ2AwxGIwBKo6UB0QLRPs1Ktob+qRTdrsdq1evxo4dO5j5yGAwMLD2HKkAypseFi0VJ8S2ElD7FQB9o1V5eXn0+PHjaGlpwVNPPQWNRsNmisFgDABxFO0Q20qpo+EDdPox9h02mw1//OMfQSlFWVkZ/vWvf9FFixaxeGoGg9Hfy2oKseUUtdf8d58LI9C515ibm3v757fffhvt7e0KNlUMBqNfxZHar1FH3SZQobXPO9PR0YENGzZAFMXbv8vPz8eBAwfs7PSawWD067Kab/oMIn+9TzvR1tbme/HixYKCgoLkzz///B6Xnj/84Q/QaDQ0MTFxfWxs7DrmA8lgMNDX4YO2iqVUtFVLakTmmwblkHck7ws2NTUZzp07V3D69Gns3bsXR48ehcPhcPrasLAwZGZmYvr06UhNTc1LTEx8TKlUshlkMBh9I4726nUUoh2uO2tbAcECVeJWl8XRaDSmnjhxovT48eP4/PPPUVxc/L2ldHdotVo8/vjjmD59OkaPHm0fPXq0iuNYhVkGg9HLy2pFxPOQ5hieA75pL6jDmEoUQWWuvMtqtaYeOnQIO3fuRF2dtIib1tZW5OTkoL6+Hv7+/kqDwcBmkcFgDKx8jmJ7OXXU/A6ivR7KmDWQBblWwoBSivPnz+86cuTIgmPHjuHEiRO4cOFCjyGDGo0GBoMBEyZMwMSJE/HII48kBwYGnmfTyGAwBo448uYE++XXK4W2AgAA5zsCqsEbCIhcchnWEydOvH/8+HEcOnQIJSUlTkUxIyMDkyZNwvjx4zF69GjCHMMZDMaAFEf+2l+po+GvAG7uFRIFVAl/Auc3wi2HbaPRmJqfn1+6dOlSNDR8P0zx6aefxpo1aypSUlKGKRTM5ZHBYAxQcRQ7Kqm9ag0o3/S938sDJkMR9xoBcf+AJDs7m77zzju3f/b19UVeXl6DwWAI7+3BoIIF1HKOim0lkOkyQNQJD0RkTmNj49TDhw8fNBqNbmU7kslkGDduXGlqamoax3EwGo2pX375ZWlTU5Ok9gghCAgIwJQpU6aFhYXlwrU95sAvv/zSWF9f79Jn6HQ6TJs2bXxISMhJURRRVVX17pEjR1ZYrd7JOUoIAcdx8PHxgb+/PzQaDfz8/BAUFLQ7KChoa0BAQK5KpXKpLYfDgaKiIltxcbFSEAS3+hMSEoJp06YN1ul0VQ6HA8XFxa0lJSV+lFJwHAdCyO0+3/3/d/6uq5/vvGZnf3PlvxzHIT4+fmF4ePhuQjpvKZPJlHTw4MGKpqam/hVAQpCQkIBJkyYRtVrteZkE/vque4QRAITWAnBtRVSmNbgtKj//+c+3f/jhh4tNps4E4FlZWUhLS+tFYaSg1stUaD0D0VIOaqvt/Oe4AUXsyw/EEzAkJCQ3JSUl980335x67px7UU8JCQkjbgmhTqcrGz58uNvtXbt27eCzzz47WKfTVbkgINHFxcXYu3evS22PGzcOkydPTgZwkuM4xMXFvVhdXb3ilVde6e1hXnDzHxYuXIgJEyYgLi5uRURExGa53PltplAoMGzYsPDy8nLj1q1b3frQqVOn4tFHH00AUKVQKJCcnBxbVlZm3LJly4D5/o0ZMwavvvpq2C1hBAC1Wn2+o6MDf/nLX/q1b7GxsXjttdcKnT3QJJt41HKRis1Hnf9NtIBv/BuoYHG7s0OHDl2ycOFCAJ0lW1966SV09eXyzEpsVwgtR6ijdiO11/4R/PUdEMyHIXZ8CypaILQcAbVWPRBhORzHYdiwYdNWrlxpio2N9Vp7v/rVrwqHDh0q+f3bt2/Hnj17Ki0WS69fu0KhwKRJk8i6dev6bLx37dqF7OxszJs3772tW7fShoaGWV1Z2Fqt1vSjH/1o8I9//GO4W7XTm+31Bunp6QgODt6Mu8oxp6amXgkLC+vXvs2ePRtDhgwZ7cwdULI48sZ9oGJ7NyfYZyHc+IR6cuOtXLkyx8/PD5mZmUhPTyfetBJFywXKN3xIHdXr7I6r70MwfQGxvRzUceO7/VMAVGgD37QfDwocx2HkyJFBv/jFL7zW3kMPPTR69erVVxISEiS//91338VXX31F7XZ7r1+7SqXClClTlMuXL+/zcf/ggw+wevXqnKKiohqe57vaCqh66qmnMmfOnOnWPHizPW8TFBSEtLQ0u4+Pzz39jomJWTJy5Mh+69uoUaPw2GOPre/qcJeTZm21+QrNh3uqvgX+xj8gtpe6LZBJSUlZCxcuxKpVq+CNKBgqtCsEUy51VP+WOq78Dvz1jyG0nelcPotdJ9QVmr8C7YX6OP2FXC6HwWBInj9/vtcEMi0tLTY7OxvuWACbNm1Cfn4+7W6/jRDiFfX09fV1PP7442+mpaX1+bifO3cO2dnZMd09DEJDQ/dnZmZ6LT2ft9tzl7S0NAwaNCjLmYgHBAR8nZ6e3m99mzNnDqKiotZ5pUyC2JrfRnmzC/lxTXDUvQNqb6Tu1o957bXXVowdO9Yjq1G0fEsd9f+X2iuz7Y76P4NvOQrRWg0qNLuUm5I6jBBb8x+ojBcBAQHnp0yZYg8KCoK3av2kp6eTJUuWwA0PBWzevBlnz54t6CpCilLqtRjR6OjoV2bNmtUv426xWLB582aUlpYanV0rx3FISUmJmT17ttceXN5sDx7sN+p0utxu9lxzvbHVA+nlnzFu3Lhp3Rlf0mrImI+4nPBW7LgER+1boEKHu1/kza6e+t194iyYDlB71SvUXv0b8Dc+gdhxsXPZTB2Sl+GC6TAeqDRMHIdBgwZlSY0s6u5U2sfHB4899tjijIwMyf25cOEC3n//fUN1dfUbvZ15SaFQYPjw4TnubAN4g7q6Ouzbty/QaDQ+7Ozvfn5+taNGjfJoz9HT9uDlw47hw4fndeV+RwhBZGTkir6OctNoNMjMzERISEiud2rIUB5ia4GUaYPQVgS+5n8kxWrfOXCSrcS6zdT+7XPUXrcJQlsBqOM6IFp6FHTCqcD5DOliD7XwuxK0Dwh+fn658fHxXm0zNDR0xyOPPOLWe0+dOoVt27a9fO3atQXoZbeNkJCQjf0ljgBw7NgxVFVV5XVlPUZGRnpNvL3dnjtWY3R09E+6u5e1Wu35vhbwJ598Eg899JC+p5wMnOu+jd9SKrZJNePAtx6Dvfq/aG/UtqZCu4Jv2kttF1ZQW+UvwTd9CtFeDwhtLgkapxoERcRyqBI/hDJuLQDi5JnQAtF68YFaWisUCoSEhHj1YSWXyzFkyJA8d5dIOTk5+Pvf/77LbDYn9Mae4x1WbmlwcHC/jb3RaER5ebnSWfE4Qgi0Wm1OYGCgRwcynrQHLx94aLXaqp6+N4mJie8NHz68T/o0dOhQzJgxI1er1TbCW3WrRUuFezVkqACh9SRslf+HKuN/N5jIe/Zt67EvbeVUMO2D0Pw1qNgBUOpy3wjnA1nAo5AFZoLze4iAkNvPCE4ZScV7Ss9SiG1nwamTH6iltU6n87pVFhAQsDs0NPTRmpoat9r46KOPEBgYWLlgwYLb4aHe3HO8+WBolLrfmpmZiYiIiNs/8zyPmpoaHDp0yK0+VFdXo6OjI1WtVt+TqEWtVhdKEe+esllJbW/ixIlYs2bN4sjIyB3dvY7nebS3t0ebTKalFy9eXH8r5SDuOIhJSkpa70ru1bCwsPWjRo1acafPrEajwZIlS2AwGBqio6OfDQgI2K9SqW4/pB0OB3Jzc+natWsljX1WVhYGDx48zZWVqcviSDsueuhCcxa2b5+vVMauA+c3nLhutFKAUlChNVowHbgiNO2FaLss5bYFCAdOFQ9ZcBZkuqlBRK41OX2lJgVwUpebWisfqGX1nREO3qwjrlKpzvv7+3vUt3fffRfBwcF0+vTppDfydUq99oSEBCxatChn2LBhWXffUHa7HZcuXfpo586diz/77DOX2zSbzbBarakAypz0z+HNFHyEEEdv+AnL5XIEBATUBgQEvD5o0KDXx40bF7h//37jW2+9BQAwGAzQ6/XrXPQkaExLS8P27dsBAHPnzsWSJUvWx8XFretqLGpra3/7z3/+U/Iyf/Lkyavudivy3HK0XvZ4QKmjEfaqlyAP+3cqD/23IHAaE4isC0HkAeqAaKmgQtM+CC1Hu3W7uWe3gFOAyHwh858IWeBMcL7De3xUcJohEMwH7u2NrQ4Ml25EuzduxE2bNkGn09Hx48cP6PBNpVKJpKSkJc8880zF1atX1xcWFsLVciA8z+sfpLnXarWm2bNnE1EU6ZYtW5Ceng5XRUgmk2HIkCGrDAbDhujoaCxbtqzb8FKLxYJDhw6tLSoqktTHuXPnIjIycqPnpVnvsaNveGefkDrgaPgrBHOeUR7+E3B+I/VEpmkEkQOUBxWtvpQ3tYnNX0MwHoSkrOScEoRTg1PGgwuaBlnA5C6tRKc3tzKyS1FnuCaOUg/SunPx8ff3L4iMjFwx0K87KCho69ChQ10WR57nvb5dMBBQq9WYNWtWDCHkSnx8fKYUCzgkJGTj1KlTN4wZM2Ztd8IoiiIqKipq9uzZA6mRMGPGjBkt5eHt+rJaaPHqQIq2atgv/xdkviMaZKELwKkHQ+y4AMH0FcTWk5Kdr4k8CDLd5E4rUZPs1h1K5MF9cu0MuOzis3jx4hNdRZa4/4Cm3n4oSF66Ukr7LL2Uq1n24R0/2tq5c+cSqdmz1Go1srKySE/WptlsTt23b1+MlCTZQUFBmDVrlj0wMLAQHtetdpacQeiNSBEKob0EQnspiFwHypvcaoVThkERmQ0uYIJHZguR+XZ1VM/Uqh84deoUTp06xQYCnh/AQUKm/YaGhvVw/fS/TKfTld35Ga4up+/eC+4pmofneZw+fbpUqtU4Z84cDBs2LFzqXq7cxcdcL08fdVsYQRSQhz/nsTDeVMc+sTQYjIFKSUkJfvaznyUA2O7K69944w1kZGT0yd5wbW3tf3766aeS3pOcnIypU6fu9vPzM3lcQ6YL/xcQTg4qOnp/BIgchNMAhOs8gBGt3brpcKo4yHRTvDM5XUTQEI4l2GXcv/TWstpgMCAxMXFVX5RKvnkIs/706dOSrcb4+PiF7uyFu75RwvkBoqn3rp4owPnEg6hiwSkjAU4J6jBDtNWA2qq7PBThtOldWnyStZFvdv4HmT+7wxgM3OsaExoauhF9IO7nz5+vlGo13qwz9aw7YciSxJHIg9xf+vZomKrBBUyGPPgJcJrh5E6xo7w5QWzOq+SbPoPoxNeSyL3n/U8d17q8dgaDgbtTkbW5KzxSMJvNqTk5OQlSK5U+8cQT0Ov129z9XM51N5feSsZNINNlQBH5gp7zTSV3W4FErquSBc8liojnQRR6J4J2w2s9ETuq+/jaGYz7k7Fjx3aZisyb8DyPgoKC0k8++UTS++bNm4f09PQRnvjdui6Oqt5JK8Qp9ZDrfwwi9+/WmZDTjiXyoDlOQglLvJYYQrR84/yzfeLZ3dALhIWFYeLEiWwg7kPS09MREBDwNXo/k9HLUoUxKioKs2bNuqLT6co80iaXX6hJ7B1x1I4FUYaT26fWtitUbC2gYssJKraXUwhtt/1rZMEzxhNOdVfkziVQW43nx8l8Sxi1Og+RJOqh7G7oBZKTk7F8+fLdy5YtY4OBgeHKAxeTN6SkpGzv7UqgFosFhw8ffkPqIUxWVhaSkpIGe3rdctfFcfjNl3vXIZdobgqPYAFv/IKKlmJQewMgOgC5Pzif+DaZLgOcbyohipCTkAcC9jv3BgUIrUWQ+3iWlkloK2mgQpvTIeLUiewOQ685De+eP3/++ubm5tKdO3eyAcH9cRATERGxyhvRUN0dwnz77bffSPVpTE1NRUZGxlaNRuOxa43rBzLKUML5xFLRWgUv5x0DQMHf+Jjy1/8JKpjuWTaLHZVQRK+knCqeOPO5FDsqPO6GYMrtYkk9CEQZQtgtgd7c3C9bvHjxwpaWll05OTlsQNC/ZQ1WrFhRpdfr13aTLPakKym/4PkhTLLUDE+zZ89GXFzcMm8It1zKClzmPwHeFkfReBAO+zUqGL+4RxhvJc0V20vA128FUcVRZ6+hds/midrqqNh+xunfZAET3KlDxpBIeHj47qVLl75pMplePnHiBBsQ9FsCCej1+rU9pSzryerzZEl76xBGatadjIwMPPzww4t7yuYkiuLt+tveyQQOgNM9BkK86/ApWM6Bv/Fpj25CQutp8E174CyrOKGCR33gm/aBCq1OQpoU4HQZ7I5B36RRGzRo0CvLly/P7avEp/dr/PJA7pvVasXRo0ep0WhMdbeNq1evviR1OQ107jXq9fpuRd1ms+HYsWP0xo0bk712IHNziUmIehi8HTroUpZwKgJdpSyT+XlkNQqmL5xfr28aOFUsW1Kj7w4OkpOTp61YscIeFRXVq0IM7x4cPFxfXy/pOjmOa79fD2S6s/jy8/Pp22+/jZqaml3uhN12dHTg0KFDG6TG1C9YsAAjR44c3F20jiAIKCoqohs3bsSlS5cO9vTQ4KSG9smD5w48q0MZ4bYwOxq2g/JNzv0vQ570WvQNAy7n9hs1apRq5cqV/V5W1BVaWlrCDhw4kCMlK7iPjw9kMpkJD5DlyPM8zpw5Y3vvvfdQU1OD0tLSZGelIFw4hCmRajVGRUVhxowZFf7+/lXdCWNxcfG1d955BzU1NSgpKVFarVZv7Tne3oNTksZBdiolz2JvPxk17rnaiM3HqNhFHW5OnQSZ1sCsRvRPjZuJEyeS7Oxs+vvf/75f+1JVVYWioqJMk8lEb1mcoijCbDbj4sWLyMnJgdFolNRmcHAw1Gp14YNgOdrtdtTW1r5x+PDhl7ds2XL79wUFBZg6deq/S9m7vHkIM0LqIczcuXORmJg4zNl12mw21NXVvZGXl/fy5s2bv9e/GTNmvBQbG7vRa+IIma9DEbYY9iu/GyjSCM5PevUyam+kjvqtoKLFqdUoD3sa4NRMqfoJlUqF6dOn+5nN5rY7b7r+YONG74YPx8TEwFn9mIHA0aNHodfrt+t0uu2uJCX+4osvYLHcew8dP34clZWV28PDw3e4Is43Lc/Sjz/+WHKfz549i23bttE7t0va29vxzTffoKSkxOl7CgsLUVlZuSEqKmpjV0txt2JrZLrJRGbaT4W2M/2/DFMngSj0Euu42uCo/RNE25UutjBHQub/MLMa0e8lZNvnzZuXbDabK3bs2PFAXJNGo0FKSgr6IibZXdwRKGcUFxdj1KhRvn5+fu2uHMK4+7l5eXnIy8tzq38GgyHM3995dJ579jangjx8Ge6OVukXcQyeBRAJl0FFOGrfpkJbodNUaIRTQRH5C4BTMXUaAAQGBp5ftGjR4pkzZz4Q1zNz5kwkJyeP6KtDkv48ST9z5gwaGxvXunIIc/jw4Q1SI2E8pbCwsNvEvm7PEOc7lMjD/qNfv2hEEQZZ4HQiSRivbqK8+WCX8djysGdA1IOZ1TiAiIyM3LF8+fJl97tADh06FLNnzz7haczv/UJZWRkqKipe7q7MhSiKuHDhQoHUdGTeoKKiAhUVFc87HA7viiPAQa5/msj8DP02+IrwnwCcj6vJGmG/8iblb+zpUhhl2nGQ658izOl74BEdHb31xRdfzJw/f/59ew3PPvsskpOTJ/Sl1djfFBUVobW1Namrvzc3Nyfl5OQY3K117qX+pbq85yg0H6OA4OL+YwbE9jJQau/b5bR2DGSBM4hLJRgcplTH5d+WCu3F3ewUxEMR959BAHPdGajo9fr9L7zwgl6v1zf09yENJGYfWrVqFSZNmkR6O1nDQKOwsBC1tbX/T6fTTbjbv5TneRQWFlb84x//6Pf+BQYGjr67f07F0VH7J1DeOECHm4Aow6GIeWU0unXmpYBoh9hWRO21fwJ1XO+2PWX8/4DItCYmQQMbf3//xkWLFhGdTtfvbj5wMeHqkiVL3ktISHjxh2Qx3qKmpgbl5eXjExMT7zmEqq+vf8GdSBh4NyUaysvLDUlJSff0j7u/krsScIpQKOP+G0QRXNi1KHZAtNVSR+1Garv0arfCyCn1ne2pwtk+Yx9GlcDDE9+5c+eSjz76aPcTTzwx4MZq/Pjx+PWvf42PP/74lbVr15IhQ4ZIEkZvv7a/RbmgoAAmk2nu3WGGeXl57w2EOPrCwkKYzeZZLi2rOc0wiJZzA+zulINTRkIR8zI4TSJxtqdIhfZo8KYrvOkghKZ/OY2XvrPSIKeKhiL2N+DUiT84YQwODj7x/PPPjxcEwaWoldDQ0NKeBFKpVFYZDAbExMS47AytUqkq3I2kSU5OXvjqq69i4cKF+2prazNNJhMIId22yXEc4uPj8dxzz3mlqiQhBEqlEnq9HnFxcXlxcXHztFr3VyAKheLKyJEjERHhWtRXSEgIlEpllbfa6w2USiVaW1sz9Xr9nlvfoaampqWiKOKnP/0pBoJPbWtra2ZYWNj+O7/jxNkXRGw7Q21Va7yWYdvziqlacJpkyKNe/C7WmYoAtYIKHQaIlgJqrQHfchRi8xE4z8v4fVckTpMEReQqcOpBzGJkMBhw0XIcQTj1UNpV2YC+Xc6pwGnHQx72byAAqLWKUtEOKrSAWmtArRchtpVDtNd1W8L1O/efYMj8xkIe8bMRRBFUxr4CDAbDqVZ0tbQQmo9Sx5U3erbCen01HQzOJ6Hz1JnaAKENlDeC8i0uieF3zt0+IKp4yIJmQh78BGEJJRgMhlviCErhuLqZ8sbPuk4V5nLWHD0I8em07vp6qc4pwSljwPmmQh7yJIgPW0YzGAxPMoETAnn4M0FUaDeKzXldJGhwRRijoAhbBMgDIbYchdhxAaKt1mPBdaUWNvGJA6dOAhcwCTK/UcxaZDAYXrAcbx8CmxP4xv+tFFuOQrTVSYntA6cZDnlQFmSB0wgIARVtoG1FVGg7A7GjEtRWe9PNRvSWIoLIQ0F8YiFTJ4LTjkZnLWw5m2kGg+Fdcew8GLZBbP6aCi0nIXacB7XVo8sIGk4NzmcQOE0KZIGPg9MMJ85inEXrJUotFRA7LoA6roHaG+/YSxRdFmAi04LIA0GUoSDKKHA+CSC+KeB84gjAVtAMBqMXxfG7HIjXqdheBmqthGhvABWaAWoHKAGRaQB5IDhVFIg6CTLfFOJK3DMVbaD2Okpt9aCORlBHEyhvBsR2QLR1HsJQoTOShSgBmRrgNCCcBkTuD8h1IPIQEFUEiDKKEE7JZpXBYPStOH7P4dpxg1L+pjiCgMjUgDwomch15z2y2KgIKrQkQGivhGjtjNm+WUCLcKrOBLScZgRkmjLiatIJBoPB6BNxZDAYjAcclpuLwWAwmDgyGAwGE0cGg8Fwm/8PpmclDVTXuiMAAAAASUVORK5CYII="


# --- SERVER ROUTE (QUAN TRỌNG) ---
@app.route('/')
def index():
    ui_settings = _load_ui_settings()
    ui_zoom = ui_settings["zoom"]
    return render_template_string(
        HTML_TEMPLATE,
        ui_zoom=ui_zoom,
        ui_hide_top=ui_settings["hide_top"],
        indiez_logo=INDIEZ_LOGO_DATA_URI,
        package_quick_filters={p: _load_package_quick_filters(p)[0] for p in PACKAGE_QUICK_FILTER_PLATFORMS},
        package_quick_filter_defaults={p: _default_package_quick_filters(p) for p in PACKAGE_QUICK_FILTER_PLATFORMS},
        native_file_dialog=bool(webview and getattr(webview, "windows", None)),
        ui_zoom_px=round(16 * ui_zoom, 2),
        ui_default_zoom=UI_DEFAULT_ZOOM,
        default_event_names=sorted(event_specific_params.keys()),
        current_profile_name=active_profile_name,
        profile_revision=profile_revision,
        default_ad_event_data=_default_ad_event_payload(),
        sdk_check_presets=_load_sdk_check_presets(),
    )


@app.get('/api/sdk-check-presets')
def get_sdk_check_presets():
    refresh_requested = request.args.get("refresh", "").strip().lower() in {"1", "true", "yes"}
    presets, source = _fetch_sdk_check_presets(force_remote=refresh_requested)
    return jsonify({
        'ok': bool(presets),
        'source': source,
        'presets': presets,
        'error': '' if presets else 'sdk_check_presets_unavailable',
    })


@app.get('/api/package-log-presets')
def get_package_log_presets():
    refresh_requested = request.args.get("refresh", "").strip().lower() in {"1", "true", "yes"}
    presets, source = _fetch_package_log_presets(force_remote=refresh_requested)
    return jsonify({
        'ok': bool(presets),
        'source': source,
        'presets': presets,
        'error': '' if presets else 'package_log_presets_unavailable',
    })


@app.post('/api/services-checker/start')
def start_services_checker():
    try:
        url = _start_services_checker()
        return jsonify({'ok': True, 'url': url})
    except Exception as exc:
        logging.exception("Services Checker start failed")
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.get('/api/services-checker/host')
def get_services_checker_host():
    return jsonify({'ok': True, **_services_checker_host_status()})


@app.post('/api/services-checker/host')
def replace_services_checker_host():
    upload = request.files.get('host_file')
    if not upload or not upload.filename:
        return jsonify({'ok': False, 'error': 'host_file_required'}), 400
    try:
        target = _save_services_checker_host(upload)
        return jsonify({
            'ok': True,
            'filename': os.path.basename(target),
            **_services_checker_host_status(),
        })
    except Exception as exc:
        logging.exception("Services Checker host file save failed")
        return jsonify({'ok': False, 'error': str(exc)}), 400


@app.post('/api/services-checker/reload')
def reload_services_checker():
    try:
        url, restarted = _reload_services_checker()
        return jsonify({'ok': True, 'url': url, 'restarted': restarted})
    except Exception as exc:
        logging.exception("Services Checker reload failed")
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.get('/api/profiles')
def get_profiles():
    return jsonify(_profile_payload())


@app.post('/api/profiles/select')
def select_profile():
    data = request.get_json(silent=True) or {}
    profile_name = data.get('profile_name', '')
    if not profile_name:
        return jsonify({'ok': False, 'error': 'profile_name_required'}), 400
    try:
        if not _set_active_profile(profile_name):
            return jsonify({'ok': False, 'error': 'profile_not_found'}), 404
        socketio.emit('update_default_ad_events', _default_ad_event_payload())
        _emit_profile_runtime_refresh()
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    return jsonify({'ok': True, **_profile_payload()})


@app.post('/api/profiles/reload')
def reload_profile():
    if active_profile_name:
        _set_active_profile(active_profile_name)
    else:
        _set_active_profile()
    socketio.emit('update_default_ad_events', _default_ad_event_payload())
    _emit_profile_runtime_refresh()
    return jsonify({'ok': True, **_profile_payload()})


@app.post('/api/profiles/import')
def import_profile():
    upload = request.files.get('profile_file')
    if not upload or not upload.filename:
        return jsonify({'ok': False, 'error': 'profile_file_required'}), 400
    try:
        filename = _sanitize_profile_filename(upload.filename)
        target = os.path.join(PROFILE_DIR, filename)
        os.makedirs(PROFILE_DIR, exist_ok=True)
        upload.save(target)
        _set_active_profile(filename)
        socketio.emit('update_default_ad_events', _default_ad_event_payload())
        _emit_profile_runtime_refresh()
        return jsonify({'ok': True, **_profile_payload()})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400


@app.post('/api/ipa/inspect')
def inspect_ipa():
    upload = request.files.get('ipa_file')
    if not upload or not upload.filename:
        return jsonify({'ok': False, 'error': 'ipa_file_required'}), 400
    tmp_path = None
    try:
        filename = _sanitize_ipa_filename(upload.filename)
        preset_name = str(request.form.get('preset_name') or '').strip()
        os.makedirs(PROFILE_DIR, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix="eventinspector_ipa_", suffix=".ipa", delete=False) as tmp:
            tmp_path = tmp.name
            upload.save(tmp_path)
        ipa_info = _extract_brightsdk_version_from_ipa(tmp_path)
        presets = _load_sdk_check_presets()
        bright_compare = _compare_bright_skad_preset(
            ipa_info.get('brightsdk_version'),
            ipa_info.get('skadnetwork_identifiers') or [],
            preset_name=preset_name,
            preset=presets.get(preset_name),
        )
        return jsonify({
            'ok': True,
            'filename': filename,
            'app_name': ipa_info.get('app_name') or 'Not Found',
            'brightsdk_version': ipa_info.get('brightsdk_version') or 'Not Found',
            'skadnetwork_identifiers': ipa_info.get('skadnetwork_identifiers') or [],
            'bright_compare': bright_compare,
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


@app.post('/api/profiles/open-folder')
def open_profile_folder():
    try:
        os.makedirs(PROFILE_DIR, exist_ok=True)
        if sys.platform == 'darwin':
            subprocess.Popen(['open', PROFILE_DIR])
        elif sys.platform.startswith('win'):
            os.startfile(PROFILE_DIR)
        else:
            subprocess.Popen(['xdg-open', PROFILE_DIR])
        return jsonify({'ok': True, 'profile_dir': PROFILE_DIR})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

@app.post('/check_update')
def check_update():
    try:
        remote_update = _load_remote_update_module()
    except Exception as e:
        return jsonify({'ok': False, 'status': 'error', 'error': f'updater_unavailable: {e}'})
    try:
        result = remote_update.check_for_updates()
        return jsonify(result)
    except Exception as e:
        return jsonify({'ok': False, 'status': 'error', 'error': str(e)})


@app.post('/clear_update_cache')
def clear_update_cache():
    try:
        remote_update = _load_remote_update_module()
    except Exception as e:
        return jsonify({'ok': False, 'error': f'updater_unavailable: {e}'})
    try:
        remote_update.clear_update_cache(include_all_channels=True)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.post('/restart_app')
def restart_app():
    cmd = os.getenv('EVENTINSPECTOR_RESTART_CMD')
    args = os.getenv('EVENTINSPECTOR_RESTART_ARGS', '')
    if not cmd:
        # ``python -c 'import Log_checker; run_server(...)'`` is the supported
        # local test mode and has no desktop process to relaunch. Reset the
        # runtime instead of returning an error that the UI cannot act on.
        _clear_default_ad_event_history()
        _reset_runtime_for_platform_switch()
        return jsonify({'ok': True, 'mode': 'local_runtime_reset'})
    _clear_default_ad_event_history()
    argv = [cmd]
    if args:
        argv.append(args)
    try:
        subprocess.Popen(argv, creationflags=creation_flags)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    os._exit(0)


def _package_session_platform_resolver():
    """Return f(package_id, stored_platform) -> 'android' | 'ios'.

    Sessions saved before the platform column existed are matched against
    the project list (read only); otherwise an id with a dot is an Android
    package and one without is an iOS bundle search name.
    """
    android_ids, ios_ids = set(), set()
    for project in _load_package_log_presets() or []:
        if isinstance(project, dict):
            if (project.get("android") or {}).get("id"):
                android_ids.add(str(project["android"]["id"]).strip().lower())
            if (project.get("ios") or {}).get("id"):
                ios_ids.add(str(project["ios"]["id"]).strip().lower())

    def resolve(package_id, stored=None):
        if stored in ("android", "ios"):
            return stored
        key = str(package_id or "").strip().lower()
        if key in ios_ids and key not in android_ids:
            return "ios"
        if key in android_ids:
            return "android"
        return "android" if "." in key else "ios"

    return resolve


def _package_session_platform(conn, session_id):
    row = conn.execute("SELECT package_id, platform FROM package_log_sessions WHERE id = ?", (session_id,)).fetchone()
    if not row:
        return None
    return _package_session_platform_resolver()(row["package_id"], row["platform"])


@app.get('/api/package-log/sessions')
def package_log_sessions_api():
    """Saved sessions, newest first, with paging.

    Row counts come from package_log_sessions.row_count; only sessions not
    backfilled yet fall back to counting their rows.
    """
    limit = max(1, min(request.args.get('limit', default=500, type=int) or 500, 1000))
    offset = max(request.args.get('offset', default=0, type=int) or 0, 0)
    platform_filter = (request.args.get('platform') or '').strip().lower()
    if platform_filter not in ('', 'android', 'ios'):
        return jsonify({"ok": False, "error": "invalid_platform"}), 400
    resolve_platform = _package_session_platform_resolver()
    conn = _get_package_db_connection()
    try:
        # The platform of older sessions is resolved in Python, so filter
        # on ids first, then page and count rows only for that page.
        candidates = [
            row["id"]
            for row in conn.execute("SELECT id, package_id, platform FROM package_log_sessions ORDER BY id DESC")
            if not platform_filter or resolve_platform(row["package_id"], row["platform"]) == platform_filter
        ]
        total = len(candidates)
        page_ids = candidates[offset:offset + limit]
        rows = []
        if page_ids:
            marks = ",".join("?" * len(page_ids))
            rows = conn.execute(
                f"""
                SELECT s.id, s.package_id, s.started_at, s.ended_at, s.status, s.platform,
                       CASE WHEN s.row_count IS NULL
                            THEN (SELECT COUNT(*) FROM package_log_entries e WHERE e.session_id = s.id)
                            ELSE s.row_count END AS row_count
                FROM package_log_sessions s
                WHERE s.id IN ({marks})
                ORDER BY s.id DESC
                """,
                page_ids,
            ).fetchall()
        sessions = []
        for row in rows:
            started_label = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(row["started_at"])) if row["started_at"] else ""
            sessions.append({
                "id": row["id"],
                "package_id": row["package_id"],
                "started_at": row["started_at"],
                "ended_at": row["ended_at"],
                "started_label": started_label,
                "status": row["status"],
                "platform": resolve_platform(row["package_id"], row["platform"]),
                "row_count": row["row_count"] or 0,
            })
        return jsonify({
            "ok": True,
            "sessions": sessions,
            "total": total,
            "offset": offset,
            "limit": limit,
            "has_more": offset + len(sessions) < total,
        })
    finally:
        conn.close()


@app.get('/api/package-log/maintenance')
def package_log_maintenance_api():
    """Size, retention preview and Compact state for the History view."""
    cfg = _load_package_log_retention()
    conn = _get_package_db_connection()
    try:
        plan = _package_log_retention_plan(conn, cfg)
        has_seq_index = bool(conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'index' AND name = 'idx_package_log_entries_session_seq'"
        ).fetchone())
    finally:
        conn.close()
    with lock:
        capturing = bool(target_package_name)
    oldest = min((item["started_at"] for item in plan["sessions"]), default=None)
    return jsonify({
        "ok": True,
        "usage": plan["usage"],
        "total_sessions": plan["total_sessions"],
        "total_rows": plan["total_rows"],
        "policy": {key: cfg[key] for key in ("days", "max_bytes", "confirmed")},
        "cleanup": {
            "sessions": len(plan["sessions"]),
            "rows": plan["rows"],
            "estimated_bytes": plan["estimated_bytes"],
            "by_age": sum(1 for item in plan["sessions"] if item["reason"] == "age"),
            "by_size": sum(1 for item in plan["sessions"] if item["reason"] == "size"),
            "oldest_label": time.strftime("%Y-%m-%d", time.localtime(oldest)) if oldest else "",
        },
        "has_seq_index": has_seq_index,
        "capturing": capturing,
    })


@app.post('/api/package-log/retention/apply')
def package_log_retention_apply_api():
    """Delete sessions per the retention policy.  Requires explicit confirm."""
    payload = request.get_json(silent=True) or {}
    if not payload.get("confirm"):
        return jsonify({"ok": False, "error": "confirmation_required"}), 400
    result = _apply_package_log_retention(confirmed_by_user=True)
    return jsonify(result), (200 if result.get("ok") else 400)


@app.post('/api/package-log/compact')
def package_log_compact_api():
    """Rebuild the history file (VACUUM) and add the scroll-back index.

    VACUUM locks the database for its whole run, so it is refused while a
    capture is writing.
    """
    payload = request.get_json(silent=True) or {}
    if not payload.get("confirm"):
        return jsonify({"ok": False, "error": "confirmation_required"}), 400
    with lock:
        if target_package_name:
            return jsonify({"ok": False, "error": "capture_running"}), 409
    conn = _get_package_db_connection()
    try:
        before = _package_log_db_usage(conn)
        started = time.time()
        conn.execute(PACKAGE_LOG_SEQ_INDEX_SQL)
        conn.commit()
        conn.execute("VACUUM")
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        after = _package_log_db_usage(conn)
    finally:
        conn.close()
    return jsonify({"ok": True, "before": before, "after": after, "seconds": round(time.time() - started, 1)})


@app.get('/api/package-log/logs')
def package_log_rows_api():
    session_id = request.args.get('session_id', type=int)
    keyword1 = (request.args.get('q1') or request.args.get('q') or '').strip().lower()
    keyword2 = (request.args.get('q2') or '').strip().lower()
    keyword3 = (request.args.get('q4') or '').strip().lower()
    tag_keyword = (request.args.get('q3') or '').strip().lower()
    quick_filter_id = (request.args.get('qf') or '').strip()
    quick_platform = (request.args.get('qp') or '').strip()
    offset = max(request.args.get('offset', default=0, type=int) or 0, 0)
    limit = request.args.get('limit', default=2000, type=int) or 2000
    limit = max(1, min(limit, 5000))
    if not session_id:
        return jsonify({"ok": False, "error": "session_id_required"}), 400
    conn = _get_package_db_connection()
    try:
        base_where = """
            FROM package_log_entries
            WHERE session_id = ?
        """
        params = [session_id]

        def search_clause():
            return """
              AND (
                lower(coalesce(raw_log, '')) LIKE ?
                OR lower(coalesce(message, '')) LIKE ?
                OR lower(coalesce(tag, '')) LIKE ?
                OR lower(coalesce(device_name, '')) LIKE ?
              )
            """

        def tag_search_clause():
            return """
              AND lower(coalesce(tag, '')) LIKE ?
            """

        if keyword1:
            base_where += search_clause()
            like1 = f"%{keyword1}%"
            params.extend([like1, like1, like1, like1])
        if keyword2:
            base_where += search_clause()
            like2 = f"%{keyword2}%"
            params.extend([like2, like2, like2, like2])
        if keyword3:
            base_where += search_clause()
            like3 = f"%{keyword3}%"
            params.extend([like3, like3, like3, like3])
        if tag_keyword:
            base_where += tag_search_clause()
            params.append(f"%{tag_keyword}%")
        if quick_filter_id:
            quick_clause, quick_params = _package_quick_filter_sql(quick_filter_id, _package_session_platform(conn, session_id) or quick_platform)
            if quick_clause is None:
                return jsonify({"ok": False, "error": "unknown_quick_filter"}), 400
            base_where += quick_clause
            params.extend(quick_params)

        total_rows = conn.execute(f"SELECT COUNT(*) AS c {base_where}", params).fetchone()["c"] or 0
        sql = f"""
            SELECT time_display, device_id, device_name, tag, message, raw_log
            {base_where}
            ORDER BY id DESC
            LIMIT ? OFFSET ?
        """
        query_params = params + [limit, offset]
        rows = conn.execute(sql, query_params).fetchall()
        rows = list(reversed(rows))
        loaded_count = offset + len(rows)
        return jsonify({
            "ok": True,
            "session_id": session_id,
            "rows": [dict(r) for r in rows],
            "offset": offset,
            "limit": limit,
            "total_rows": total_rows,
            "has_more": loaded_count < total_rows,
        })
    finally:
        conn.close()


@app.post('/api/package-log/clear')
def package_log_clear_api():
    payload = request.get_json(silent=True) or {}
    if not payload.get("clear_all"):
        return jsonify({"ok": False, "error": "confirmation_required"}), 400
    conn = _get_package_db_connection()
    try:
        session_count = conn.execute("SELECT COUNT(*) AS c FROM package_log_sessions").fetchone()["c"] or 0
        row_count = conn.execute("SELECT COUNT(*) AS c FROM package_log_entries").fetchone()["c"] or 0
        conn.execute("DELETE FROM package_log_entries")
        conn.execute("DELETE FROM package_log_sessions")
        conn.commit()
        try:
            conn.execute("VACUUM")
        except Exception:
            pass
        return jsonify({
            "ok": True,
            "deleted_sessions": session_count,
            "deleted_rows": row_count,
        })
    finally:
        conn.close()


@app.get('/api/package-log/raw')
def package_log_raw_api():
    """Return the raw record of one live row (the UI receives message only)."""
    seq = request.args.get('seq', type=int)
    if not seq:
        return jsonify({"ok": False, "error": "seq_required"}), 400
    with lock:
        for row in reversed(package_log_cache):
            row_seq = row.get('seq', 0)
            if row_seq == seq:
                return jsonify({"ok": True, "seq": seq, "raw_log": row.get('log', ''), "source": "live"})
            if row_seq < seq:
                break
        session_id = active_package_log_session_id
    session_id = request.args.get('session_id', type=int) or session_id
    if not session_id:
        return jsonify({"ok": False, "error": "row_not_found"}), 404
    conn = _get_package_db_connection()
    try:
        row = conn.execute(
            "SELECT raw_log FROM package_log_entries WHERE session_id = ? AND seq = ? LIMIT 1",
            (session_id, seq),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return jsonify({"ok": False, "error": "row_not_found"}), 404
    return jsonify({"ok": True, "seq": seq, "raw_log": row["raw_log"] or "", "source": "history"})


@app.get('/api/package-log/older')
def package_log_older_rows_api():
    """Rows of the live session older than `before_seq`, oldest first.

    The UI calls this when the user scrolls above its in-memory store, so
    rows evicted from the live view stay reachable without leaving the tab.
    Within one session the writer inserts rows in seq order, so walking the
    (session_id, id) index backwards returns them newest first.
    """
    before_seq = request.args.get('before_seq', type=int)
    if not before_seq or before_seq <= 1:
        return jsonify({"ok": True, "rows": [], "has_more": False})
    limit = max(1, min(request.args.get('limit', default=2000, type=int) or 2000, 5000))
    session_id = request.args.get('session_id', type=int)
    if not session_id:
        with lock:
            session_id = active_package_log_session_id
    if not session_id:
        return jsonify({"ok": True, "rows": [], "has_more": False})
    conn = _get_package_db_connection()
    try:
        rows = conn.execute(
            """
            SELECT seq, time_display, device_id, device_name, level, tag, message, is_error
            FROM package_log_entries
            WHERE session_id = ? AND seq IS NOT NULL AND seq < ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (session_id, before_seq, limit + 1),
        ).fetchall()
    finally:
        conn.close()
    has_more = len(rows) > limit
    rows = rows[:limit]
    rows.reverse()
    return jsonify({
        "ok": True,
        "session_id": session_id,
        "rows": [
            {
                "seq": row["seq"],
                "device_id": row["device_id"] or "",
                "device_name": row["device_name"] or "",
                "time": row["time_display"] or "",
                "time_display": row["time_display"] or "",
                "level": row["level"] or "",
                "tag": row["tag"] or "",
                "message": row["message"] or "",
                "pid": "",
                "tid": "",
                "is_error": bool(row["is_error"]),
            }
            for row in rows
        ],
        "has_more": has_more,
    })


@app.post('/api/package-log/export')
def package_log_export_api():
    payload = request.get_json(silent=True) or {}
    try:
        session_id = int(payload.get('session_id'))
    except Exception:
        return jsonify({"ok": False, "error": "session_id_required"}), 400

    keyword1 = str(payload.get('q1') or '').strip().lower()
    keyword2 = str(payload.get('q2') or '').strip().lower()
    keyword3 = str(payload.get('q4') or '').strip().lower()
    tag_keyword = str(payload.get('q3') or '').strip().lower()
    quick_filter_id = str(payload.get('qf') or '').strip()
    quick_platform = str(payload.get('qp') or '').strip()
    filtered_only = bool(payload.get('filtered_only', True))
    conn = _get_package_db_connection()
    try:
        session = conn.execute(
            "SELECT id, package_id, started_at FROM package_log_sessions WHERE id = ?",
            (session_id,),
        ).fetchone()
        if not session:
            return jsonify({"ok": False, "error": "session_not_found"}), 404

        sql = """
            SELECT raw_log, message
            FROM package_log_entries
            WHERE session_id = ?
        """
        params = [session_id]

        def search_clause():
            return """
              AND (
                lower(coalesce(raw_log, '')) LIKE ?
                OR lower(coalesce(message, '')) LIKE ?
                OR lower(coalesce(tag, '')) LIKE ?
                OR lower(coalesce(device_name, '')) LIKE ?
              )
            """

        def tag_search_clause():
            return """
              AND lower(coalesce(tag, '')) LIKE ?
            """

        if keyword1:
            like1 = f"%{keyword1}%"
            sql += search_clause()
            params.extend([like1, like1, like1, like1])
        if keyword2:
            like2 = f"%{keyword2}%"
            sql += search_clause()
            params.extend([like2, like2, like2, like2])
        if keyword3:
            like3 = f"%{keyword3}%"
            sql += search_clause()
            params.extend([like3, like3, like3, like3])
        if tag_keyword:
            sql += tag_search_clause()
            params.append(f"%{tag_keyword}%")
        if quick_filter_id and filtered_only:
            quick_clause, quick_params = _package_quick_filter_sql(quick_filter_id, _package_session_platform(conn, session_id) or quick_platform)
            if quick_clause is None:
                return jsonify({"ok": False, "error": "unknown_quick_filter"}), 400
            sql += quick_clause
            params.extend(quick_params)

        sql += " ORDER BY id ASC"
        rows = conn.execute(sql, params).fetchall()

        export_dir = os.path.join(_package_history_dir(), "exports")
        os.makedirs(export_dir, exist_ok=True)
        started_at = session["started_at"] or time.time()
        started_label = time.strftime("%Y%m%d_%H%M%S", time.localtime(started_at))
        package_stub = re.sub(r'[^A-Za-z0-9._-]+', '_', session["package_id"] or f"session_{session_id}")
        suffix = "filtered" if filtered_only and (keyword1 or keyword2 or keyword3 or tag_keyword or quick_filter_id) else "full"
        filename = f"package_log_{session_id}_{package_stub}_{started_label}_{suffix}.log"
        path = os.path.join(export_dir, filename)

        chosen_path = None
        try:
            if webview and getattr(webview, "windows", None):
                selected = webview.windows[0].create_file_dialog(
                    webview.SAVE_DIALOG,
                    directory=export_dir,
                    save_filename=filename,
                    file_types=("Log files (*.log)", "All files (*.*)"),
                )
                if selected:
                    chosen_path = selected[0] if isinstance(selected, (list, tuple)) else selected
        except Exception:
            logging.exception("Package log save dialog failed")

        if chosen_path:
            path = chosen_path
        elif webview and getattr(webview, "windows", None):
            return jsonify({"ok": False, "cancelled": True, "session_id": session_id})

        with open(path, "w", encoding="utf-8") as f:
            for row in rows:
                text = row["raw_log"] or row["message"] or ""
                f.write(text.rstrip("\n"))
                f.write("\n")

        return jsonify({"ok": True, "path": path, "row_count": len(rows), "session_id": session_id})
    finally:
        conn.close()

# --- BACKEND LOGIC ---

def _normalize_load_ads_provider(value):
    provider = str(value or "").strip().lower()
    return provider or "Unknown"


def send_to_sheet(device_name, ad_source, ad_format, raw_log, log_type, provider="Unknown"):
    """Gửi data lên Google Sheet nếu đang bật Recording cho loại Log cụ thể"""
    state = recording_states.get(log_type)
    if state and state["is_recording"] and state["current_sheet"]:
        payload = {
            "sheet_name": state["current_sheet"], 
            "device_name": device_name, 
            "provider": _normalize_load_ads_provider(provider),
            "ad_source": ad_source, 
            "ad_format": ad_format, 
            "raw_log": raw_log
        }
        threading.Thread(
            target=lambda: requests.post(
                G_SHEET_URL,
                json=payload,
                timeout=(RECORD_SHEET_CONNECT_TIMEOUT_SEC, RECORD_SHEET_READ_TIMEOUT_SEC),
            ),
            daemon=True,
        ).start()


def _clean_record_sheet_name(response_text, fallback_name):
    """Keep a failed Apps Script HTML response out of the recording button label."""
    candidate = str(response_text or "").strip()
    if not candidate:
        return fallback_name

    # Some deployments return JSON, while the current sheet script returns a plain name.
    if candidate.startswith("{"):
        try:
            payload = json.loads(candidate)
        except (TypeError, ValueError):
            payload = None
        if isinstance(payload, dict):
            candidate = str(
                payload.get("sheet_name")
                or payload.get("name")
                or payload.get("result")
                or ""
            ).strip()

    # HTML means the deployment returned an error/login page, not a sheet name.
    if (
        not candidate
        or len(candidate) > 200
        or "<!doctype" in candidate[:200].lower()
        or "<html" in candidate[:200].lower()
        or "<head" in candidate[:200].lower()
    ):
        return fallback_name
    return candidate


def _create_record_sheet_in_background(tab_name, requested_name, request_id):
    try:
        response = requests.post(
            G_SHEET_URL,
            json={"action": "create_or_get_sheet", "sheet_name": requested_name},
            timeout=(RECORD_SHEET_CONNECT_TIMEOUT_SEC, RECORD_SHEET_READ_TIMEOUT_SEC),
        )
        if response.status_code >= 400:
            raise requests.RequestException(f"HTTP {response.status_code}")
        resolved_name = _clean_record_sheet_name(response.text, requested_name)
    except Exception as exc:
        resolved_name = requested_name
        logging.warning("Load Ads sheet setup failed; recording locally with %r: %s", requested_name, exc)

    state = recording_states.get(tab_name)
    if not state or not state.get("is_recording") or state.get("record_request_id") != request_id:
        return
    state["current_sheet"] = resolved_name
    socketio.emit("record_status", {
        "tab_name": tab_name,
        "is_recording": True,
        "current_sheet": resolved_name,
    })

def process_load_ads_unity_log(line, device_id):
    """Xử lý TrackingService ad_impression cho Tab Load Ads.

    Unity Android records historically used ``eventName`` + ``e`` while the
    iOS UnityFramework record uses ``EventName`` + ``params``. Both are the
    same TrackingService record format, so parse them through the shared event
    parser instead of making Load Ads depend on one JSON spelling.
    """
    # CHỈ XỬ LÝ NẾU ĐANG GHI (RECORDING)
    if not recording_states["LoadAds"]["is_recording"]:
        return

    if not TRACKING_SERVICE_EVENT_PATTERN.search(str(line or "")):
        return

    try:
        event_name, params, _json_string = find_and_parse_event(line)
        if event_name != "ad_impression" or not isinstance(params, dict):
            return

        # Keep the existing Load Ads columns while accepting both old Unity
        # names and the newer iOS TrackingService names. This is only field
        # extraction after the record has been parsed; record matching itself
        # is shared and does not depend on the ad parameter set.
        src = (
            params.get("ad_source")
            or params.get("ad_network")
            or params.get("network_name")
            or params.get("ad_platform")
        )
        fmt = params.get("ad_format")
        if params.get("mediation_ad_unit_name") == "MREC":
            fmt = "MREC"

        if not src or not fmt:
            return

        d_name = get_device_name(device_id)
        snapshot = None
        with lock:
            dedup_key = (device_id, str(src), str(fmt), "tracking")
            if dedup_key not in unique_load_ads and _accept_exact_record("load_ads", device_id, line, src, fmt, "tracking"):
                unique_load_ads.add(dedup_key)
                _live_table_append("load_ads", {
                    "device_id": device_id,
                    "device_name": d_name,
                    "ad_source": src,
                    "ad_format": fmt,
                    "raw_log": line.strip()
                })
                snapshot = True
        if snapshot is not None:
            send_to_sheet(d_name, src, fmt, line.strip(), "LoadAds")
    except Exception:
        pass


def _first_ios_max_field(pattern, text):
    match = pattern.search(text or "")
    if not match:
        return ""
    return next((value.strip() for value in match.groups() if value and value.strip()), "")


def _split_ios_max_log_records(raw_text):
    """Split a relay chunk at iOS syslog headers before MAX parses it.

    tidevice can return more than one physical syslog record in a single
    stdout chunk.  MAX viewability payloads are multiline, so parsing the
    whole chunk as one record leaks the following SDK/postback lines into the
    preceding Load Ads row.
    """
    text = str(raw_text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not text.strip():
        return []

    starts = [
        match.start()
        for match in re.finditer(rf"(?m)^{IOS_LOG_TIMESTAMP_PATTERN}\s+", text)
    ]
    if len(starts) <= 1:
        return [text.strip()]

    records = []
    if starts[0] > 0 and text[:starts[0]].strip():
        records.append(text[:starts[0]].strip())
    for index, start in enumerate(starts):
        end = starts[index + 1] if index + 1 < len(starts) else len(text)
        record = text[start:end].strip()
        if record:
            records.append(record)
    return records


def _ios_max_viewability_block_end(text):
    """Return the end of the MAX parameter dictionary, if it is complete."""
    closing_line = re.search(r"(?m)^\s*}\s*(?:\n|$)", str(text or ""))
    if closing_line:
        return closing_line.end()
    stripped = str(text or "").rstrip()
    if stripped.endswith("}"):
        return len(stripped)
    return None


def _process_ios_max_viewability_record(raw_record, device_id):
    raw_record = str(raw_record or "").strip()
    if not raw_record:
        return

    record_starts = bool(IOS_LOG_RECORD_START_PATTERN.match(raw_record))
    with lock:
        previous = incomplete_ios_max_load_ads_logs.get(device_id)
        starts_viewability = MAX_LOAD_ADS_VIEWABILITY_KEYWORD in raw_record

        # A timestamped record is a new syslog record, never a continuation
        # of the prior MAX dictionary. Drop an incomplete stale block instead
        # of attaching the next SDK/postback record to it.
        if record_starts and previous and not starts_viewability:
            previous = None
            incomplete_ios_max_load_ads_logs.pop(device_id, None)

        if starts_viewability:
            buffered = raw_record
        else:
            if not previous:
                return
            buffered = f"{previous}\n{raw_record}"

        block_end = _ios_max_viewability_block_end(buffered)
        if block_end is None:
            if len(buffered) <= 50000:
                incomplete_ios_max_load_ads_logs[device_id] = buffered
            else:
                incomplete_ios_max_load_ads_logs.pop(device_id, None)
            return
        buffered = buffered[:block_end].rstrip()
        incomplete_ios_max_load_ads_logs.pop(device_id, None)

    ad_format = _first_ios_max_field(MAX_LOAD_ADS_IOS_FORMAT_PATTERN, buffered).upper()
    ad_network = _first_ios_max_field(MAX_LOAD_ADS_IOS_NETWORK_PATTERN, buffered)
    if not ad_format or not ad_network:
        return

    d_name = get_ios_device_name(device_id)
    provider = "MAX"
    dedup_key = (device_id, provider.lower(), ad_network.casefold(), ad_format, "max_viewability")
    row = {
        "device_id": device_id,
        "device_name": d_name,
        "provider": provider,
        "ad_source": ad_network,
        "ad_network": ad_network,
        "ad_format": ad_format,
        "adapter_class": _first_ios_max_field(MAX_LOAD_ADS_IOS_ADAPTER_CLASS_PATTERN, buffered),
        "adapter_version": _first_ios_max_field(MAX_LOAD_ADS_IOS_ADAPTER_VERSION_PATTERN, buffered),
        "raw_log": buffered,
    }
    snapshot = None
    with lock:
        if dedup_key in unique_load_ads_ext:
            return
        if not _accept_exact_record("load_ads_ext", device_id, buffered, provider, ad_network, ad_format, "max_viewability"):
            return
        unique_load_ads_ext.add(dedup_key)
        _live_table_append("load_ads_ext", row)
        snapshot = True
    send_to_sheet(d_name, ad_network, ad_format, buffered, "LoadAdsExt", provider)


def _process_ios_max_viewability_log(line, device_id):
    """Collect MAX viewability dictionaries without absorbing later records."""
    for record in _split_ios_max_log_records(line):
        _process_ios_max_viewability_record(record, device_id)


def _process_ios_max_delegate_log(line, device_id):
    """Collect one-line MAX delegate callbacks used by iOS Load Ads."""
    raw_line = str(line or "").strip()
    if not raw_line or not any(keyword in raw_line for keyword in MAX_LOAD_ADS_IOS_DELEGATE_KEYWORDS):
        return

    format_match = MAX_LOAD_ADS_FORMAT_PATTERN.search(raw_line)
    network_match = MAX_LOAD_ADS_IOS_DELEGATE_NETWORK_PATTERN.search(raw_line)
    if not format_match or not network_match:
        return

    ad_format = format_match.group(1).strip().upper()
    ad_network = next(
        (value.strip() for value in network_match.groups() if value and value.strip()),
        "",
    )
    if not ad_format or not ad_network:
        return

    d_name = get_ios_device_name(device_id)
    provider = "MAX"
    dedup_key = (device_id, provider.lower(), ad_network.casefold(), ad_format, "max_delegate")
    snapshot = None
    with lock:
        if dedup_key in unique_load_ads_ext:
            return
        if not _accept_exact_record("load_ads_ext", device_id, raw_line, provider, ad_network, ad_format, "max_delegate"):
            return
        unique_load_ads_ext.add(dedup_key)
        _live_table_append("load_ads_ext", {
            "device_id": device_id,
            "device_name": d_name,
            "provider": provider,
            "ad_source": ad_network,
            "ad_network": ad_network,
            "ad_format": ad_format,
            "raw_log": raw_line,
        })
        snapshot = True
    send_to_sheet(d_name, ad_network, ad_format, raw_line, "LoadAdsExt", provider)


def process_load_ads_max_log(line, device_id):
    """Process MAX Load Ads callbacks for Android and multiline viewability logs for iOS."""
    # The visible Load Ads tab is backed by LoadAdsExt.  LoadAds is the
    # legacy hidden Unity table, so MAX rows must use the same state, cache,
    # socket event, and sheet target as the table users actually record.
    if not recording_states["LoadAdsExt"]["is_recording"]:
        if active_platform == "ios":
            with lock:
                incomplete_ios_max_load_ads_logs.pop(device_id, None)
        return

    if active_platform == "ios":
        _process_ios_max_delegate_log(line, device_id)
        _process_ios_max_viewability_log(line, device_id)
        return

    if active_platform != "android":
        return
    if MAX_LOAD_ADS_REVENUE_KEYWORD not in str(line or ""):
        return

    raw_line = str(line or "").strip()
    format_match = MAX_LOAD_ADS_FORMAT_PATTERN.search(raw_line)
    network_match = MAX_LOAD_ADS_NETWORK_PATTERN.search(raw_line)
    if not format_match or not network_match:
        return

    ad_format = format_match.group(1).strip().upper()
    ad_network = network_match.group(2).strip()
    if not ad_format or not ad_network:
        return

    d_name = get_device_name(device_id)
    provider = "MAX"
    dedup_key = (device_id, provider.lower(), ad_network.casefold(), ad_format, "max_revenue")
    snapshot = None
    with lock:
        if dedup_key in unique_load_ads_ext:
            return
        if not _accept_exact_record("load_ads_ext", device_id, raw_line, provider, ad_network, ad_format, "max_revenue"):
            return
        unique_load_ads_ext.add(dedup_key)
        _live_table_append("load_ads_ext", {
            "device_id": device_id,
            "device_name": d_name,
            "provider": provider,
            "ad_source": ad_network,
            "ad_network": ad_network,
            "ad_format": ad_format,
            "raw_log": raw_line,
        })
        snapshot = True
    send_to_sheet(d_name, ad_network, ad_format, raw_line, "LoadAdsExt", provider)

def process_load_ads_ext_log(line, device_id):
    """Xử lý log cho Tab 2: Load Ads Ext (AppMetrica AdRevenue)"""
    # CHỈ XỬ LÝ NẾU ĐANG GHI (RECORDING)
    if not recording_states["LoadAdsExt"]["is_recording"]:
        return

    ios_payload_part = ""
    if active_platform == "ios":
        for keyword in IOS_LOAD_ADS_EXT_ADREVENUE_KEYWORDS:
            if keyword in line:
                ios_payload_part = line.split(keyword, 1)[1].strip()
                break
    if ios_payload_part:
        try:
            json_str, raw_log = _buffer_ios_load_ads_ext_payload(device_id, ios_payload_part, line)
            if not json_str:
                return
            data = _loads_adrevenue_json_payload(json_str)
            ad_revenue = data.get("adRevenue") if isinstance(data.get("adRevenue"), dict) else {}
            if not ad_revenue and isinstance(data, dict) and any(key in data for key in ("AdRevenueValue", "Currency", "AdNetwork", "Payload")):
                ad_revenue = data
            payload = ad_revenue.get("Payload") if isinstance(ad_revenue.get("Payload"), dict) else {}

            provider = _normalize_load_ads_provider(
                payload.get("ad_platform") or ad_revenue.get("ad_platform")
            )
            ad_network = ad_revenue.get("AdNetwork") or payload.get("ad_network")
            fmt = payload.get("ad_format") or ad_revenue.get("AdFormat") or ad_revenue.get("AdType")
            if fmt and str(fmt).strip().lower() == "mrec":
                fmt = "MREC"

            if ad_network and fmt:
                d_name = get_ios_device_name(device_id)
                snapshot = None
                with lock:
                    dedup_key = (device_id, provider, ad_network, fmt, "ios_metrica")
                    if dedup_key not in unique_load_ads_ext and _accept_exact_record("load_ads_ext", device_id, raw_log or line, provider, ad_network, fmt, "ios_metrica"):
                        unique_load_ads_ext.add(dedup_key)
                        _live_table_append("load_ads_ext", {
                            "device_id": device_id,
                            "device_name": d_name,
                            "provider": provider,
                            "ad_network": ad_network,
                            "ad_format": fmt,
                            "raw_log": raw_log or line.strip()
                        })
                        snapshot = True
                if snapshot is not None:
                    send_to_sheet(d_name, ad_network, fmt, raw_log or line.strip(), "LoadAdsExt", provider)
            return
        except:
            pass

    match = LOAD_ADS_EXT_ADREVENUE_PATTERN.search(line)
    if match:
        try:
            parsed = parse_appmetrica_adrevenue_text(f"AdRevenue{{{match.group(1)}}}") or {}
            payload = parsed.get("payload") if isinstance(parsed.get("payload"), dict) else {}

            provider = _normalize_load_ads_provider(
                payload.get("ad_platform") or parsed.get("ad_platform")
            )
            ad_network = parsed.get("adNetwork") or payload.get("ad_network")
            fmt = parsed.get("adType") or payload.get("ad_format")
            if fmt and str(fmt).strip().lower() == "mrec":
                fmt = "MREC"

            if ad_network and fmt:
                d_name = get_device_name(device_id)
                snapshot = None
                with lock:
                    dedup_key = (device_id, provider, ad_network, fmt, "metrica")
                    if dedup_key not in unique_load_ads_ext and _accept_exact_record("load_ads_ext", device_id, line, provider, ad_network, fmt, "metrica"):
                        unique_load_ads_ext.add(dedup_key)
                        _live_table_append("load_ads_ext", {
                            "device_id": device_id,
                            "device_name": d_name, 
                            "provider": provider,
                            "ad_network": ad_network,
                            "ad_format": fmt, 
                            "raw_log": line.strip()
                        })
                        snapshot = True
                if snapshot is not None:
                    # Gửi với type "LoadAdsExt"
                    send_to_sheet(d_name, ad_network, fmt, line.strip(), "LoadAdsExt", provider)
        except: pass

def _extract_event_json_candidate(text):
    """Extract one JSON object without counting braces inside string values."""
    raw_text = str(text or "")
    start = raw_text.find("{")
    if start >= 0:
        decoder = json.JSONDecoder()
        for candidate_text in (raw_text, _decode_ios_levelplay_json_text(raw_text)):
            candidate_start = candidate_text.find("{")
            if candidate_start < 0:
                continue
            try:
                parsed, end = decoder.raw_decode(candidate_text[candidate_start:])
                if isinstance(parsed, dict):
                    return candidate_text[candidate_start:candidate_start + end]
            except (TypeError, ValueError, json.JSONDecodeError):
                pass
    # Keep the existing Objective-C/plist-style fallback for callers that
    # receive a non-JSON diagnostic around the event payload.
    return extract_json_object_from_text(raw_text)


def _mapping_value_case_insensitive(data, *keys):
    if not isinstance(data, dict):
        return None
    normalized = {str(key).casefold(): value for key, value in data.items()}
    for key in keys:
        value = normalized.get(str(key).casefold())
        if value is not None:
            return value
    return None


def _parse_tracking_service_event(log_entry):
    """Parse both Android and iOS TrackingService event spellings."""
    text = str(log_entry or "")
    marker = TRACKING_SERVICE_EVENT_PATTERN.search(text)
    if not marker:
        return None, None, None

    after_keyword = text[marker.end():]
    json_str = _extract_event_json_candidate(after_keyword)
    candidates = []
    if json_str:
        candidates.extend([json_str, _decode_ios_levelplay_json_text(json_str)])
    raw_tail = after_keyword.strip()
    if raw_tail:
        candidates.extend([raw_tail, _decode_ios_levelplay_json_text(raw_tail)])

    data = None
    seen = set()
    for candidate in candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            parsed = json.loads(candidate)
        except (TypeError, ValueError, json.JSONDecodeError):
            continue
        if isinstance(parsed, dict):
            data = parsed
            break

    if isinstance(data, dict):
        event_name = _mapping_value_case_insensitive(
            data, "eventName", "event_name", "name", "event"
        )
        params = _mapping_value_case_insensitive(data, "e", "params", "parameters")
        if isinstance(params, str):
            try:
                params = json.loads(params)
            except (TypeError, ValueError, json.JSONDecodeError):
                params = {}
        if not isinstance(params, dict):
            params = {}
        if event_name:
            event_name = str(event_name).strip()
            if event_name:
                wrapped = {
                    "eventName": event_name,
                    "e": params,
                    "source": "firebase",
                }
                return event_name, params, json.dumps(wrapped, ensure_ascii=False)

    # If the payload is malformed, keep the event visible as a FAILED/INFO
    # record instead of silently dropping it from every event tab. A line that
    # starts an unfinished JSON object is left for the iOS stream buffer.
    if not json_str and "{" not in after_keyword:
        name_match = re.search(
            r"[\"']?(?:eventname|event_name|name)[\"']?\s*(?:=|:)\s*[\"']([^\"']+)[\"']",
            after_keyword,
            re.IGNORECASE,
        )
        if name_match:
            event_name = name_match.group(1).strip()
            wrapped = {"eventName": event_name, "e": {}, "source": "firebase"}
            return event_name, {}, json.dumps(wrapped, ensure_ascii=False)

    return None, None, None


def find_and_parse_event(log_entry):
    """Parse TrackingService and AppMetrica regular event records."""
    event_name, params, json_string = _parse_tracking_service_event(log_entry)
    if event_name:
        return event_name, params, json_string

    match = METRICA_REGULAR_EVENT_PATTERN.search(str(log_entry or ""))
    if match:
        try:
            event_name = match.group(1)
            params = json.loads(match.group(2))
            if not isinstance(params, dict):
                params = {}
            wrapped = {
                "eventName": event_name,
                "e": params,
                "source": "appmetrica",
            }
            return event_name, params, json.dumps(wrapped, ensure_ascii=False)
        except (TypeError, ValueError, json.JSONDecodeError):
            pass
    return None, None, None


def _tracking_service_payload_fragment(raw_line):
    """Return the payload after a TrackingService marker, if present.

    Some iOS syslog relays hard-wrap a long TrackingService JSON message and
    repeat the complete syslog prefix on the continuation line.  The second
    line can therefore look like ``...TrackingService->Track: ode_game...``
    even though it is not a new event.  Keeping only the text after the
    repeated marker lets the stream reassembler join the JSON without
    inserting a newline into a split key/string.
    """
    text = str(raw_line or "").rstrip("\r\n")
    marker = TRACKING_SERVICE_EVENT_PATTERN.search(text)
    if not marker:
        return None
    # The separator after the colon is not part of the JSON.  Remove only
    # leading transport whitespace; the continuation itself is joined
    # byte-for-byte afterwards so split keys/strings remain intact.
    return text[marker.end():].lstrip().rstrip()


def _tracking_buffer_expects_continuation(buffered):
    """Tell whether a buffered JSON payload can still accept a fragment."""
    payload = _tracking_service_payload_fragment(buffered)
    if not payload:
        return False

    in_string = False
    escaped = False
    for char in payload:
        if escaped:
            escaped = False
            continue
        if char == "\\" and in_string:
            escaped = True
            continue
        if char == '"':
            in_string = not in_string

    # A wrap can happen in a JSON string (the user's ``"m`` example), after
    # a delimiter, or immediately after an object/array opener.  In all of
    # those states a repeated marker is more likely a continuation than a
    # new event, even when its fragment starts with ``{``.
    return in_string or payload.rstrip().endswith((":", ",", "[", "{"))


def _process_ios_event_stream_line(device_id, raw_line):
    """Return completed event records while buffering split iOS JSON lines."""
    text = str(raw_line or "").rstrip("\r\n")
    if not text.strip():
        return []

    completed = []
    with lock:
        buffered = incomplete_ios_event_logs.get(device_id, "")

    # A long iOS TrackingService line may be emitted as multiple timestamped
    # lines, each with the full ``TrackingService->Track:`` prefix repeated.
    # Try the payload-only concatenation before treating the timestamp as a
    # new syslog record.  This is important when the wrap occurs in the
    # middle of a JSON key, e.g. ``"m`` + ``ode_game"``.
    if buffered:
        fragment = _tracking_service_payload_fragment(text)
        if fragment is not None:
            joined = f"{buffered}{fragment}"
            joined_parsed = find_and_parse_event(joined)
            if joined_parsed[0]:
                with lock:
                    incomplete_ios_event_logs.pop(device_id, None)
                completed.append((joined, joined_parsed))
                return completed

            # A repeated marker followed by a non-object fragment is a
            # continuation of the buffered JSON. Keep it without a newline;
            # a newline could corrupt a split key or string value.
            if fragment and (
                not fragment.lstrip().startswith("{")
                or _tracking_buffer_expects_continuation(buffered)
            ):
                if len(joined) <= 200000:
                    with lock:
                        incomplete_ios_event_logs[device_id] = joined
                else:
                    with lock:
                        incomplete_ios_event_logs.pop(device_id, None)
                return completed

    # A timestamped line starts a new syslog record. Finish the previous
    # TrackingService record before considering this line independently.
    if buffered and IOS_LOG_RECORD_START_PATTERN.match(text.strip()):
        parsed = find_and_parse_event(buffered)
        if parsed[0]:
            completed.append((buffered, parsed))
        buffered = ""
        with lock:
            incomplete_ios_event_logs.pop(device_id, None)

    combined = f"{buffered}\n{text}" if buffered else text
    parsed = find_and_parse_event(combined)
    if parsed[0]:
        with lock:
            incomplete_ios_event_logs.pop(device_id, None)
        completed.append((combined, parsed))
        return completed

    if TRACKING_SERVICE_EVENT_PATTERN.search(combined):
        if len(combined) <= 200000:
            with lock:
                incomplete_ios_event_logs[device_id] = combined
        else:
            with lock:
                incomplete_ios_event_logs.pop(device_id, None)
    return completed


def _safe_runtime_call(label, callback, *args):
    """Run one log consumer without allowing it to kill the device reader."""
    try:
        return callback(*args)
    except Exception as exc:
        print(f"WARNING: {label} failed: {exc}")
        return None


def _dispatch_event_record(raw_log, device_id, parsed=None):
    """Send one parsed event to every relevant tab independently."""
    if parsed is None:
        parsed = _safe_runtime_call("event parser", find_and_parse_event, raw_log)
    if not parsed:
        parsed = (None, None, None)
    event_name, params, json_string = parsed
    if event_name:
        # Each tab is isolated deliberately. A malformed profile/filter or a
        # transient formatter error must not stop the shared device stream.
        _safe_runtime_call(
            "callback/ad-event consumer",
            process_callback_and_ad_event_log,
            raw_log,
            device_id,
            event_name,
            params,
            json_string,
        )
        _safe_runtime_call(
            "default-ad-event consumer",
            _record_default_ad_event_hit,
            event_name,
            params,
            device_id,
        )
        _safe_runtime_call(
            "validator consumer",
            process_event_validator_log,
            event_name,
            params,
            json_string,
            raw_log,
            device_id,
        )
        _safe_runtime_call(
            "specific-event consumer",
            cache_specific_event_log,
            event_name,
            params,
            json_string,
            raw_log,
            device_id,
        )
    else:
        _safe_runtime_call("callback consumer", process_callback_and_ad_event_log, raw_log, device_id)


def _dispatch_android_event_record(raw_log, device_id, parsed=None):
    """Dispatch one complete Android record to Load Ads and event tabs."""
    # Load Ads/Unity historically parsed the physical adb line itself.  Keep
    # that behavior for normal records, but also feed it the reconstructed
    # record when logcat wrapped a long TrackingService JSON payload.
    _safe_runtime_call("Android Load Ads/Unity event", process_load_ads_unity_log, raw_log, device_id)
    _dispatch_event_record(raw_log, device_id, parsed)


def _process_android_event_stream_line(device_id, raw_line):
    """Reassemble wrapped Android TrackingService records before dispatch.

    ``adb logcat`` normally returns one message per line, but Unity can emit a
    long JSON message that the relay exposes as several physical lines.  The
    old per-line parser silently discarded every fragment.  Only records that
    contain the TrackingService marker are buffered; ordinary callbacks keep
    their existing one-line behavior.
    """
    text = str(raw_line or "").rstrip("\r\n")
    if not text.strip():
        return

    with lock:
        buffered = incomplete_android_event_logs.get(device_id, "")

    # A new timestamp starts a new logcat record.  Flush a completed previous
    # record before handling the new one; never merge unrelated log messages.
    if buffered and ANDROID_LOG_RECORD_START_PATTERN.match(text.strip()):
        parsed = find_and_parse_event(buffered)
        if parsed[0]:
            _dispatch_android_event_record(buffered, device_id, parsed)
        buffered = ""
        with lock:
            incomplete_android_event_logs.pop(device_id, None)

    combined = f"{buffered}\n{text}" if buffered else text
    parsed = find_and_parse_event(combined)
    if parsed[0]:
        with lock:
            incomplete_android_event_logs.pop(device_id, None)
        _dispatch_android_event_record(combined, device_id, parsed)
        return

    if TRACKING_SERVICE_EVENT_PATTERN.search(combined):
        if len(combined) <= 200000:
            with lock:
                incomplete_android_event_logs[device_id] = combined
        else:
            with lock:
                incomplete_android_event_logs.pop(device_id, None)
        return

    _dispatch_event_record(text, device_id)


def _flush_android_event_log_buffer(device_id):
    """Dispatch a final complete Android event before a reader is replaced."""
    with lock:
        buffered = incomplete_android_event_logs.pop(device_id, "")
    if not buffered:
        return
    parsed = _safe_runtime_call("Android final event parser", find_and_parse_event, buffered)
    if parsed and parsed[0]:
        _dispatch_android_event_record(buffered, device_id, parsed)


def _process_android_runtime_line(line, device_id):
    """Process one Android logcat line without coupling tab consumers."""
    _safe_runtime_call("Android Load Ads/MAX", process_load_ads_max_log, line, device_id)
    _safe_runtime_call("Android Load Ads Ext", process_load_ads_ext_log, line, device_id)
    _safe_runtime_call("Android SDK Check", _process_sdk_check_line, line, device_id)
    _safe_runtime_call("Android AdRevenue", process_adrevenue_log, line, device_id)
    _safe_runtime_call("Android Price Rotation", process_price_rotation_log, line, device_id)
    _safe_runtime_call("Android Installation ID", process_installation_id_log, line, device_id)
    _process_android_event_stream_line(device_id, line)


def _process_ios_runtime_line(raw_line, device_id):
    """Process one iOS physical line and complete buffered event records."""
    line = str(raw_line or "")
    if not line:
        return

    _safe_runtime_call("iOS Load Ads/MAX", process_load_ads_max_log, line, device_id)
    _safe_runtime_call("iOS Load Ads/Unity", process_load_ads_unity_log, line, device_id)
    _safe_runtime_call("iOS Load Ads Ext", process_load_ads_ext_log, line, device_id)
    _safe_runtime_call("iOS AdRevenue", process_adrevenue_log, line, device_id)
    _safe_runtime_call("iOS Price Rotation", process_price_rotation_log, line, device_id)
    _safe_runtime_call("iOS SDK Check", _process_sdk_check_line, line, device_id)

    with lock:
        had_buffer = bool(incomplete_ios_event_logs.get(device_id))
    is_record_start = bool(IOS_LOG_RECORD_START_PATTERN.match(line.strip()))
    is_tracking_line = bool(TRACKING_SERVICE_EVENT_PATTERN.search(line))
    completed = _safe_runtime_call(
        "iOS event stream parser",
        _process_ios_event_stream_line,
        device_id,
        line,
    ) or []

    for event_raw, parsed in completed:
        _dispatch_event_record(event_raw, device_id, parsed)

    # A non-TrackingService line is independent unless it was a continuation
    # of the buffered TrackingService JSON. This also keeps AppMetrica regular
    # events and ordinary callbacks flowing through the normal dispatcher.
    continuation = had_buffer and not is_record_start
    if not is_tracking_line and not continuation:
        current_parsed = _safe_runtime_call("iOS current-event parser", find_and_parse_event, line)
        current_was_completed = any(
            str(event_raw or "").strip() == line.strip()
            for event_raw, _parsed in completed
        )
        if not current_was_completed:
            _dispatch_event_record(line, device_id, current_parsed)


def _flush_ios_event_log_buffer(device_id):
    """Dispatch a final buffered iOS event when a reader is replaced/stopped."""
    with lock:
        buffered = incomplete_ios_event_logs.pop(device_id, "")
    if not buffered:
        return
    parsed = _safe_runtime_call("iOS final event parser", find_and_parse_event, buffered)
    if parsed and parsed[0]:
        _dispatch_event_record(buffered, device_id, parsed)

def process_callback_and_ad_event_log(log_entry, device_id, event_name=None, actual_params=None, json_string=None):
    global incomplete_impression_logs
    if is_paused: return

    # --- 0. Process iOS LevelPlay callbacks ---
    ios_callback_key = next((key for key in IOS_LEVELPLAY_CALLBACK_KEYS if key in log_entry), "")
    if ios_callback_key:
        display_name = IOS_LEVELPLAY_CALLBACK_KEYS.get(ios_callback_key, ios_callback_key)
        after_keyword = log_entry.split(ios_callback_key, 1)[1].strip()
        json_str = extract_json_object_from_text(after_keyword)
        details = "Callback fired"
        json_data_for_log = "{}"
        if json_str:
            data = _parse_callback_json_payload(json_str)
            if data is not None:
                fallback_format = ""
                if "[Ad,LevelPlay,Banner]" in ios_callback_key:
                    fallback_format = "banner"
                elif "[Ad,LevelPlay,Mrec" in ios_callback_key:
                    fallback_format = "mrec"
                if "_OnImpressionDataReady" in ios_callback_key or "_OnLevelPlayImpressionDataReadyEvent" in ios_callback_key:
                    display_name = _levelplay_impression_event_name(data, fallback_format)
                details = format_json_html(data)
                json_data_for_log = json.dumps(data, ensure_ascii=False)
            else:
                details = f'<div class="text-xs font-mono break-all text-red-600">JSON Parse Error</div><div class="text-xs font-mono break-all">{html.escape(json_str)}</div>'
        elif after_keyword:
            details = f'<div class="text-xs font-mono break-all">{html.escape(after_keyword.lstrip(":").strip())}</div>'
        _append_callback_ad_row({
            "device_id": device_id,
            "device_name": get_device_name(device_id),
            "type": "Callback Levelplay",
            "event_name": display_name,
            "details": details,
            "raw_log": log_entry.strip(),
            "json_data": json_data_for_log
        })
        return

    # --- 0b. Process Android LevelPlay impression data mapper callbacks ---
    if LEVELPLAY_IMPRESSION_DATA_MAPPER_KEYWORD in log_entry:
        after_keyword = log_entry.split(LEVELPLAY_IMPRESSION_DATA_MAPPER_KEYWORD, 1)[1].strip()
        json_str = extract_json_object_from_text(after_keyword)
        details_target = None
        json_data_for_log = "{}"
        details = "LevelPlay impression data"
        if json_str:
            try:
                data = _parse_callback_json_payload(json_str) or {}
                details_target = data.get("LevelPlayImpressionData", data)
                if not isinstance(details_target, (dict, list)):
                    details_target = data
                details = format_json_html(details_target)
                json_data_for_log = json.dumps(details_target, ensure_ascii=False)
            except Exception:
                details = f'<div class="text-xs font-mono break-all text-red-600">JSON Parse Error</div><div class="text-xs font-mono break-all">{html.escape(json_str)}</div>'
        elif after_keyword:
            details = f'<div class="text-xs font-mono break-all">{html.escape(after_keyword.lstrip(":").strip())}</div>'

        display_name = _levelplay_impression_data_event_name(details_target or {})
        _append_callback_ad_row({
            "device_id": device_id,
            "device_name": get_device_name(device_id),
            "type": "Callback Levelplay",
            "event_name": display_name,
            "details": details,
            "raw_log": log_entry.strip(),
            "json_data": json_data_for_log,
        })
        return

    # --- 0c. Process AscendX callbacks ---
    ascendx_event = _find_named_callback_event(log_entry, ASCENDX_CALLBACK_EVENTS)
    if ascendx_event:
        _record_named_provider_callback(log_entry, device_id, ascendx_event, "Callback Ascendx")
        return

    # --- 0d. Process CloudX callbacks ---
    cloudx_event = _find_named_callback_event(log_entry, CLOUDX_CALLBACK_EVENTS)
    if cloudx_event:
        _record_named_provider_callback(log_entry, device_id, cloudx_event, "Callback Cloudx")
        return

    # --- 0e. Process Gadsme callbacks ---
    if GADSME_SERVICE_KEYWORD in log_entry:
        try:
            after_keyword = log_entry.split(GADSME_SERVICE_KEYWORD, 1)[1]
            if "->" in after_keyword:
                method_part = after_keyword.split("->", 1)[1].split(":", 1)[0].strip()
            else:
                method_part = after_keyword.split(":", 1)[0].strip() or "Gadsme"
            json_str = extract_json_object_from_text(after_keyword)
            details = "N/A"
            json_data_for_log = "{}"
            if json_str:
                try:
                    data = json.loads(json_str)
                    details = format_json_html(data)
                    json_data_for_log = json_str
                except:
                    details = f'<div class="text-xs font-mono break-all text-red-600">JSON Parse Error</div><div class="text-xs font-mono break-all">{json_str}</div>'
            _append_callback_ad_row({
                "device_id": device_id,
                "device_name": get_device_name(device_id),
                "type": "Callback Gadsme",
                "event_name": method_part or "Gadsme",
                "details": details,
                "raw_log": log_entry.strip(),
                "json_data": json_data_for_log
            })
            return
        except:
            pass

    # --- 0b. Process Adverty callbacks ---
    if ADVERTY5_KEYWORD in log_entry:
        try:
            after_keyword = log_entry.split(ADVERTY5_KEYWORD, 1)[1]
            if "->" in after_keyword:
                method_part = after_keyword.split("->", 1)[1].split(":", 1)[0].strip()
            else:
                method_part = after_keyword.split(":", 1)[0].strip() or "Adverty"
            json_str = extract_json_object_from_text(after_keyword)
            details = "N/A"
            json_data_for_log = "{}"
            if json_str:
                try:
                    data = json.loads(json_str)
                    details = format_json_html(data)
                    json_data_for_log = json_str
                except:
                    details = f'<div class="text-xs font-mono break-all text-red-600">JSON Parse Error</div><div class="text-xs font-mono break-all">{json_str}</div>'
            _append_callback_ad_row({
                "device_id": device_id,
                "device_name": get_device_name(device_id),
                "type": "Callback Adverty5",
                "event_name": method_part or "Adverty",
                "details": details,
                "raw_log": log_entry.strip(),
                "json_data": json_data_for_log
            })
            return
        except:
            pass
    
    # --- 1. HANDLING BUFFERED IMPRESSION DATA (Split Logs) ---
    # Check if we are currently buffering for this device
    with lock:
        current_buffer = incomplete_impression_logs.get(device_id, "")

    current_split_key = ""
    if "_OnLevelPlayImpressionDataReadyEvent" in current_buffer:
        current_split_key = "_OnLevelPlayImpressionDataReadyEvent"
    elif "_OnImpressionDataReadyEvent" in current_buffer:
        current_split_key = "_OnImpressionDataReadyEvent"

    split_impression_key = ""
    if "_OnLevelPlayImpressionDataReadyEvent" in log_entry:
        split_impression_key = "_OnLevelPlayImpressionDataReadyEvent"
    elif "_OnImpressionDataReadyEvent" in log_entry:
        split_impression_key = "_OnImpressionDataReadyEvent"

    is_split_impression_start = bool(split_impression_key)

    # If we are buffering, OR if this is a new split impression event
    if current_buffer or is_split_impression_start:
        with lock:
            # Re-read buffer inside lock to be safe (though simple logic here is fine)
            current_buffer = incomplete_impression_logs.get(device_id, "")
            current_split_key = ""
            if "_OnLevelPlayImpressionDataReadyEvent" in current_buffer:
                current_split_key = "_OnLevelPlayImpressionDataReadyEvent"
            elif "_OnImpressionDataReadyEvent" in current_buffer:
                current_split_key = "_OnImpressionDataReadyEvent"
            
            # Case A: Start of new log
            if current_buffer and current_split_key and split_impression_key == current_split_key:
                # Continuation line may repeat the same callback prefix; append instead of
                # resetting the buffer so two split lines become one merged record.
                current_buffer += "\n" + log_entry
            elif is_split_impression_start:
                # Reset buffer with current line
                current_buffer = log_entry
                current_split_key = split_impression_key
            elif current_split_key and current_split_key in log_entry:
                # Case B: Continuation line
                current_buffer += "\n" + log_entry # Add newline to separate lines if needed, or just string concat
            else:
                # Buffer existed but this line is not the matching continuation.
                # Keep the old buffer for its own pair and let this line continue
                # through the normal parsing path below.
                incomplete_impression_logs[device_id] = current_buffer
                current_buffer = ""

            parse_buffer = current_buffer
            if current_split_key:
                callback_marker = f"LevelPlayAdService->{current_split_key}:"
                buffer_parts = []
                for buffer_line in current_buffer.splitlines():
                    cleaned_line = buffer_line.strip()
                    if callback_marker in cleaned_line:
                        cleaned_line = cleaned_line.split(callback_marker, 1)[1].strip()
                    buffer_parts.append(cleaned_line)
                parse_buffer = ''.join(buffer_parts)

            if not parse_buffer:
                pass
            else:
                # Try to find JSON
            # 1. Find first '{'
                start_idx = parse_buffer.find('{')
            
            # If no '{' yet, just keep buffering (unless it's been too long?)
            if parse_buffer and start_idx != -1:
                # 2. Count braces to find end
                open_braces = 0
                end_idx = -1
                for i in range(start_idx, len(parse_buffer)):
                    if parse_buffer[i] == '{': open_braces += 1
                    elif parse_buffer[i] == '}': open_braces -= 1
                    
                    if open_braces == 0:
                        end_idx = i
                        break
                
                if end_idx != -1:
                    # Found complete JSON
                    json_str = parse_buffer[start_idx : end_idx+1]
                    details = ""
                    json_data_for_log = "{}"
                    display_name = "_OnLevelPlayImpressionDataReadyEvent"
                    
                    try:
                        data = _parse_callback_json_payload(json_str)
                        if data is None:
                            data = json.loads(json_str)
                        if 'impressionData' in data:
                            details_target = data['impressionData']
                        else:
                            details_target = data
                        display_name = _levelplay_impression_event_name(details_target)
                        details = format_json_html(details_target)
                        json_data_for_log = json.dumps(details_target, ensure_ascii=False)
                    except:
                        details = f'<div class="text-xs font-mono break-all text-red-600">JSON Parse Error</div><div class="text-xs font-mono break-all">{json_str}</div>'
                        json_data_for_log = "{}"

                    # Clear buffer
                    incomplete_impression_logs[device_id] = ""
                    
                    # Emit
                    _append_callback_ad_row({
                        "device_id": device_id,
                        "device_name": get_device_name(device_id),
                        "type": "Callback Levelplay",
                        "event_name": display_name,
                        "details": details,
                        "raw_log": current_buffer.strip(),
                        "json_data": json_data_for_log
                    })
                    return # Done processing this line/buffer
                else:
                     # JSON start found but not ended -> Update buffer and wait for next line
                     incomplete_impression_logs[device_id] = current_buffer
                     return # Consumed line
            elif current_buffer:
                # No JSON start found yet (e.g. log line is "_OnImpressionDataReadyEvent:" and JSON is on next line)
                incomplete_impression_logs[device_id] = current_buffer
                return # Consumed line

    # --- 2. Process "ad_" events ---
    if event_name and event_name.startswith("ad_"):
        try:
            details = format_json_html(actual_params) if actual_params else "No params"
            _append_callback_ad_row({
                "device_id": device_id,
                "device_name": get_device_name(device_id),
                "type": "Ad Event",
                "event_name": event_name,
                "details": details,
                "raw_log": log_entry.strip(),
                "json_data": json_string,
            })
        except: pass
    
    # --- 3. Process Other Callbacks ---
    callback_match = CALLBACK_LOG_PATTERN.search(log_entry)
    if callback_match:
        found_key = callback_match.group(1)
        # Skip if it is impression data (handled above)
        if "_OnImpressionDataReadyEvent" in found_key or "_OnLevelPlayImpressionDataReadyEvent" in found_key: return

        details = "N/A"
        display_name = CALLBACK_DISPLAY_NAMES.get(found_key, found_key)
        json_data_for_log = "{}"
        
        if found_key == "Receive Ironsource Impression Data LevelPlayImpressionData":
             try:
                 payload = log_entry.split(found_key, 1)[1].strip()
                 parsed_data = parse_levelplay_impression_text(payload)
                 if parsed_data:
                     display_name = _levelplay_impression_event_name(parsed_data)
                     details = format_json_html(parsed_data)
                     json_data_for_log = json.dumps(parsed_data, ensure_ascii=False)
                 else:
                     details = f'<div class="text-xs font-mono break-all">{html.escape(payload)}</div>' if payload else "Ironsource Impression Data"
             except:
                 details = "Ironsource Impression Data"
        elif "Listener" in found_key:
             # Try parsing adInfo if present
             try:
                 if 'adInfo:' in log_entry:
                     # Simple parsing for adInfo string
                     parts = log_entry.split('adInfo:')[1].strip().split(',')
                     ad_info = {}
                     for part in parts:
                         if ':' in part:
                             k, v = part.split(':', 1)
                             ad_info[k.strip()] = v.strip()
                     details = format_json_html(ad_info)
                     json_data_for_log = json.dumps(ad_info)
                 else:
                     details = "Listener Fired (No adInfo)"
             except:
                 details = "Listener Fired"

        _append_callback_ad_row({
            "device_id": device_id,
            "device_name": get_device_name(device_id),
            "type": "Callback Levelplay",
            "event_name": display_name,
            "details": details,
            "raw_log": log_entry.strip(),
            "json_data": json_data_for_log,
        })

def process_event_validator_log(event_name, actual_params, json_string, log_entry, device_id):
    global validator_seq
    if is_paused or not validator_active: 
        return
    source = "firebase"
    try:
        payload = json.loads(json_string)
        source = payload.get("source", "firebase")
    except:
        pass
    with lock:
        required_all = []
        required_all.extend(default_params)
        required_all.extend(required_params)
        specific = event_specific_params.get(event_name, [])
        required_all.extend(specific)

        required_set = set(required_all)
        actual_set = set(actual_params.keys())
        missing = sorted(required_set - actual_set)
        strange = sorted(actual_set - required_set) if required_set else []
        status = "PASSED" if not missing and not strange else "FAILED"
        
        # New Formatting Logic: Highlight missing -> Show full JSON
        details_html = ""
        if not specific and event_specific_params:
            closest = None
            best_dist = 3
            for known in event_specific_params.keys():
                d = _levenshtein_distance_limit(event_name, known, limit=2)
                if d < best_dist:
                    best_dist = d
                    closest = known
                    if best_dist == 1:
                        break
            if closest and best_dist <= 2:
                details_html += f'<div class="text-orange-600 font-bold mb-2">Possible typo: "{event_name}" ~ "{closest}"</div>'
        if missing:
            details_html += format_param_issue_html("Missing", missing, "text-red-600")
        if strange:
            details_html += format_param_issue_html("Strange", strange, "text-orange-600")
        
        details_html += format_json_html(actual_params)
        
        if not _accept_exact_record("validator", device_id, log_entry, event_name, source):
            return
        validator_seq += 1
        row = {"device_id": device_id, "event_name": event_name, "device_name": get_device_name(device_id), "status": status, "details": details_html, "raw_log": log_entry.strip(), "json_data": json_string, "source": source, "seq": validator_seq}
        validator_results.append(row)
        # The UI receives only this new row (coalesced by live_table_emitter)
        # instead of the whole table on every event.
        validator_pending.append(row)

def _build_specific_result(item):
    """Validation row for one cached event, or None when the name filter skips it.

    This is the exact per-row logic the table always used; a full rebuild and
    a single new event now share it, so both produce identical rows.
    """
    try:
        data = json.loads(item['json_data'])
        evt = data.get('eventName')
        params = data.get('e', {})

        source = data.get('source', 'firebase')

        # Name Filter
        if specific_event_name_filters and not any(evt.startswith(f) for f in specific_event_name_filters):
            return None

        # Param Validation (Default Params + Default Events)
        status = "INFO"
        details = ""

        required_all = []
        required_all.extend(default_params)
        # If user input params, combine with defaults
        if specific_event_params_filters:
            required_all.extend(specific_event_params_filters)
        # Add event-specific params if event is in default event list
        if evt in event_specific_params:
            required_all.extend(event_specific_params.get(evt, []))

        if required_all:
            required_set = set(required_all)
            actual_set = set(params.keys())
            missing = sorted(required_set - actual_set)
            strange = sorted(actual_set - required_set)
            status = "PASSED" if not missing and not strange else "FAILED"

            if missing:
                details += format_param_issue_html("Missing", missing, "text-red-600")

            if strange:
                details += format_param_issue_html("Strange", strange, "text-orange-600")

            # Show full JSON
            details += format_json_html(params)
        else:
            # Show all params json
            details = format_json_html(params)

        return {"device_id": item['device_id'], "device_name": get_device_name(item['device_id']), "status": status, "event_name": evt, "details": details, "raw_log": item['log'], "json_data": item['json_data'], "source": source, "seq": item.get('seq', 0)}
    except Exception:
        return None


def _apply_specific_filter_and_emit():
    """Rebuild every row (filters or profile changed) and send the whole table."""
    global specific_event_results, specific_event_pending
    with lock:
        res = []
        for item in event_log_cache:
            row = _build_specific_result(item)
            if row is not None:
                res.append(row)
        specific_event_results = res
        # The full table already contains anything still waiting for a delta.
        specific_event_pending = []
    socketio.emit('update_specific_event_table', specific_event_results)


def _emit_specific_event_snapshot(to=None):
    """Send the current table (rows already built) to one client.

    New rows reach clients as deltas, so a page that (re)connects needs the
    rows it missed once.  Pending rows are left for the next delta; the
    client ignores any row it already holds.
    """
    with lock:
        rows = list(specific_event_results)
    if to is not None:
        socketio.emit('update_specific_event_table', rows, to=to)
    else:
        socketio.emit('update_specific_event_table', rows)


def _flush_specific_event_delta():
    """Send rows added since the last flush; `min_seq` says which old rows left the cache."""
    global specific_event_pending
    with lock:
        if not specific_event_pending:
            return 0
        rows = specific_event_pending
        specific_event_pending = []
        min_seq = event_log_cache[0].get('seq', 0) if event_log_cache else 0
    socketio.emit('specific_event_delta', {"rows": rows, "min_seq": min_seq})
    return len(rows)


def _live_table_append(name, row):
    """Append one row to a live table and queue it for the next UI delta.

    Used by Load Ads, Load Ads Ext, Callback & Ads and Price Rotation: the
    row keeps every field it had, plus a `seq` the UI uses to order rows and
    to drop rows that left the table.
    """
    global live_table_seq
    with lock:
        live_table_seq += 1
        row["seq"] = live_table_seq
        LIVE_TABLE_ROWS[name].append(row)
        LIVE_TABLE_PENDING[name].append(row)


def _flush_live_table(name):
    """Send the rows queued for one table; `min_seq` marks rows that left it."""
    with lock:
        rows = LIVE_TABLE_PENDING[name]
        if not rows:
            return 0
        LIVE_TABLE_PENDING[name] = []
        table = LIVE_TABLE_ROWS[name]
        min_seq = table[0].get("seq", 0) if table else 0
    socketio.emit(LIVE_TABLE_DELTA_EVENTS[name], {"rows": rows, "min_seq": min_seq})
    return len(rows)


def _clear_live_table_pending(*names):
    with lock:
        for name in names or LIVE_TABLE_PENDING.keys():
            LIVE_TABLE_PENDING[name] = []


def _flush_validator_delta():
    """Send Validator rows added since the last flush; `min_seq` marks rows that left the table."""
    global validator_pending
    with lock:
        if not validator_pending:
            return 0
        rows = validator_pending
        validator_pending = []
        min_seq = validator_results[0].get('seq', 0) if validator_results else 0
    socketio.emit('validator_delta', {"rows": rows, "min_seq": min_seq})
    return len(rows)


def _emit_validator_snapshot(to=None):
    """Send the current Validator table to one (re)connecting client."""
    with lock:
        rows = list(validator_results)
    if to is not None:
        socketio.emit('update_validator_table', rows, to=to)
    else:
        socketio.emit('update_validator_table', rows)


def live_table_emitter():
    """Coalesce live table rows (Specific Validator, Validator): one delta per interval."""
    while True:
        time.sleep(TABLE_DELTA_INTERVAL_SECONDS)
        flushes = [("Specific Validator", _flush_specific_event_delta), ("Validator", _flush_validator_delta),
                   ("Revenue", _flush_adrevenue_delta)]
        flushes += [(name, lambda name=name: _flush_live_table(name)) for name in LIVE_TABLE_DELTA_EVENTS]
        for name, flush in flushes:
            try:
                flush()
            except Exception as exc:
                logging.warning("%s delta failed: %s", name, exc)


def cache_specific_event_log(event_name, params, json_string, log_entry, device_id):
    global specific_event_seq, specific_event_results
    if is_paused: return
    with lock:
        if not _accept_exact_record("specific", device_id, log_entry, event_name):
            return
        specific_event_seq += 1
        item = {'log': log_entry, 'device_id': device_id, 'json_data': json_string, 'seq': specific_event_seq}
        event_log_cache.append(item)
        # Only the new event is validated and rendered; the rest of the table
        # is unchanged, so the UI receives just this row.
        row = _build_specific_result(item)
        if row is not None:
            specific_event_results.append(row)
            specific_event_pending.append(row)
        oldest_seq = event_log_cache[0].get('seq', 0)
        if specific_event_results and specific_event_results[0].get('seq', 0) < oldest_seq:
            specific_event_results = [r for r in specific_event_results if r.get('seq', 0) >= oldest_seq]

def _record_adrevenue_log(device_id, source, event_name, parsed_data, raw_log, raw_details=None, raw_event_prefix=None, skip_validation=False):
    global adrevenue_seq, adrevenue_rendered
    payload = parsed_data if isinstance(parsed_data, dict) else {}
    item = {
        "device_id": device_id,
        "device_name": get_device_name(device_id),
        "status": "INFO",
        "event_name": event_name,
        "source": source,
        "details": format_json_html(payload) if payload else (raw_details or ""),
        "raw_details": raw_details or "",
        "raw_log": raw_log.strip(),
        "json_data": json.dumps(payload, ensure_ascii=False) if payload else "{}",
        "parsed_data": payload,
        "skip_validation": skip_validation,
    }
    if raw_event_prefix:
        item["raw_event_prefix"] = raw_event_prefix
    with lock:
        if _accept_exact_record("adrevenue", device_id, raw_log, source, event_name, raw_event_prefix or ""):
            adrevenue_seq += 1
            item["seq"] = adrevenue_seq
            adrevenue_logs.append(item)
            # Validate and render only this item; earlier rows are unchanged.
            row = _build_adrevenue_row(item)
            adrevenue_rendered.append(row)
            adrevenue_pending.append(row)
            oldest_seq = adrevenue_logs[0].get("seq", 0)
            if adrevenue_rendered and adrevenue_rendered[0].get("seq", 0) < oldest_seq:
                adrevenue_rendered = [r for r in adrevenue_rendered if r.get("seq", 0) >= oldest_seq]

def _normalize_ios_appmetrica_adrevenue(ad_revenue):
    payload = ad_revenue.get("Payload") if isinstance(ad_revenue.get("Payload"), dict) else {}
    return {
        "adRevenue": ad_revenue.get("AdRevenueValue"),
        "currency": ad_revenue.get("Currency"),
        "adType": ad_revenue.get("AdType"),
        "adNetwork": ad_revenue.get("AdNetwork"),
        "adUnitId": ad_revenue.get("AdUnitId"),
        "adUnitName": ad_revenue.get("AdUnitName"),
        "adPlacementId": ad_revenue.get("AdPlacementId"),
        "adPlacementName": ad_revenue.get("AdPlacementName"),
        "precision": ad_revenue.get("Precision"),
        "payload": payload,
    }

def _normalize_ios_appsflyer_adrevenue(data):
    revenue = data.get("adRevenueData") if isinstance(data.get("adRevenueData"), dict) else {}
    if not revenue:
        revenue = data.get("logRevenue") if isinstance(data.get("logRevenue"), dict) else {}
    additional = data.get("additionalParameters") if isinstance(data.get("additionalParameters"), dict) else {}
    if not additional:
        additional = data.get("additionalParams") if isinstance(data.get("additionalParams"), dict) else {}
    return {
        "ad_network": {
            "name": "ad_revenue",
            "monetization_network": revenue.get("monetizationNetwork"),
            "event_revenue_currency": revenue.get("currencyIso4217Code"),
            "mediation_network": revenue.get("mediationNetwork"),
            "event_revenue": revenue.get("eventRevenue"),
            "payload": additional,
        }
    }

IOS_ADREVENUE_KEYWORDS = {
    "appmetrica": (
        "AppMetricaTrackingHandler->_Handle:",
        "AppMetricaAdRevenueTrackingHandler->Handle:",
        "Appmetrica TrackAdRevenueEventForOther:",
        "AppMetrica TrackAdRevenueEventForOther:",
        "Appmetrica TrackAdRevenueEvent:",
        "AppMetrica TrackAdRevenueEvent:",
    ),
    "appsflyer": (
        "AppsFlyerTrackingHandler->Track:",
        "AppsFlyerAdRevenueTrackingHandler->Handle:",
    ),
}

def _ios_adrevenue_source_and_payload(line):
    for source, keywords in IOS_ADREVENUE_KEYWORDS.items():
        for keyword in keywords:
            if keyword in line:
                if source == "appsflyer" and keyword == "AppsFlyerTrackingHandler->Track:" and "[Tracking,AppsFlyer,Ad]" not in line:
                    return None, "", ""
                if source == "appmetrica" and keyword == "AppMetricaTrackingHandler->_Handle:" and "[Tracking,AppMetrica]" not in line:
                    return None, "", ""
                return source, keyword, line.split(keyword, 1)[1].strip()
    return None, "", ""

def _buffer_ios_adrevenue_payload(device_id, source, payload_part, raw_line):
    buffer_key = f"{device_id}:{source}"
    with lock:
        current = incomplete_ios_adrevenue_logs.get(buffer_key)
        if current:
            current["payload"] += payload_part
            current["raw_log"] += "\n" + raw_line.strip()
        else:
            current = {"payload": payload_part, "raw_log": raw_line.strip()}
        if len(current["payload"]) > 50000:
            incomplete_ios_adrevenue_logs.pop(buffer_key, None)
            return None, ""

        json_str = extract_json_object_from_text(current["payload"])
        if json_str:
            incomplete_ios_adrevenue_logs.pop(buffer_key, None)
            return json_str, current["raw_log"]

        incomplete_ios_adrevenue_logs[buffer_key] = current
    return None, ""

def _buffer_ios_load_ads_ext_payload(device_id, payload_part, raw_line):
    with lock:
        current = incomplete_ios_load_ads_ext_logs.get(device_id)
        if current:
            current["payload"] += payload_part
            current["raw_log"] += "\n" + raw_line.strip()
        else:
            current = {"payload": payload_part, "raw_log": raw_line.strip()}
        if len(current["payload"]) > 50000:
            incomplete_ios_load_ads_ext_logs.pop(device_id, None)
            return None, ""

        json_str = extract_json_object_from_text(current["payload"])
        if json_str:
            incomplete_ios_load_ads_ext_logs.pop(device_id, None)
            return json_str, current["raw_log"]

        incomplete_ios_load_ads_ext_logs[device_id] = current
    return None, ""

def _loads_adrevenue_json_payload(json_str):
    for candidate in (json_str, _decode_ios_levelplay_json_text(json_str)):
        try:
            data = json.loads(candidate) if candidate else {}
            return data if isinstance(data, dict) else {}
        except Exception:
            pass
    return {}

def _is_adjust_tag_log(line):
    text = str(line or "")
    return bool(re.search(
        r'(?:'
        r'AdjustSdk|AdjustTrackingHandler|AdjustService|'
        r'\[\s*Tracking\s*,\s*Adjust\b|'
        r'\[\s*Adjust(?:Sdk)?\s*\]|'
        r'(?:^|[\s(\[])Adjust(?:\s|:|\)|\])'
        r')',
        text,
        re.IGNORECASE,
    ))


def _buffer_adjust_adrevenue_param(device_id, param_type, payload_part, raw_line):
    buffer_key = f"{device_id}:{param_type}"
    with lock:
        current = incomplete_adjust_adrevenue_logs.get(buffer_key)
        if current:
            current["payload"] += payload_part
            current["raw_log"] += "\n" + raw_line.strip()
        else:
            current = {"payload": payload_part, "raw_log": raw_line.strip()}
        if len(current["payload"]) > 50000:
            incomplete_adjust_adrevenue_logs.pop(buffer_key, None)
            return None, ""

        json_str = extract_json_object_from_text(current["payload"])
        if json_str:
            incomplete_adjust_adrevenue_logs.pop(buffer_key, None)
            return json_str, current["raw_log"]

        incomplete_adjust_adrevenue_logs[buffer_key] = current
    return None, ""


def _resume_adjust_adrevenue_param_if_needed(device_id, line):
    with lock:
        pending_keys = [key for key in incomplete_adjust_adrevenue_logs if key.startswith(f"{device_id}:")]
    for key in pending_keys:
        param_type = key.split(":", 1)[1]
        json_str, raw_log = _buffer_adjust_adrevenue_param(device_id, param_type, line.strip(), line)
        if json_str:
            return param_type, json_str, raw_log
    return None, "", ""


def process_adrevenue_log(line, device_id):
    handled = False

    if "AdRevenue Received:" in line:
        match = ADREVENUE_LOG_PATTERN.search(line)
        if match:
            content = match.group(1)
            ad_data = {}
            try:
                payload_obj = {}
                payload_match = re.search(r'payload=(\{.*?\})(?:,|$)', content)

                clean_content = content
                if payload_match:
                    payload_str = payload_match.group(1)
                    try:
                        payload_obj = json.loads(payload_str)
                    except:
                        payload_obj = {}
                    clean_content = content.replace(f'payload={payload_str}', '')

                parts = [p.strip() for p in clean_content.split(',') if p.strip()]
                for part in parts:
                    if '=' in part:
                        k, v = part.split('=', 1)
                        ad_data[k.strip()] = v.strip()

                if payload_obj:
                    ad_data['payload'] = payload_obj
            except:
                ad_data = {}

            with lock:
                _record_adrevenue_log(device_id, "appmetrica", "AdRevenue - Appmetrica", ad_data, line, content)
            handled = True

    if (not handled) and "AppsFlyer" in line and "ADREVENUE-" in line and "preparing data:" in line:
        match = APPSFLYER_ADREVENUE_PATTERN.search(line)
        if match:
            event_prefix = match.group(1).upper()
            json_str = match.group(2)
            appsflyer_data = {}
            try:
                appsflyer_data = json.loads(json_str)
            except:
                appsflyer_data = {}
            with lock:
                _record_adrevenue_log(device_id, "appsflyer", "AdRevenue - Appsflyer", appsflyer_data, line, json_str, event_prefix)
            handled = True

    purchase_match = re.search(
        r'AdjustTrackingHandler\s*->\s*_LogPurchaseVerificationResult\s*:',
        line,
        re.IGNORECASE,
    )
    if (not handled) and purchase_match:
        payload_part = line[purchase_match.end():].strip()
        json_str = extract_json_object_from_text(payload_part) or ""
        if json_str:
            verification_data = _loads_adrevenue_json_payload(json_str)
            with lock:
                _record_adrevenue_log(
                    device_id,
                    "adjust",
                    "Adjust Purchase Verification",
                    verification_data,
                    line,
                    json_str,
                    "purchase_verification",
                    skip_validation=True,
                )
            handled = True

    initialize_match = re.search(
        r'AdjustService\s*->\s*Initialize\s*:',
        line,
        re.IGNORECASE,
    )
    if (not handled) and initialize_match:
        json_str = extract_json_object_from_text(line)
        if json_str:
            adjust_config_data = _loads_adrevenue_json_payload(json_str)
            config_data = adjust_config_data.get("config") if isinstance(adjust_config_data.get("config"), dict) else adjust_config_data
            if isinstance(config_data, dict) and config_data:
                with lock:
                    _record_adrevenue_log(
                        device_id,
                        "adjust",
                        "Adjust Config",
                        config_data,
                        line,
                        json_str,
                        "config",
                        skip_validation=True,
                    )
                handled = True

    adjust_track_match = re.search(
        r'AdjustTrackingHandler\s*->\s*Track\s*:',
        line,
        re.IGNORECASE,
    )
    if (
        (not handled)
        and re.search(r'\[\s*Tracking\s*,\s*Adjust\s*,\s*Iap\s*\]', line, re.IGNORECASE)
        and adjust_track_match
    ):
        payload_part = line[adjust_track_match.end():].strip()
        json_str = extract_json_object_from_text(payload_part) or ""
        if json_str:
            iap_data = _loads_adrevenue_json_payload(json_str)
            event_name = "Adjust Subscription" if isinstance(iap_data.get("subscription"), dict) else "Adjust IAP"
            with lock:
                _record_adrevenue_log(device_id, "adjust", event_name, iap_data, line, json_str, "iap", skip_validation=True)
            handled = True

    # Some Android/iOS integrations emit AdjustTrackingHandler->Track without
    # the older [Tracking,Adjust,Iap] prefix.  Keep those records visible too;
    # do not require a particular app/process name or platform.
    if (not handled) and adjust_track_match:
        payload_part = line[adjust_track_match.end():].strip()
        json_str = extract_json_object_from_text(payload_part) or ""
        if json_str:
            adjust_data = _loads_adrevenue_json_payload(json_str)
            event_name = "Adjust Subscription" if isinstance(adjust_data.get("subscription"), dict) else "AdRevenue - Adjust Track"
            with lock:
                _record_adrevenue_log(
                    device_id,
                    "adjust",
                    event_name,
                    adjust_data,
                    line,
                    json_str,
                    "track",
                    skip_validation=True,
                )
            handled = True

    ios_source, ios_keyword, ios_payload_part = (None, "", "")
    if (not handled) and active_platform == "ios":
        ios_source, ios_keyword, ios_payload_part = _ios_adrevenue_source_and_payload(line)
    if (not handled) and active_platform == "ios" and ios_source:
        try:
            json_str, raw_log = _buffer_ios_adrevenue_payload(device_id, ios_source, ios_payload_part, line)
            if not json_str:
                return True
            data = _loads_adrevenue_json_payload(json_str)
            if ios_source == "appmetrica":
                ad_revenue = data.get("adRevenue") if isinstance(data.get("adRevenue"), dict) else {}
                if ad_revenue:
                    normalized = _normalize_ios_appmetrica_adrevenue(ad_revenue)
                elif isinstance(data, dict) and any(key in data for key in ("AdRevenueValue", "Currency", "AdNetwork", "Payload")):
                    normalized = _normalize_ios_appmetrica_adrevenue(data)
                else:
                    normalized = data
                event_name = "AdRevenue - Appmetrica"
            else:
                normalized = _normalize_ios_appsflyer_adrevenue(data) if data else {}
                event_name = "AdRevenue - Appsflyer"
            with lock:
                _record_adrevenue_log(device_id, ios_source, event_name, normalized, raw_log or line, json_str)
            handled = True
        except:
            pass

    if (not handled) and _is_adjust_tag_log(line):
        param_type, json_str, raw_log = _resume_adjust_adrevenue_param_if_needed(device_id, line)
        if json_str:
            adjust_data = _loads_adrevenue_json_payload(json_str)
            with lock:
                _record_adrevenue_log(device_id, "adjust", f"AdRevenue - Adjust {param_type}", adjust_data, raw_log or line, json_str, param_type)
            handled = True

    if (not handled) and _is_adjust_tag_log(line):
        lvl_match = re.search(r'\blvl_(signature|signed_data)\b\s+(.+)$', line)
        if lvl_match:
            field_name = f"lvl_{lvl_match.group(1)}"
            field_value = lvl_match.group(2).strip()
            lvl_data = {field_name: field_value}
            with lock:
                _record_adrevenue_log(
                    device_id,
                    "adjust",
                    "Adjust LVL",
                    lvl_data,
                    line,
                    json.dumps(lvl_data, ensure_ascii=False),
                    field_name,
                    skip_validation=True,
                )
            handled = True

    if (not handled) and _is_adjust_tag_log(line) and re.search(r'\b(callback_params|partner_params)\b', line, re.IGNORECASE):
        match = re.search(r'\b(callback_params|partner_params)\b', line, re.IGNORECASE)
        if match:
            param_type = match.group(1).lower()
            tail = line[match.end():]
            json_str, raw_log = _buffer_adjust_adrevenue_param(device_id, param_type, tail, line)
            if not json_str:
                return True
            adjust_data = _loads_adrevenue_json_payload(json_str)
            with lock:
                _record_adrevenue_log(device_id, "adjust", f"AdRevenue - Adjust {param_type}", adjust_data, raw_log or line, json_str, param_type)
            handled = True

    if (not handled) and _is_adjust_tag_log(line) and re.search(r'\bsource\b', line, re.IGNORECASE):
        match = re.search(r'(?:(?:\bsource\s+)|(?:"source"\s*:\s*"))(ironsource[A-Za-z0-9_.-]*)', line, re.IGNORECASE)
        if match:
            source_data = {"source": match.group(1)}
            with lock:
                _record_adrevenue_log(device_id, "adjust", "Adjust adrevenue", source_data, line, json.dumps(source_data, ensure_ascii=False), "source", skip_validation=True)
            handled = True

    if (not handled) and _is_adjust_tag_log(line) and ("Response:" in line or "Response string:" in line):
        match = re.search(r'Response(?:\s+string)?\s*:\s*(\{.*\})', line, re.IGNORECASE)
        if match:
            json_str = match.group(1).strip()
            response_data = _loads_adrevenue_json_payload(json_str)
            message = str(response_data.get("message", "")).lower()
            is_adjust_revenue_response = (
                bool(response_data.get("timestamp")) and (
                    (message.startswith("revenue tracked")) or
                    (message == "ad revenue tracked")
                )
            )
            if is_adjust_revenue_response:
                with lock:
                    event_name = "Adjust IAP" if message.startswith("revenue tracked") else "Adjust adrevenue"
                    _record_adrevenue_log(device_id, "adjust", event_name, response_data, line, json_str, "response", skip_validation=True)
                handled = True

    # New rows were queued by _record_adrevenue_log and leave as a delta.
    return handled

def _build_adrevenue_row(item):
    """Validated Revenue row for one recorded item (the table's exact per-row logic)."""
    parsed_data = item.get("parsed_data") or {}
    source = (item.get("source") or "appmetrica").lower()

    if source == "appsflyer":
        ad_network = parsed_data.get("ad_network") if isinstance(parsed_data.get("ad_network"), dict) else {}
        root_data = ad_network if ad_network else parsed_data
        payload_data = root_data.get("payload") if isinstance(root_data.get("payload"), dict) else {}
        custom_params = payload_data.get("custom_parameters") if isinstance(payload_data.get("custom_parameters"), dict) else {}
        validate_maps = [root_data, payload_data, custom_params]
        details_target = root_data if root_data else parsed_data
    else:
        payload_data = parsed_data.get("payload") if isinstance(parsed_data.get("payload"), dict) else {}
        validate_maps = [parsed_data, payload_data]
        details_target = parsed_data if parsed_data else item.get("raw_details", "")

    if item.get("skip_validation"):
        if item.get("event_name") == "Adjust Config":
            sending_background = parsed_data.get("IsSendingInBackgroundEnabled")
            cost_data_attr = parsed_data.get("IsCostDataInAttributionEnabled")
            ad_services = parsed_data.get("IsAdServicesEnabled")
            idfa_reading = parsed_data.get("IsIdfaReadingEnabled")
            is_adjust_config_ok = (
                sending_background is True and
                cost_data_attr is True and
                ad_services is True and
                idfa_reading is True
            )
            return ({
                **item,
                "status": "PASSED" if is_adjust_config_ok else "FAILED",
                "details": format_adjust_config_html(details_target) if isinstance(details_target, dict) else (details_target or item.get("details", "")),
            })
        return ({
            **item,
            "status": "INFO",
            "details": format_json_html(details_target) if isinstance(details_target, dict) else (details_target or item.get("details", "")),
        })

    required_all = list(adrevenue_default_params)
    normalized_source = _normalize_adrevenue_sheet_key(source)
    for alias in (
        normalized_source,
        _normalize_adrevenue_sheet_key(item.get("event_name", "")),
        "all",
    ):
        required_all.extend(adrevenue_source_params.get(alias, []))
    seen_required = []
    seen_set = set()
    for param in required_all:
        if param not in seen_set:
            seen_set.add(param)
            seen_required.append(param)

    missing = []
    skipped_ios = []
    for param in seen_required:
        if any(isinstance(m, dict) and param in m for m in validate_maps):
            continue
        if active_platform == "ios" and (
            (source == "appsflyer" and param == "custom_parameters") or
            (source == "appmetrica" and param == "autoCollected")
        ):
            skipped_ios.append(param)
            continue
        missing.append(param)

    actual_keys = set()
    for candidate in validate_maps:
        if isinstance(candidate, dict):
            actual_keys.update(candidate.keys())
    strange = sorted(actual_keys - set(seen_required)) if seen_required else []

    status = "INFO"
    if seen_required:
        status = "PASSED" if not missing else "FAILED"

    summary_parts = []
    if seen_required:
        if missing:
            summary_parts.append(format_param_issue_html("Missing", missing, "text-red-600", chunk_size=1))
        else:
            summary_parts.append("<div class='mb-2 text-xs font-medium text-green-600'>All required params found</div>")
        if skipped_ios:
            summary_parts.append(format_param_issue_html("Skip in iOS", skipped_ios, "text-slate-500", chunk_size=1))
        if strange:
            summary_parts.append(format_param_issue_html("Strange", strange, "text-orange-600", chunk_size=1))

    details_html = ''.join(summary_parts) + format_json_html(details_target)
    return ({
        **item,
        "status": status,
        "details": details_html,
    })


def _apply_adrevenue_filter_and_emit():
    """Rebuild every Revenue row (params or profile changed) and send the whole table."""
    global adrevenue_rendered, adrevenue_pending
    with lock:
        adrevenue_rendered = [_build_adrevenue_row(item) for item in adrevenue_logs]
        rendered = list(adrevenue_rendered)
        # The full table already contains anything still waiting for a delta.
        adrevenue_pending = []
    socketio.emit('update_adrevenue_table', rendered)


def _flush_adrevenue_delta():
    """Send Revenue rows added since the last flush; `min_seq` marks rows that left the table."""
    global adrevenue_pending
    with lock:
        if not adrevenue_pending:
            return 0
        rows = adrevenue_pending
        adrevenue_pending = []
        min_seq = adrevenue_logs[0].get("seq", 0) if adrevenue_logs else 0
    socketio.emit('adrevenue_delta', {"rows": rows, "min_seq": min_seq})
    return len(rows)


def _emit_live_table_snapshots(to):
    """Full tables for one (re)connecting client; new rows then arrive as deltas."""
    with lock:
        payloads = [
            ('update_load_ads', list(load_ads_events)),
            ('update_load_ads_ext', list(load_ads_ext_events)),
            ('update_callback_ad_table', list(callback_ad_logs)),
            ('update_price_rotation_table', list(price_rotation_logs)),
            ('update_adrevenue_table', list(adrevenue_rendered)),
        ]
    for event, rows in payloads:
        socketio.emit(event, rows, to=to)


def _sdk_result_group_sort_key(rows):
    statuses = {row.get("status") for row in rows if row.get("status")}
    if "FAILED" in statuses:
        return 0
    if "PASSED" in statuses:
        return 1
    if "FOUND" in statuses:
        return 2
    return 3


def _emit_sdk_check_results():
    res = []
    with lock:
        if active_platform == "ios" and not connected_devices_info:
            res.append({"status": "HEADER", "display_text": "--- iOS ---", "device_name": "iOS", "device_id": "ios"})
            res.append({"status": "WAITING", "display_text": "Waiting for iOS device...", "device_id": "ios"})

        def append_sdk_network_rows(device_id, network_key, block):
            rows = []
            if _normalize_sdk_network_name(block.get("display_name")) == "googleplayservices":
                return rows
            expected_key = block.get("expected_key") or _match_sdk_expected_key(block.get("display_name"))
            expected = sdk_check_expected_map.get(expected_key, {})
            rows.append({"status": "LABEL", "display_text": block.get("display_name", network_key), "device_id": device_id})

            actual_sdk = block.get("sdk_version", "")
            expected_sdk = expected.get("sdk", "")
            actual_adapter = block.get("adapter_version", "")
            if not actual_adapter and block.get("adapter_missing"):
                actual_adapter = "MISSING"
            expected_adapter = expected.get("adapter", "")

            if expected.get("source") == "cloudx":
                cloudx_status = _sdk_result_status(actual_adapter, expected_adapter)
                cloudx_actual_display = actual_adapter
                if cloudx_actual_display and cloudx_actual_display.upper() not in {"NOT FOUND", "MISSING"}:
                    cloudx_actual_display = _extract_sdk_comparable_version(cloudx_actual_display, expected_adapter)
                rows.append({
                    "status": cloudx_status,
                    "display_text": f"Adapter Version  Actual: {cloudx_actual_display or 'NOT FOUND'}  Expected: {expected_adapter or '—'}",
                    "device_id": device_id
                })
                expected_native = expected.get("sdk", "")
                if expected.get("source") == "cloudx" and expected_native:
                    actual_native = block.get("native_version", "")
                    native_status = _sdk_result_status(actual_native, expected_native)
                    native_actual_display = actual_native
                    if native_actual_display and native_actual_display.upper() not in {"NOT FOUND", "MISSING"}:
                        native_actual_display = _extract_sdk_comparable_version(native_actual_display, expected_native)
                    rows.append({
                        "status": native_status,
                        "display_text": f"Native Version  Actual: {native_actual_display or 'NOT FOUND'}  Expected: {expected_native}",
                        "device_id": device_id
                    })
                return rows

            if bool(expected_sdk) ^ bool(expected_adapter):
                if expected_sdk:
                    actual_single = actual_sdk
                    expected_single = expected_sdk
                else:
                    actual_single = actual_adapter
                    expected_single = expected_adapter
                single_status = _sdk_result_status(actual_single, expected_single)
                actual_single_display = actual_single
                if actual_single_display and actual_single_display.upper() not in {"NOT FOUND", "MISSING"}:
                    actual_single_display = _extract_sdk_comparable_version(actual_single_display, expected_single)
                rows.append({
                    "status": single_status,
                    "display_text": f"Version  Actual: {actual_single_display or 'NOT FOUND'}  Expected: {expected_single}",
                    "device_id": device_id
                })
            else:
                sdk_status = _sdk_result_status(actual_sdk, expected_sdk)
                actual_sdk_display = actual_sdk
                if actual_sdk_display and actual_sdk_display.upper() not in {"NOT FOUND", "MISSING"}:
                    actual_sdk_display = _extract_sdk_comparable_version(actual_sdk_display, expected_sdk)
                rows.append({
                    "status": sdk_status,
                    "display_text": f"SDK  Actual: {actual_sdk_display or 'NOT FOUND'}  Expected: {expected_sdk}",
                    "device_id": device_id
                })

                adapter_status = _sdk_result_status(actual_adapter, expected_adapter)
                actual_adapter_display = actual_adapter
                if actual_adapter_display and actual_adapter_display.upper() not in {"NOT FOUND", "MISSING"}:
                    actual_adapter_display = _extract_sdk_comparable_version(actual_adapter_display, expected_adapter)
                rows.append({
                    "status": adapter_status,
                    "display_text": f"Adapter  Actual: {actual_adapter_display or 'NOT FOUND'}  Expected: {expected_adapter}",
                    "device_id": device_id
                })
            return rows

        for dev in connected_devices_info:
            res.append({"status": "HEADER", "display_text": f"--- {dev['name']} ---", "device_name": dev['name'], "device_id": dev['id']})
            device_state = sdk_check_runtime_state.get(dev['id'], {})
            if not device_state and not sdk_check_expected_order:
                wait_text = "Waiting for iOS SDK logs..." if active_platform == "ios" else "Waiting for IntegrationHelper logs..."
                res.append({"status": "WAITING", "display_text": wait_text, "device_id": dev['id']})
                continue

            _ensure_sdk_expected_blocks_for_device(dev['id'])
            device_state = sdk_check_runtime_state.get(dev['id'], {})
            in_list_keys = [key for key in _sdk_state_network_keys(dev['id']) if key in sdk_check_expected_map]
            not_in_list_keys = [key for key in _sdk_state_network_keys(dev['id']) if key not in sdk_check_expected_map]

            in_list_groups = []
            for network_key in in_list_keys:
                rows = append_sdk_network_rows(dev['id'], network_key, device_state.get(network_key, {}))
                if rows:
                    in_list_groups.append(rows)
            for rows in sorted(in_list_groups, key=_sdk_result_group_sort_key):
                res.extend(rows)

            if not_in_list_keys:
                res.append({"status": "SECTION", "display_text": "----------- Not in List -------", "device_id": dev['id']})
                not_in_list_groups = []
                for network_key in not_in_list_keys:
                    rows = append_sdk_network_rows(dev['id'], network_key, device_state.get(network_key, {}))
                    if rows:
                        not_in_list_groups.append(rows)
                for rows in sorted(not_in_list_groups, key=_sdk_result_group_sort_key):
                    res.extend(rows)
    socketio.emit('update_sdk_check_table', res)


def _sdk_check_block_for_device(device_id, network_name):
    device_state = sdk_check_runtime_state.setdefault(device_id, {})
    raw_network_key = _normalize_sdk_network_name(network_name)
    expected_key = _match_sdk_expected_key(network_name) or raw_network_key
    storage_key = expected_key or raw_network_key
    block = device_state.get(storage_key)
    if not block:
        expected = sdk_check_expected_map.get(expected_key, {})
        block = {
            "display_name": network_name.strip() or expected.get("display_name", network_name.strip()),
            "sdk_version": "",
            "adapter_version": "",
            "native_version": "",
            "observed_network_name": "",
            "adapter_missing": False,
            "verification": "",
            "expected_key": expected_key,
            "updated_at": time.time(),
        }
        if expected.get("display_name"):
            block["display_name"] = expected.get("display_name")
        device_state[storage_key] = block
    else:
        expected_key = block.get("expected_key") or _match_sdk_expected_key(network_name) or raw_network_key
        expected = sdk_check_expected_map.get(expected_key, {})
        block["display_name"] = expected.get("display_name") or network_name.strip() or block["display_name"] or expected.get("display_name", network_name.strip())
        block["expected_key"] = expected_key
        block["updated_at"] = time.time()
    return block


def _process_sdk_cloudx_adapter_metadata_line(line, device_id):
    """Read Cloudx adapter metadata without using the IntegrationHelper state machine."""
    if active_platform != "android":
        return False

    match = SDK_CLOUDX_ADAPTER_METADATA_PATTERN.search(str(line or ""))
    if not match:
        return False

    actual_network = match.group(1).strip()
    adapter_version = match.group(2).strip()
    native_match = SDK_CLOUDX_NATIVE_METADATA_PATTERN.search(str(line or ""))
    expected_key = _match_cloudx_sdk_expected_key(actual_network)
    if not expected_key:
        return False

    expected = sdk_check_expected_map.get(expected_key, {})
    display_name = expected.get("display_name") or actual_network
    block = _sdk_check_block_for_device(device_id, display_name)
    changed = (
        block.get("adapter_version") != adapter_version
        or block.get("adapter_missing")
        or block.get("expected_key") != expected_key
    )
    block["adapter_version"] = adapter_version
    block["adapter_missing"] = False
    if expected.get("source") == "cloudx" and native_match:
        native_version = native_match.group(1).strip()
        if block.get("native_version") != native_version:
            changed = True
        block["native_version"] = native_version
    block["expected_key"] = expected_key
    block["display_name"] = display_name
    block["updated_at"] = time.time()
    return changed


def _process_sdk_cloudx_http_metadata_line(line, device_id):
    """Read the exact CloudX SDK telemetry entry without touching - Cloudx rows."""
    if active_platform != "android":
        return False
    raw_line = str(line or "")
    if not SDK_CLOUDX_HTTP_METADATA_PATTERN.search(raw_line):
        return False

    expected_key = _normalize_sdk_network_name("CloudX")
    expected = sdk_check_expected_map.get(expected_key, {})
    if not expected or expected.get("source") == "cloudx":
        return False

    plugin_version = _extract_json_field_version(raw_line, "pluginVersion")
    native_version = _extract_json_field_version(raw_line, "sdkVersion")
    if not plugin_version and not native_version:
        return False

    adapter_version = _extract_sdk_comparable_version(plugin_version, expected.get("adapter")) if plugin_version else ""
    sdk_version = _extract_sdk_comparable_version(native_version, expected.get("sdk")) if native_version else ""
    if not adapter_version and not sdk_version:
        return False

    display_name = expected.get("display_name") or "CloudX"
    block = _sdk_check_block_for_device(device_id, display_name)
    changed = False
    if adapter_version and (
        block.get("adapter_version") != adapter_version or block.get("adapter_missing")
    ):
        block["adapter_version"] = adapter_version
        block["adapter_missing"] = False
        changed = True
    if sdk_version and block.get("sdk_version") != sdk_version:
        block["sdk_version"] = sdk_version
        changed = True
    block["expected_key"] = expected_key
    block["display_name"] = display_name
    block["updated_at"] = time.time()
    return changed


# --- THREADS & PROCESSES ---

def _process_sdk_check_line(line, device_id):
    if is_paused or not sdk_check_active:
        return

    changed = False
    with lock:
        if _process_sdk_max_stream_line(line, device_id):
            changed = True
        if _process_sdk_cloudx_http_metadata_line(line, device_id):
            changed = True
        if _process_sdk_cloudx_adapter_metadata_line(line, device_id):
            changed = True

        if "IntegrationHelper" in line:
            current_network = sdk_check_current_network.get(device_id, "")
            header_match = SDK_HEADER_PATTERN.search(line)
            if header_match:
                current_network = header_match.group(1).strip()
                _sdk_check_block_for_device(device_id, current_network)
                sdk_check_current_network[device_id] = current_network
                changed = True
            else:
                sdk_match = SDK_VERSION_LINE_PATTERN.search(line) or IOS_SDK_VERSION_LINE_PATTERN.search(line)
                adapter_match = SDK_ADAPTER_VERSION_LINE_PATTERN.search(line) or IOS_ADAPTER_VERSION_LINE_PATTERN.search(line)
                adapter_missing_match = SDK_ADAPTER_MISSING_PATTERN.search(line) or IOS_ADAPTER_MISSING_PATTERN.search(line)
                verification_match = SDK_VERIFICATION_PATTERN.search(line) or IOS_VERIFICATION_PATTERN.search(line)

                if verification_match:
                    current_network = verification_match.group(1).strip()
                    block = _sdk_check_block_for_device(device_id, current_network)
                    status = verification_match.group(2).strip().upper()
                    if block.get("verification") != status:
                        block["verification"] = status
                        changed = True
                    sdk_check_current_network[device_id] = current_network

                target_network = current_network
                if target_network:
                    block = _sdk_check_block_for_device(device_id, target_network)
                    block_expected_key = block.get("expected_key") or _match_sdk_expected_key(target_network) or _normalize_sdk_network_name(target_network)
                    if block_expected_key != "firebasecrashlytics":
                        if sdk_match:
                            sdk_version = sdk_match.group(1).strip()
                            if block.get("sdk_version") != sdk_version:
                                block["sdk_version"] = sdk_version
                                changed = True
                        if adapter_match:
                            adapter_version = adapter_match.group(1).strip()
                            if block.get("adapter_version") != adapter_version or block.get("adapter_missing"):
                                block["adapter_version"] = adapter_version
                                block["adapter_missing"] = False
                                changed = True
                        elif adapter_missing_match:
                            if not block.get("adapter_missing") or block.get("adapter_version"):
                                block["adapter_missing"] = True
                                block["adapter_version"] = ""
                                changed = True

        if _process_sdk_external_line(line, device_id):
            changed = True

    if changed:
        _emit_sdk_check_results()


ANDROID_LOGCAT_LOOKBACK_SECONDS = 1.0


def _android_logcat_start_cursor():
    """Return a logcat cursor just before opening a continuous stream.

    ``adb logcat`` otherwise replays the device ring buffer every time its
    subprocess is recreated.  A one-second lookback covers the tiny startup
    race without replaying minutes of old events after a reconnect.
    """
    cursor_time = time.time() - ANDROID_LOGCAT_LOOKBACK_SECONDS
    return time.strftime("%m-%d %H:%M:%S.000", time.localtime(cursor_time))


def _android_logcat_stream_command(device_id, pid=None):
    command = [
        ADB_EXECUTABLE,
        "-s",
        device_id,
        "logcat",
        "-v",
        "threadtime",
        "-T",
        _android_logcat_start_cursor(),
    ]
    if pid:
        command.append(f"--pid={pid}")
    return command


def adb_log_reader(device_id):
    print(f"INFO: Starting log reader for {device_id}")
    proc = None
    reader_thread = threading.current_thread()
    try:
        proc = subprocess.Popen(
            _android_logcat_stream_command(device_id),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            bufsize=1,
            text=True,
            encoding='utf-8',
            errors='ignore',
            creationflags=creation_flags,
        )
        
        for line in iter(proc.stdout.readline, ''):
            if not line: break
            if active_platform != "android":
                continue
            _process_android_runtime_line(line, device_id)

    except Exception as e:
        print(f"Error {device_id}: {e}")
    finally:
        _safe_runtime_call("Android final event buffer", _flush_android_event_log_buffer, device_id)
        if proc and proc.poll() is None:
            try:
                proc.terminate()
            except Exception:
                pass
        with lock:
            if active_log_readers.get(device_id) is reader_thread:
                active_log_readers.pop(device_id, None)

def ios_log_reader(device_id):
    print(f"INFO: Starting iOS log reader for {device_id}")
    proc = None
    reader_thread = threading.current_thread()
    try:
        cmd = _resolve_ios_syslog_command(device_id)
        if not cmd:
            print("WARNING: No iOS syslog command found. Install idevicesyslog or tidevice.")
            return
        print(f"INFO: iOS syslog command for {device_id}: {' '.join(cmd)}")
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='ignore',
            creationflags=creation_flags,
            # tidevice is a Python program: keep its pipe output unbuffered so
            # lines arrive as they are logged instead of in bursts.
            env=dict(os.environ, PYTHONUNBUFFERED="1"),
        )
        with lock:
            active_ios_log_processes[device_id] = proc
            active_ios_log_started_at[device_id] = time.time()
            active_ios_log_last_seen[device_id] = time.time()
            active_ios_log_commands[device_id] = os.path.basename(cmd[0])

        for raw_chunk in iter(proc.stdout.readline, ''):
            if not raw_chunk:
                break
            with lock:
                active_ios_log_last_seen[device_id] = time.time()
            if active_platform != "ios":
                continue
            # Package Log detects the actual stream framing. The executable
            # name alone cannot tell us whether the relay payload is newline
            # or NUL-delimited.
            _safe_runtime_call(
                "iOS Package Log stream",
                process_ios_package_log_stream_chunk,
                device_id,
                raw_chunk,
            )
            logical_lines = raw_chunk.replace("\x00", "\n").splitlines()

            for raw_line in logical_lines:
                if not raw_line:
                    continue
                _process_ios_runtime_line(raw_line, device_id)
    except Exception as e:
        print(f"iOS log reader error {device_id}: {e}")
    finally:
        _flush_ios_event_log_buffer(device_id)
        _safe_runtime_call("iOS Package Log NUL flush", _flush_ios_package_log_frame_buffer, device_id)
        _safe_runtime_call("iOS Package Log line flush", _flush_ios_package_log_buffer, device_id)
        with lock:
            # A stalled-reader restart must never let an old reader's finally
            # block remove the dictionaries belonging to its replacement.
            current_reader = active_ios_log_readers.get(device_id)
            current_process = active_ios_log_processes.get(device_id)
            owns_reader_state = current_reader is reader_thread
            owns_process_state = current_process is proc
            if owns_reader_state:
                active_ios_log_readers.pop(device_id, None)
                active_ios_log_started_at.pop(device_id, None)
                active_ios_log_last_seen.pop(device_id, None)
                active_ios_log_commands.pop(device_id, None)
            if owns_process_state:
                active_ios_log_processes.pop(device_id, None)
            if owns_reader_state or (current_reader is None and current_process is None):
                incomplete_ios_max_load_ads_logs.pop(device_id, None)
                incomplete_ios_event_logs.pop(device_id, None)
                active_ios_package_log_stream_modes.pop(device_id, None)
        if proc and proc.poll() is None:
            try:
                proc.terminate()
            except Exception:
                pass

def _stop_ios_log_reader(device_id):
    with lock:
        proc = active_ios_log_processes.get(device_id)
        thread = active_ios_log_readers.get(device_id)

    if proc and proc.poll() is None:
        try:
            proc.terminate()
            proc.wait(timeout=2)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass

    if thread and thread is not threading.current_thread():
        thread.join(timeout=2)

    with lock:
        # Keep a still-running old reader registered. device_manager will see
        # it on the next pass and retry the stop instead of starting a second
        # reader for the same device.
        if thread and thread.is_alive():
            return False
        if active_ios_log_readers.get(device_id) is thread:
            active_ios_log_readers.pop(device_id, None)
            active_ios_log_started_at.pop(device_id, None)
            active_ios_log_last_seen.pop(device_id, None)
            active_ios_log_commands.pop(device_id, None)
        if active_ios_log_processes.get(device_id) is proc:
            active_ios_log_processes.pop(device_id, None)
        ios_device_missing_polls.pop(device_id, None)
        if device_id not in active_ios_log_readers and device_id not in active_ios_log_processes:
            incomplete_ios_max_load_ads_logs.pop(device_id, None)
            incomplete_ios_event_logs.pop(device_id, None)
            active_ios_package_log_stream_modes.pop(device_id, None)
    return True

def device_manager():
    global connected_devices_info
    while True:
        platform = active_platform
        generation = platform_generation
        try:
            if platform == "ios":
                observed_ids = _observe_ios_device_ids()
                readers_to_stop = []
                stale_reader_ids = set()
                with lock:
                    if generation != platform_generation or active_platform != platform:
                        continue
                    known_ids = set(active_ios_log_readers) | set(active_ios_log_processes)
                    known_ids.update(
                        device.get("id", "")
                        for device in connected_devices_info
                        if device.get("id")
                    )
                    ids = _stabilize_ios_device_ids(observed_ids, known_ids)
                    # A streaming reader survives discovery misses; a reader
                    # gone from discovery and silent is stale; an alive but
                    # silent reader of a connected phone is restarted.
                    readers_to_stop, stale_reader_ids = _ios_readers_to_stop(observed_ids, time.time())
                    ids.difference_update(stale_reader_ids)
                    ids.update(set(active_ios_log_readers) - stale_reader_ids)
                    ids.update(set(active_ios_log_processes) - stale_reader_ids)

                for did, reason in readers_to_stop:
                    print(f"WARNING: Stopping iOS log reader for {did} ({reason})")
                    _stop_ios_log_reader(did)

                with lock:
                    if generation != platform_generation or active_platform != platform:
                        continue
                    # Only start readers for a freshly observed device. A
                    # grace-kept ID must remain visible, but must not spawn a
                    # new tidevice process while discovery is temporarily
                    # empty.
                    active_ios_reader_ids = set(active_ios_log_readers) | set(active_ios_log_processes)
                    for did in observed_ids - active_ios_reader_ids:
                        t = threading.Thread(target=ios_log_reader, args=(did,), daemon=True)
                        active_ios_log_readers[did] = t
                        t.start()

                    connected_devices_info = [_make_device_info(i, 'ios') for i in ids]
                if generation != platform_generation or active_platform != platform:
                    continue
                if connected_devices_info:
                    socketio.emit('device_status', {"platform": platform, "connected_devices": connected_devices_info})
                else:
                    socketio.emit('device_status', {
                        "platform": platform,
                        "connected_devices": [],
                        "message": _ios_device_status_message()
                    })
                _emit_sdk_check_results()
            else:
                with lock:
                    if generation != platform_generation or active_platform != platform:
                        continue
                    ios_readers_to_stop = set(active_ios_log_readers) | set(active_ios_log_processes)
                for did in ios_readers_to_stop:
                    _stop_ios_log_reader(did)

                if generation != platform_generation or active_platform != platform:
                    continue
                output = subprocess.run([ADB_EXECUTABLE, 'devices'], capture_output=True, text=True, creationflags=creation_flags).stdout
                ids = {l.split('\t')[0] for l in output.strip().split('\n')[1:] if '\tdevice' in l}
                
                with lock:
                    if generation != platform_generation or active_platform != platform:
                        continue
                    # A reader can exit while ADB still reports the device
                    # (for example after a transient logcat failure). Remove
                    # dead thread handles so the next pass starts a fresh
                    # reader instead of leaving every event tab stalled.
                    for did in list(active_log_readers.keys()):
                        reader = active_log_readers.get(did)
                        if did in ids and reader is not None and not reader.is_alive():
                            active_log_readers.pop(did, None)
                    for did in ids - set(active_log_readers.keys()):
                        t = threading.Thread(target=adb_log_reader, args=(did,), daemon=True)
                        active_log_readers[did] = t
                        t.start()
                    
                    for did in set(active_log_readers.keys()) - ids:
                        del active_log_readers[did]
                    
                    connected_devices_info = [_make_device_info(i, 'android') for i in ids]
                if generation != platform_generation or active_platform != platform:
                    continue
                if connected_devices_info:
                    socketio.emit('device_status', {"platform": platform, "connected_devices": connected_devices_info})
                else:
                    socketio.emit('device_status', {
                        "platform": platform,
                        "connected_devices": [],
                        "message": f"Waiting... (ADB: {ADB_EXECUTABLE})"
                    })
        except Exception as e:
            if generation == platform_generation and active_platform == platform:
                socketio.emit('device_status', {
                    "platform": platform,
                    "connected_devices": [],
                    "message": f"{'iOS' if platform == 'ios' else 'ADB'} error: {e}"
                })
        time.sleep(3)

def _append_package_log_row(device_id, raw_log, time_str="", time_display="", level="", tag="", message="", is_error=False,
                            pid="", tid=""):
    global package_log_seq
    device_name = ""
    raw_text = str(raw_log or "").strip()
    if len(raw_text) > MAX_PACKAGE_LOG_RECORD_CHARS:
        raw_text = raw_text[:MAX_PACKAGE_LOG_RECORD_CHARS] + "\n[Package Log record truncated]"
    message_text = str(message or "").strip()
    if len(message_text) > MAX_PACKAGE_LOG_RECORD_CHARS:
        message_text = message_text[:MAX_PACKAGE_LOG_RECORD_CHARS] + "\n[Package Log message truncated]"
    with lock:
        # A Package Log consumer can deliver one final line after Clear All
        # asks its subprocess to stop. Do not let that tail recreate rows
        # after the UI has already been cleared.
        if not target_package_name:
            return
        session_id = active_package_log_session_id
        if not _accept_exact_record("package", device_id, raw_text, session_id or ""):
            return
        device_name = get_device_name(device_id)
        # Assign `seq` inside the lock that also orders the cache, so seq
        # order always equals cache order.
        package_log_seq += 1
        seq = package_log_seq
        package_log_cache.append({
            'seq': seq,
            'stream': package_log_stream_id,
            'device_id': device_id,
            'device_name': device_name,
            'log': raw_text,
            'time': time_str,
            'time_display': time_display or time_str,
            'level': level,
            'tag': tag,
            'message': message_text,
            'pid': str(pid or ""),
            'tid': str(tid or ""),
            'timestamp': time.time(),
            'is_error': is_error
        })
        if session_id:
            # Keep the cache append and its database enqueue in the same
            # critical section so Clear All cannot drain the queue and then
            # receive a late row from the old session.
            _enqueue_package_log_db_row((
                session_id,
                time.time(),
                time_display or time_str,
                device_id,
                device_name,
                level,
                tag,
                message_text,
                raw_text,
                1 if is_error else 0,
                seq,
            ))
        return True

def _ios_package_log_matches_target(raw_log, bundle_search):
    """Match iOS package logs by the record's process header, not its body."""
    lines = [line.strip() for line in str(raw_log or "").splitlines() if line.strip()]
    if not lines:
        return False
    first_line = lines[0]
    search = str(bundle_search or "").casefold()
    if IOS_LOG_RECORD_START_PATTERN.match(first_line):
        process_match = re.match(
            rf'^{IOS_LOG_TIMESTAMP_PATTERN}\s+\S+\s+'
            r'([^\s(\[]+)(?:\([^)]*\))?\[\d+\]\s+<[^>]+>:',
            first_line,
        )
        if process_match:
            return search in process_match.group(1).casefold()
        return search in first_line.casefold()
    return search in str(raw_log or "").casefold()


def process_ios_package_log_line(log_obj, feed_max=False):
    # Package Log has its own Pause/Resume control in the UI.  The global
    # validation pause must not silently stop package capture.
    if active_platform != "ios":
        return
    raw_log = log_obj.get("raw_log", "")
    # Only NUL-framed records need a second MAX-parser entry point. Normal
    # newline streams are already sent to the MAX parser line-by-line by the
    # iOS reader; feeding the completed Package Log record again can duplicate
    # or merge unrelated records.
    if feed_max and raw_log:
        process_load_ads_max_log(raw_log, log_obj.get("device_id", ""))
    with lock:
        bundle_search = target_package_name
    if not bundle_search:
        return
    if not _ios_package_log_matches_target(raw_log, bundle_search):
        return

    time_str = ""
    time_display = ""
    match = IOS_LOG_TIMESTAMP_CAPTURE_PATTERN.match(raw_log.strip())
    if match:
        time_str = match.group("timestamp")
        time_display = time_str.split()[-1]
    level = ""
    level_match = re.search(r'<(Error|Fault|Warning|Notice|Debug|Info)>', raw_log, re.IGNORECASE)
    if level_match:
        level_name = level_match.group(1).lower()
        level = {"error": "E", "fault": "F", "warning": "W", "notice": "I", "debug": "D", "info": "I"}.get(level_name, level_name[:1].upper())
    is_error = level in {"E", "F"}
    tag = log_obj.get("tag", "")
    message = log_obj.get("message", raw_log).strip()
    pid = ""
    ios_process_match = re.match(
        rf'^{IOS_LOG_TIMESTAMP_PATTERN}\s+\S+\s+'
        r'([^\s(\[]+)(?:\([^)]*\))?\[(\d+)\]\s+<[^>]+>:\s*(.*)$',
        raw_log.strip(),
        re.DOTALL,
    )
    if ios_process_match:
        tag = ios_process_match.group(1).strip()
        pid = ios_process_match.group(2)
        message = ios_process_match.group(3).strip()
    _safe_runtime_call(
        "iOS Package Log row",
        _append_package_log_row,
        log_obj.get("device_id", ""),
        raw_log,
        time_str,
        time_display,
        level,
        tag,
        message,
        is_error,
        pid,
    )


def _flush_ios_package_log_buffer(device_id):
    """Flush one buffered iOS syslog record into Package Log only."""
    with lock:
        raw_record = active_ios_package_log_buffers.pop(device_id, "")
    if not raw_record or active_platform != "ios":
        return
    process_ios_package_log_line(_normalize_ios_log_line(raw_record, device_id))


def _flush_ios_package_log_frame_buffer(device_id):
    """Flush one NUL-framed tidevice record into Package Log only."""
    with lock:
        raw_record = active_ios_package_log_frame_buffers.pop(device_id, "")
    if not raw_record or active_platform != "ios":
        return
    process_ios_package_log_line(
        _normalize_ios_log_line(raw_record.strip("\r\n"), device_id),
        feed_max=True,
    )


def process_ios_package_log_stream_chunk(device_id, raw_chunk):
    """Consume iOS syslog in either newline or NUL-delimited form.

    tidevice forwards the syslog relay payload as-is, so its stdout is not
    guaranteed to be NUL-delimited. Detect the framing from the stream instead
    of assuming it from the executable name; otherwise every newline record is
    accumulated into one giant buffer and a single matching app line can pull
    kernel/daemon records into Package Log.
    """
    if active_platform != "ios":
        return
    if not raw_chunk:
        return

    switch_to_nul = False
    with lock:
        stream_mode = active_ios_package_log_stream_modes.get(device_id)
        if stream_mode is None:
            stream_mode = "nul" if "\x00" in raw_chunk else "line"
            active_ios_package_log_stream_modes[device_id] = stream_mode
        elif stream_mode == "line" and "\x00" in raw_chunk:
            # A relay may expose the first delimiter only after readline()
            # has already returned a few newline chunks. Finish the normal
            # record before switching modes rather than mixing buffers.
            active_ios_package_log_stream_modes[device_id] = "nul"
            stream_mode = "nul"
            switch_to_nul = True

    if switch_to_nul:
        _flush_ios_package_log_buffer(device_id)

    if stream_mode == "nul":
        with lock:
            buffered = active_ios_package_log_frame_buffers.get(device_id, "")
            combined = buffered + raw_chunk
            frames = combined.split("\x00")
            remainder = frames.pop()
            # A broken relay must not let an unterminated frame grow without
            # bound. A normal MAX/syslog record is far below this limit.
            active_ios_package_log_frame_buffers[device_id] = (
                remainder if len(remainder) <= MAX_IOS_PACKAGE_LOG_BUFFER_CHARS else ""
            )

        for frame in frames:
            frame = frame.strip("\r\n")
            if frame:
                process_ios_package_log_line(
                    _normalize_ios_log_line(frame, device_id),
                    feed_max=True,
                )
        return

    for raw_line in raw_chunk.splitlines():
        if raw_line:
            process_ios_package_log_stream_line(
                device_id,
                _normalize_ios_log_line(raw_line, device_id),
            )


def process_ios_package_log_stream_line(device_id, log_obj):
    """Keep iOS multiline records intact without changing other iOS tabs."""
    if active_platform != "ios":
        return
    raw_line = (log_obj.get("raw_log", "") or "").rstrip("\r\n")
    if not raw_line:
        return

    if IOS_LOG_RECORD_START_PATTERN.match(raw_line):
        _flush_ios_package_log_buffer(device_id)
        with lock:
            active_ios_package_log_buffers[device_id] = raw_line
        return

    with lock:
        current = active_ios_package_log_buffers.get(device_id)
        if current is not None:
            combined = f"{current}\n{raw_line}"
            active_ios_package_log_buffers[device_id] = (
                combined if len(combined) <= MAX_IOS_PACKAGE_LOG_BUFFER_CHARS else ""
            )
            return

    # Preserve the old behavior for unstructured/prelude lines that arrive
    # before the first timestamped syslog record.
    process_ios_package_log_line(log_obj)

def _append_package_log_row_from_consumer(device_id, logcat_process, *row_args):
    """Append only while this consumer still owns the live Package stream."""
    with lock:
        current_process = active_logcat_processes.get(device_id)
        if not target_package_name:
            return False
        # Direct parser tests do not register a process. In production a
        # registered replacement must always invalidate the old consumer.
        if current_process is not None and current_process is not logcat_process:
            return False
        return _append_package_log_row(device_id, *row_args)


def package_log_consumer(device_id, logcat_process):
    try:
        for line in iter(logcat_process.stdout.readline, ''):
            if not line: break
            with lock:
                current_process = active_logcat_processes.get(device_id)
                if not target_package_name:
                    break
                if current_process is not None and current_process is not logcat_process:
                    break
            is_error = bool(re.search(r'^\S+\s+\S+\s+\d+\s+\d+\s+[EF]\s', line))
            time_str = ""
            time_display = ""
            level = ""
            tag = ""
            pid = ""
            tid = ""
            message = line.strip()
            # Try to parse standard logcat format: MM-DD HH:MM:SS.mmm PID TID LEVEL TAG: message
            m = re.match(r'^(\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d+)\s+(\d+)\s+(\d+)\s+([A-Z])\s+([^:]+):\s*(.*)$', line.strip())
            if m:
                time_str = m.group(1)
                pid = m.group(2)
                tid = m.group(3)
                level = m.group(4)
                tag = m.group(5).strip()
                message = m.group(6).strip()
                # Prefer shorter time display to save width: HH:MM:SS.mmm
                if ' ' in time_str:
                    time_display = time_str.split(' ', 1)[1]
            _safe_runtime_call(
                "Android Package Log row",
                _append_package_log_row_from_consumer,
                device_id,
                logcat_process,
                line,
                time_str,
                time_display,
                level,
                tag,
                message,
                is_error,
                pid,
                tid,
            )
    except Exception as exc:
        print(f"WARNING: Android Package Log consumer {device_id} stopped: {exc}")
        try:
            if logcat_process.poll() is None:
                logcat_process.terminate()
        except Exception:
            pass

def package_pid_monitor():
    global active_package_pids
    while True:
        time.sleep(3)

        if active_platform != "android":
            with lock:
                for p in active_logcat_processes.values():
                    try:
                        p.terminate()
                    except Exception:
                        pass
                active_logcat_processes.clear()
                active_package_pids.clear()
            continue

        with lock:
            pkg = target_package_name
            devices_snapshot = list(connected_devices_info)

        if not pkg:
            with lock:
                for p in active_logcat_processes.values():
                    try:
                        p.terminate()
                    except Exception:
                        pass
                active_logcat_processes.clear()
                active_package_pids.clear()
            continue

        current_ids = {d['id'] for d in devices_snapshot}

        # Clean up disconnected devices
        with lock:
            for did in list(active_logcat_processes.keys()):
                if did not in current_ids:
                    try:
                        active_logcat_processes[did].terminate()
                    except Exception:
                        pass
                    active_logcat_processes.pop(did, None)
                    active_package_pids.pop(did, None)

        for device in devices_snapshot:
            did = device['id']
            try:
                res = subprocess.run([ADB_EXECUTABLE, '-s', did, 'shell', 'pidof', '-s', pkg],
                                     capture_output=True, text=True, creationflags=creation_flags)
                pid = res.stdout.strip()

                if not pid:
                    with lock:
                        if did in active_logcat_processes:
                            try:
                                active_logcat_processes[did].terminate()
                            except Exception:
                                pass
                            active_logcat_processes.pop(did, None)
                        active_package_pids.pop(did, None)
                    continue

                with lock:
                    proc = active_logcat_processes.get(did)
                    prev_pid = active_package_pids.get(did)

                restart = (pid != prev_pid) or (proc is None) or (proc.poll() is not None)

                if restart:
                    with lock:
                        if did in active_logcat_processes:
                            try:
                                active_logcat_processes[did].terminate()
                            except Exception:
                                pass
                        cmd = _android_logcat_stream_command(did, pid=pid)
                        proc = subprocess.Popen(
                            cmd,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT,
                            bufsize=1,
                            text=True,
                            encoding='utf-8',
                            errors='ignore',
                            creationflags=creation_flags,
                        )
                        active_logcat_processes[did] = proc
                        active_package_pids[did] = pid
                        threading.Thread(target=package_log_consumer, args=(did, proc), daemon=True).start()
            except Exception:
                with lock:
                    if did in active_logcat_processes:
                        try:
                            active_logcat_processes[did].terminate()
                        except Exception:
                            pass
                        active_logcat_processes.pop(did, None)
                    active_package_pids.pop(did, None)


def _package_log_ui_row(row):
    return {field: row.get(field, "") for field in PACKAGE_LOG_UI_FIELDS}


def _package_log_rows_after(since_seq):
    """Rows of the current stream with seq > since_seq, oldest first.

    The caller holds ``lock``.  The cache is ordered by seq, so this walks
    back only over the rows that are actually new.
    """
    rows = []
    for row in reversed(package_log_cache):
        if row.get('seq', 0) <= since_seq or row.get('stream') != package_log_stream_id:
            break
        rows.append(row)
    rows.reverse()
    return rows


def _emit_package_log_rows(rows, stream_id, to=None, gap=False, sync=False):
    """Send rows as ordered `package_log_delta` chunks; never skip a row."""
    chunk_size = max(1, PACKAGE_LOG_DELTA_CHUNK_ROWS)
    starts = range(0, len(rows), chunk_size) if rows else [0]
    for start in starts:
        chunk = rows[start:start + chunk_size]
        payload = {
            "stream": stream_id,
            "rows": [_package_log_ui_row(row) for row in chunk],
            "first_seq": chunk[0]['seq'] if chunk else 0,
            "last_seq": chunk[-1]['seq'] if chunk else 0,
            "gap": bool(gap and start == 0),
            "sync": bool(sync),
        }
        if to is not None:
            socketio.emit('package_log_delta', payload, to=to)
        else:
            socketio.emit('package_log_delta', payload)


def _bump_package_log_stream_locked():
    """Start a new UI stream.  The caller holds ``lock``."""
    global package_log_stream_id, package_log_last_emitted_seq
    package_log_stream_id += 1
    package_log_last_emitted_seq = package_log_seq
    return package_log_stream_id


def _emit_package_log_reset(stream_id=None):
    """Tell every client to drop its rows and follow the given stream."""
    with lock:
        if stream_id is None:
            stream_id = _bump_package_log_stream_locked()
        payload = {
            "stream": stream_id,
            "package_id": target_package_name,
            "session_id": active_package_log_session_id,
            "last_seq": package_log_seq,
        }
    socketio.emit('package_log_reset', payload)


def package_log_emitter():
    """Push only rows the clients have not received yet.

    Package Log has an independent UI pause; keep sending rows while the
    global event-validation pause is enabled.
    """
    while True:
        time.sleep(PACKAGE_LOG_DELTA_INTERVAL_SECONDS)
        _emit_pending_package_log_rows()
        _emit_package_log_db_status_if_changed()


_package_log_db_status_reported = (0, "")


def _emit_package_log_db_status_if_changed(force=False):
    """Tell the UI when history writes are failing or rows were given up."""
    global _package_log_db_status_reported
    status = (package_log_db_dropped_rows, package_log_db_last_error)
    if not force and status == _package_log_db_status_reported:
        return False
    _package_log_db_status_reported = status
    socketio.emit('package_log_db_status', {
        "dropped_rows": status[0],
        "error": status[1],
        "queued_rows": package_log_db_queue.qsize(),
    })
    return True


def _emit_pending_package_log_rows():
    """One emitter pass: send every row added since the previous pass."""
    global package_log_last_emitted_seq
    rows = []
    stream_id = 0
    with lock:
        if package_log_seq > package_log_last_emitted_seq:
            if target_package_name:
                rows = _package_log_rows_after(package_log_last_emitted_seq)
                stream_id = package_log_stream_id
            package_log_last_emitted_seq = package_log_seq
    if rows:
        _emit_package_log_rows(rows, stream_id)
    return len(rows)

# --- SOCKET HANDLERS ---
@socketio.on('change_tab')
def handle_change_tab(data): 
    # Sync logs for current tab on switch
    if data.get('tab_name') in {'Validator', 'DefaultAdEvents', 'Specific', 'AdRevenue'}:
        socketio.emit('profile_updated', _profile_payload(), to=request.sid)
    if data.get('tab_name') == 'LoadAds': socketio.emit('update_load_ads', list(load_ads_events))
    if data.get('tab_name') == 'LoadAdsExt': socketio.emit('update_load_ads_ext', list(load_ads_ext_events))
    if data.get('tab_name') == 'DefaultAdEvents': socketio.emit('update_default_ad_events', _default_ad_event_payload())
    if data.get('tab_name') == 'SdkCheck': _emit_sdk_check_results()
    if data.get('tab_name') == 'CallbackAd': socketio.emit('update_callback_ad_table', list(callback_ad_logs))
    if data.get('tab_name') == 'PriceRotation': socketio.emit('update_price_rotation_table', list(price_rotation_logs))

def _reset_runtime_for_platform_switch():
    global is_paused, validator_active, sdk_check_active, sdk_check_current_network, sdk_max_ios_pending_lines, sdk_max_ios_core_pending_lines
    global target_package_name, active_package_log_session_id
    global specific_event_name_filters, specific_event_params_filters
    global connected_devices_info, installation_id_state
    with lock:
        ios_readers_to_stop = set(active_ios_log_readers) | set(active_ios_log_processes)
    for device_id in ios_readers_to_stop:
        _stop_ios_log_reader(device_id)

    with lock:
        is_paused = False
        validator_active = False
        sdk_check_active = False
        sdk_check_search_list.clear()
        sdk_check_results.clear()
        sdk_check_input_list.clear()
        sdk_check_expected_map.clear()
        sdk_check_runtime_state.clear()
        sdk_check_current_network = {}
        sdk_max_ios_pending_lines = {}
        sdk_max_ios_core_pending_lines = {}
        sdk_check_expected_order.clear()
        ios_device_missing_polls.clear()

        load_ads_events.clear(); unique_load_ads.clear()
        load_ads_ext_events.clear(); unique_load_ads_ext.clear()
        validator_results.clear(); validator_pending.clear()
        default_ad_event_hits.clear()
        default_ad_event_simple_hits.clear()
        specific_event_name_filters = []
        specific_event_params_filters = []
        specific_event_results.clear(); specific_event_pending.clear(); event_log_cache.clear()
        adrevenue_logs.clear(); adrevenue_rendered.clear(); adrevenue_pending.clear(); _clear_live_table_pending(); adrevenue_log_cache.clear()
        callback_ad_logs.clear(); incomplete_impression_logs.clear(); incomplete_ios_adrevenue_logs.clear(); incomplete_ios_load_ads_ext_logs.clear(); incomplete_ios_max_load_ads_logs.clear(); incomplete_ios_event_logs.clear(); incomplete_android_event_logs.clear(); incomplete_adjust_adrevenue_logs.clear()
        price_rotation_logs.clear()
        package_log_cache.clear(); active_package_pids.clear(); active_ios_package_log_buffers.clear(); active_ios_package_log_frame_buffers.clear(); active_ios_package_log_stream_modes.clear()
        installation_id_state.clear()
        connected_devices_info = []
        target_package_name = ""
        _drain_package_log_db_queue()
        if active_package_log_session_id:
            _finish_package_log_session(active_package_log_session_id)
            active_package_log_session_id = None

        for state in recording_states.values():
            state["is_recording"] = False
            state["record_request_id"] = None

    socketio.emit('pause_status', {'is_paused': False})
    socketio.emit('validator_status', {'active': False})
    for tab, state in recording_states.items():
        socketio.emit('record_status', {
            "tab_name": tab,
            "is_recording": False,
            "current_sheet": state.get("current_sheet", "")
        })
    socketio.emit('update_load_ads', [])
    socketio.emit('update_load_ads_ext', [])
    socketio.emit('update_validator_table', [])
    socketio.emit('update_default_ad_events', _default_ad_event_payload())
    socketio.emit('update_specific_event_table', [])
    socketio.emit('update_adrevenue_table', [])
    socketio.emit('update_callback_ad_table', [])
    socketio.emit('update_price_rotation_table', [])
    _emit_package_log_reset()
    socketio.emit('device_status', {
        "platform": active_platform,
        "connected_devices": [],
        "message": _ios_device_status_message() if active_platform == "ios" else f"Waiting... (ADB: {ADB_EXECUTABLE})"
    })
    socketio.emit("installation_id_status", _build_installation_id_payload(""))
    socketio.emit('runtime_reset', {})
    _emit_sdk_check_results()

@socketio.on('set_platform')
def set_platform(data):
    global active_platform, platform_generation
    platform = (data or {}).get('platform', 'android')
    with lock:
        platform_changed = active_platform != ('ios' if platform == 'ios' else 'android')
        active_platform = 'ios' if platform == 'ios' else 'android'
        platform_generation += 1
    # A reconnect normally sends reset=false so the current platform's logs
    # survive.  If the client was holding a stale platform (for example after
    # a page reload while the backend was still on iOS), still perform the
    # platform transition reset; otherwise the old device list/readers can be
    # shown under the newly selected platform.
    if (data or {}).get('reset') or platform_changed:
        _reset_runtime_for_platform_switch()
    socketio.emit('platform_status', {'platform': active_platform})

@socketio.on('toggle_record')
def tr(data):
    tab_name = data.get('tab_name')
    if not tab_name or tab_name not in recording_states:
        return

    current_state = recording_states[tab_name]
    
    if not current_state["is_recording"]:
        name = str(data.get('sheet_name') or 'Log_Default').strip() or 'Log_Default'
        request_id = uuid.uuid4().hex
        # Start immediately. Sheet creation is best-effort and must never block the
        # Socket.IO handler while Apps Script is slow or unavailable.
        current_state.update({
            "is_recording": True,
            "current_sheet": name,
            "record_request_id": request_id,
        })
        socketio.emit('record_status', {
            "tab_name": tab_name,
            "is_recording": True,
            "current_sheet": name,
        })
        threading.Thread(
            target=_create_record_sheet_in_background,
            args=(tab_name, name, request_id),
            daemon=True,
        ).start()
    else:
        current_state["is_recording"] = False
        current_state["record_request_id"] = None
    
    # Emit status back with tab_name so UI knows which button to update
    if not current_state["is_recording"]:
        socketio.emit('record_status', {
            "tab_name": tab_name,
            "is_recording": False,
            "current_sheet": current_state["current_sheet"]
        })

@socketio.on('toggle_pause')
def tp(): 
    global is_paused
    is_paused = not is_paused
    socketio.emit('pause_status', {'is_paused': is_paused})

@socketio.on('clear_all_logs')
def cl():
    global sdk_check_active, sdk_check_current_network, sdk_max_ios_pending_lines, sdk_max_ios_core_pending_lines
    global target_package_name, active_package_log_session_id
    package_processes_to_stop = []
    with lock:
        load_ads_events.clear(); unique_load_ads.clear()
        load_ads_ext_events.clear(); unique_load_ads_ext.clear()
        validator_results.clear(); validator_pending.clear()
        default_ad_event_hits.clear()
        default_ad_event_simple_hits.clear()
        specific_event_results.clear(); specific_event_pending.clear(); event_log_cache.clear()
        adrevenue_logs.clear(); adrevenue_rendered.clear(); adrevenue_pending.clear(); _clear_live_table_pending(); callback_ad_logs.clear(); price_rotation_logs.clear()
        package_log_cache.clear()
        # Clean SDK check
        sdk_check_active = False
        sdk_check_results.clear()
        sdk_check_search_list.clear()
        sdk_check_input_list.clear()
        sdk_check_expected_map.clear()
        sdk_check_runtime_state.clear()
        sdk_check_current_network = {}
        sdk_check_expected_order.clear()
        sdk_max_ios_pending_lines = {}
        sdk_max_ios_core_pending_lines = {}
        incomplete_impression_logs.clear()
        incomplete_ios_adrevenue_logs.clear()
        incomplete_ios_load_ads_ext_logs.clear()
        incomplete_ios_max_load_ads_logs.clear()
        incomplete_ios_event_logs.clear()
        incomplete_android_event_logs.clear()
        incomplete_adjust_adrevenue_logs.clear()
        installation_id_state.clear()

        # Clear is also the end of the current Package Log capture. Stop the
        # Android logcat consumers and discard any iOS multiline/NUL frame so
        # their final buffered line cannot repopulate the just-cleared table.
        package_processes_to_stop = list(active_logcat_processes.values())
        active_logcat_processes.clear()
        active_package_pids.clear()
        active_ios_package_log_buffers.clear()
        active_ios_package_log_frame_buffers.clear()
        active_ios_package_log_stream_modes.clear()
        target_package_name = ""
        _drain_package_log_db_queue()
        if active_package_log_session_id:
            _finish_package_log_session(active_package_log_session_id)
            active_package_log_session_id = None

    for process in package_processes_to_stop:
        try:
            process.terminate()
        except Exception:
            pass
        
    socketio.emit('update_load_ads', [])
    socketio.emit('update_load_ads_ext', [])
    socketio.emit('update_validator_table', [])
    socketio.emit('update_default_ad_events', _default_ad_event_payload())
    socketio.emit('update_specific_event_table', [])
    socketio.emit('update_adrevenue_table', [])
    socketio.emit('update_callback_ad_table', [])
    socketio.emit('update_price_rotation_table', [])
    _emit_package_log_reset()
    socketio.emit("installation_id_status", _build_installation_id_payload(""))
    _emit_sdk_check_results()
    socketio.emit('clear_all_logs_complete', {})

@socketio.on('start_validation')
def val(p): 
    global required_params, validator_active
    required_params = p or []
    validator_active = True
    validator_results.clear(); validator_pending.clear()
    _clear_record_dedup("validator")
    socketio.emit('update_validator_table', [])
    socketio.emit('validator_status', {'active': True})

@socketio.on('stop_validation')
def stop_val():
    global validator_active
    validator_active = False
    socketio.emit('validator_status', {'active': False})

@socketio.on('update_specific_filter')
def usf(d):
    global specific_event_name_filters, specific_event_params_filters
    with lock:
        specific_event_name_filters = d.get('eventNames', [])
        specific_event_params_filters = d.get('params', [])
    _apply_specific_filter_and_emit()

@socketio.on('update_adrevenue_filter')
def uaf(_d=None):
    _apply_adrevenue_filter_and_emit()

@socketio.on('start_sdk_check')
def sdk_check(data):
    global sdk_check_search_list, sdk_check_results, sdk_check_input_list, sdk_check_active, sdk_check_expected_map, sdk_check_runtime_state, sdk_check_current_network, sdk_check_expected_order, sdk_max_ios_pending_lines, sdk_max_ios_core_pending_lines
    with lock:
        sdk_check_search_list = []
        sdk_check_results = {}
        sdk_check_input_list = []
        sdk_check_expected_map = {}
        sdk_check_runtime_state = {}
        sdk_check_current_network = {}
        sdk_check_expected_order = []
        sdk_max_ios_pending_lines = {}
        sdk_max_ios_core_pending_lines = {}
        lines = [line.rstrip('\r') for line in data.get('text', '').splitlines() if line.strip()]
        if lines:
            header_text = _normalize_sdk_network_name(lines[0])
            if header_text.startswith('adsnetwork') and ('adapter' in header_text or 'sdk' in header_text):
                for line in lines[1:]:
                    if _is_sdk_input_header_line(line):
                        continue
                    parsed = _parse_sdk_expected_line(line)
                    if not parsed:
                        continue
                    network = parsed["network"]
                    adapter = parsed["adapter"]
                    sdk = parsed["sdk"]
                    log_search = parsed["log_search"]
                    _register_sdk_expected(
                        network,
                        adapter=adapter,
                        sdk=sdk,
                        log_search=log_search,
                        source=parsed.get("source", ""),
                        match_network=parsed.get("match_network", ""),
                    )
            else:
                current_label = ""
                for line in lines:
                    if _is_sdk_input_header_line(line):
                        continue
                    search_item = _parse_sdk_search_pattern_line(line, current_label)
                    if search_item:
                        version = search_item.get("version", "")
                        field = search_item.get("field", "adapter")
                        adapter = version if field in {"adapter", "both"} else ""
                        sdk = version if field in {"sdk", "both"} else ""
                        _register_sdk_expected(
                            search_item["network"],
                            adapter=adapter,
                            sdk=sdk,
                            log_search=search_item["search_pattern"],
                            source=search_item.get("source", ""),
                            match_network=search_item.get("match_network", ""),
                        )
                        if search_item.get("source") != "cloudx":
                            sdk_check_search_list.append(search_item)
                        sdk_check_input_list.append(search_item)
                        continue

                    parsed = _parse_sdk_expected_line(line)
                    if parsed:
                        network = parsed["network"]
                        _register_sdk_expected(
                            network,
                            adapter=parsed["adapter"],
                            sdk=parsed["sdk"],
                            log_search=parsed["log_search"],
                            source=parsed.get("source", ""),
                            match_network=parsed.get("match_network", ""),
                        )
                        current_label = network
                        continue
                    current_label = line.strip()
                    sdk_check_input_list.append({"type": "label", "display_name": line})
        sdk_check_active = True
    _emit_sdk_check_results()


@socketio.on('stop_sdk_check')
def stop_sdk_check():
    global sdk_check_active, sdk_check_current_network, sdk_max_ios_pending_lines, sdk_max_ios_core_pending_lines
    with lock:
        sdk_check_active = False
        sdk_check_current_network = {}
        sdk_max_ios_pending_lines = {}
        sdk_max_ios_core_pending_lines = {}
    _emit_sdk_check_results()

@socketio.on('start_package_log')
def spl(d):
    global target_package_name, active_package_log_session_id
    pid = str((d or {}).get('package_id', '') or '').strip()
    for device_id in list(active_ios_package_log_frame_buffers):
        _flush_ios_package_log_frame_buffer(device_id)
    for device_id in list(active_ios_package_log_buffers):
        _flush_ios_package_log_buffer(device_id)
    with lock:
        # Do not let rows waiting behind a slow SQLite writer leak into the
        # next Package Log session. They belong to the capture being replaced.
        _drain_package_log_db_queue()
        if active_package_log_session_id:
            _finish_package_log_session(active_package_log_session_id)
            active_package_log_session_id = None
        target_package_name = pid
        if pid:
            active_package_log_session_id = _start_package_log_session(pid)
        package_log_cache.clear()
        active_ios_package_log_buffers.clear()
        active_ios_package_log_frame_buffers.clear()
        active_ios_package_log_stream_modes.clear()
        # Bump inside the same lock that sets the new target, so the first
        # rows of the new capture already belong to the new stream.
        stream_id = _bump_package_log_stream_locked()
    _emit_package_log_reset(stream_id)


@socketio.on('package_log_sync')
def handle_package_log_sync(data):
    """Replay rows a (re)connecting client missed, from the server cache.

    A client sends the stream it follows and the last seq it holds.  A
    different stream means it must start over; rows older than the cache
    window are reported as a gap (they remain in SQLite history).
    """
    data = data or {}
    try:
        since_seq = max(int(data.get('since_seq') or 0), 0)
    except (TypeError, ValueError):
        since_seq = 0
    try:
        limit = int(data.get('limit') or PACKAGE_LOG_SERVER_CACHE_ROWS)
    except (TypeError, ValueError):
        limit = PACKAGE_LOG_SERVER_CACHE_ROWS
    limit = max(1, min(limit, PACKAGE_LOG_SERVER_CACHE_ROWS))
    with lock:
        stream_id = package_log_stream_id
        same_stream = data.get('stream') == stream_id
        if not same_stream:
            since_seq = 0
        rows = _package_log_rows_after(since_seq)
        gap = bool(rows) and since_seq > 0 and rows[0]['seq'] > since_seq + 1
        if len(rows) > limit:
            rows = rows[-limit:]
            gap = gap or since_seq > 0
        reset_payload = {
            "stream": stream_id,
            "package_id": target_package_name,
            "session_id": active_package_log_session_id,
            "last_seq": package_log_seq,
        }
    if not same_stream:
        socketio.emit('package_log_reset', reset_payload, to=request.sid)
    _emit_package_log_rows(rows, stream_id, to=request.sid, gap=gap, sync=True)
    if package_log_db_dropped_rows or package_log_db_last_error:
        socketio.emit('package_log_db_status', {
            "dropped_rows": package_log_db_dropped_rows,
            "error": package_log_db_last_error,
            "queued_rows": package_log_db_queue.qsize(),
        }, to=request.sid)

@socketio.on('refresh_request')
def refresh():
    socketio.emit('update_load_ads', list(load_ads_events))
    socketio.emit('update_load_ads_ext', list(load_ads_ext_events))
    socketio.emit('update_default_ad_events', _default_ad_event_payload())
    socketio.emit('update_callback_ad_table', list(callback_ad_logs)) # Ensure callback data is refreshed
    socketio.emit('update_price_rotation_table', list(price_rotation_logs))
    _emit_specific_event_snapshot(to=request.sid)
    _emit_validator_snapshot(to=request.sid)
    _emit_live_table_snapshots(request.sid)
    # ... trigger others ...

@socketio.on('connect')
def connect(): 
    with lock:
        default_ad_event_clients.add(request.sid)
    _emit_specific_event_snapshot(to=request.sid)
    _emit_validator_snapshot(to=request.sid)
    _emit_live_table_snapshots(request.sid)
    socketio.emit('pause_status', {'is_paused': is_paused})
    socketio.emit('validator_status', {'active': validator_active})
    socketio.emit('platform_status', {'platform': active_platform})
    socketio.emit('profile_updated', _profile_payload(), to=request.sid)
    socketio.emit('update_default_ad_events', _default_ad_event_payload())
    socketio.emit('update_price_rotation_table', list(price_rotation_logs))
    _emit_installation_id_state()
    # Sync recording buttons on connect
    for tab, state in recording_states.items():
        socketio.emit('record_status', {
            "tab_name": tab,
            "is_recording": state["is_recording"],
            "current_sheet": state["current_sheet"]
        })

@socketio.on('disconnect')
def disconnect():
    with lock:
        default_ad_event_clients.discard(request.sid)
        last_client_disconnected = not default_ad_event_clients
    if last_client_disconnected:
        _clear_default_ad_event_history()

# --- MAIN ---
def run_server(host="0.0.0.0", port=5008):
    _clear_default_ad_event_history()
    _normalize_remote_update_config()
    threading.Thread(target=_record_app_open_audit_once, daemon=True).start()
    _set_active_profile()
    _init_package_log_db()
    threading.Thread(target=device_manager, daemon=True).start()
    threading.Thread(target=android_installation_id_monitor, daemon=True).start()
    threading.Thread(target=package_log_emitter, daemon=True).start()
    threading.Thread(target=live_table_emitter, daemon=True).start()
    threading.Thread(target=package_pid_monitor, daemon=True).start()
    threading.Thread(target=_package_log_db_writer, daemon=True).start()
    threading.Thread(target=_package_log_maintenance_loop, daemon=True).start()

    def safe_print(msg):
        try:
            print(msg)
        except UnicodeEncodeError:
            try:
                print(msg.encode("utf-8", "ignore").decode("utf-8"))
            except Exception:
                logging.info(msg)

    safe_print("===================================================")
    safe_print("  Event Inspector Started")
    safe_print(f"  - Access: http://{host}:{port}")
    safe_print("===================================================")
    socketio.run(
        app,
        host=host,
        port=port,
        use_reloader=False,
        allow_unsafe_werkzeug=True
    )

if __name__ == '__main__':
    run_server(host="0.0.0.0", port=5008)
