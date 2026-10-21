#!/bin/bash

set -u

APP_SUPPORT_DIR="${HOME}/Library/Application Support/EventInspector"
MANIFEST_URL="https://raw.githubusercontent.com/trucbm/Eventchecker/main/Updates_2_5/remote_manifest.json"
TARGET_VERSION="2026-10-21-9-2.5.0-82"

echo "Event Inspector update reset (macOS)"
echo "Target release: v2.5.0(82)"
echo "Target folder: ${APP_SUPPORT_DIR}"
echo
echo "Make sure Event Inspector is fully closed before continuing."
echo

mkdir -p "${APP_SUPPORT_DIR}"

for channel in v250; do
  state_file="${APP_SUPPORT_DIR}/update_state_${channel}.json"
  config_file="${APP_SUPPORT_DIR}/remote_update_config_${channel}.json"
  updates_dir="${APP_SUPPORT_DIR}/updates_${channel}"
  updates_tmp="${APP_SUPPORT_DIR}/updates_${channel}_tmp"

  if [ -f "${state_file}" ]; then
    rm -f "${state_file}"
    echo "Removed: ${state_file}"
  else
    echo "Skip (not found): ${state_file}"
  fi
  if [ -f "${config_file}" ]; then
    rm -f "${config_file}"
    echo "Removed: ${config_file}"
  else
    echo "Skip (not found): ${config_file}"
  fi
  if [ -d "${updates_dir}" ]; then
    rm -rf "${updates_dir}"
    echo "Removed: ${updates_dir}"
  else
    echo "Skip (not found): ${updates_dir}"
  fi
  if [ -d "${updates_tmp}" ]; then
    rm -rf "${updates_tmp}"
    echo "Removed: ${updates_tmp}"
  else
    echo "Skip (not found): ${updates_tmp}"
  fi
done

CONFIG_FILE_V250="${APP_SUPPORT_DIR}/remote_update_config_v250.json"
STATE_FILE_V250="${APP_SUPPORT_DIR}/update_state_v250.json"
UPDATES_DIR_V250="${APP_SUPPORT_DIR}/updates_v250"

cat > "${CONFIG_FILE_V250}" <<JSON
{
  "enabled": true,
  "manifest_url": "${MANIFEST_URL}",
  "manifest_urls": [
    "${MANIFEST_URL}",
    "https://github.com/trucbm/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
    "https://cdn.jsdelivr.net/gh/trucbm/Eventchecker@main/Updates_2_5/remote_manifest.json"
  ],
  "timeout_sec": 15,
  "min_interval_sec": 0
}
JSON
echo "Wrote: ${CONFIG_FILE_V250}"

cat > "${STATE_FILE_V250}" <<JSON
{
  "last_check": 0,
  "version": "${TARGET_VERSION}",
  "update_dir": "${UPDATES_DIR_V250}",
  "manifest_url": "${MANIFEST_URL}",
  "files": [
    "Log_checker.py",
    "remote_update.py",
    "remote_update_config_v250.json",
    "services_checker/app.py",
    "services_checker/axml_fallback.py",
    "services_checker/bundletool-all-1.18.1.jar"
  ]
}
JSON
echo "Wrote: ${STATE_FILE_V250}"

echo
echo "Done. Open Event Inspector and press Check Update once."
