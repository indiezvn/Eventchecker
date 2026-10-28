@echo off
setlocal

rem Official Windows update reset for the v2.5.0 channel.
rem Use this only while Event Inspector is fully closed.

set "APP_SUPPORT_DIR=%LOCALAPPDATA%\EventInspector"
set "MANIFEST_URL=https://raw.githubusercontent.com/indiezvn/Eventchecker/main/Updates_2_5/remote_manifest.json"
set "TARGET_VERSION=2026-10-28-8-2.5.0-90"
set "UPDATES_DIR_V250=%APP_SUPPORT_DIR%\updates_v250"

echo Event Inspector update reset (Windows)
echo Target release: v2.5.0(90)
echo Target folder: %APP_SUPPORT_DIR%
echo.
echo Make sure Event Inspector is fully closed before continuing.
echo.

if not exist "%APP_SUPPORT_DIR%" mkdir "%APP_SUPPORT_DIR%"

for %%C in (v250) do (
  if exist "%APP_SUPPORT_DIR%\update_state_%%C.json" (
    del /f /q "%APP_SUPPORT_DIR%\update_state_%%C.json"
    echo Removed: %APP_SUPPORT_DIR%\update_state_%%C.json
  ) else (
    echo Skip (not found): %APP_SUPPORT_DIR%\update_state_%%C.json
  )
  if exist "%APP_SUPPORT_DIR%\remote_update_config_%%C.json" (
    del /f /q "%APP_SUPPORT_DIR%\remote_update_config_%%C.json"
    echo Removed: %APP_SUPPORT_DIR%\remote_update_config_%%C.json
  ) else (
    echo Skip (not found): %APP_SUPPORT_DIR%\remote_update_config_%%C.json
  )
  if exist "%APP_SUPPORT_DIR%\updates_%%C" (
    rmdir /s /q "%APP_SUPPORT_DIR%\updates_%%C"
    echo Removed: %APP_SUPPORT_DIR%\updates_%%C
  ) else (
    echo Skip (not found): %APP_SUPPORT_DIR%\updates_%%C
  )
  if exist "%APP_SUPPORT_DIR%\updates_%%C_tmp" (
    rmdir /s /q "%APP_SUPPORT_DIR%\updates_%%C_tmp"
    echo Removed: %APP_SUPPORT_DIR%\updates_%%C_tmp
  ) else (
    echo Skip (not found): %APP_SUPPORT_DIR%\updates_%%C_tmp
  )
)

(
  echo {
  echo   "enabled": true,
  echo   "manifest_url": "%MANIFEST_URL%",
  echo   "manifest_urls": [
  echo     "%MANIFEST_URL%",
  echo     "https://github.com/indiezvn/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
  echo     "https://raw.githubusercontent.com/trucbm/Eventchecker/main/Updates_2_5/remote_manifest.json",
  echo     "https://github.com/trucbm/Eventchecker/raw/main/Updates_2_5/remote_manifest.json",
  echo     "https://cdn.jsdelivr.net/gh/indiezvn/Eventchecker@main/Updates_2_5/remote_manifest.json",
  echo     "https://cdn.jsdelivr.net/gh/trucbm/Eventchecker@main/Updates_2_5/remote_manifest.json"
  echo   ],
  echo   "timeout_sec": 15,
  echo   "min_interval_sec": 0
  echo }
) > "%APP_SUPPORT_DIR%\remote_update_config_v250.json"
echo Wrote: %APP_SUPPORT_DIR%\remote_update_config_v250.json

(
  echo {
  echo   "last_check": 0,
  echo   "version": "%TARGET_VERSION%",
  echo   "update_dir": "%UPDATES_DIR_V250:\=\\%",
  echo   "manifest_url": "%MANIFEST_URL%",
  echo   "files": [
  echo     "Log_checker.py",
  echo     "remote_update.py",
  echo     "remote_update_config_v250.json",
  echo     "services_checker/app.py",
  echo     "services_checker/axml_fallback.py",
  echo     "services_checker/bundletool-all-1.18.1.jar"
  echo   ]
  echo }
) > "%APP_SUPPORT_DIR%\update_state_v250.json"
echo Wrote: %APP_SUPPORT_DIR%\update_state_v250.json

echo.
echo Done. Open Event Inspector and press Check Update once.
pause
exit /b 0
