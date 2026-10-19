@echo off
setlocal

REM Go to project root (two levels up from this script)
pushd "%~dp0\\..\\.."

REM Refuse to build from a stale branch or mixed-version source tree.
powershell -NoProfile -ExecutionPolicy Bypass -File "build\\windows\\verify_release_source.ps1"
if errorlevel 1 (
  echo Release source validation failed. Installer build stopped.
  popd
  exit /b 1
)

REM Derive the installer version from the validated source, never from a stale local environment.
set "EVENTINSPECTOR_RELEASE_VERSION="
for /f "usebackq delims=" %%V in (`powershell -NoProfile -ExecutionPolicy Bypass -File "build\windows\verify_release_source.ps1" -PrintVersion`) do set "EVENTINSPECTOR_RELEASE_VERSION=%%V"
if not defined EVENTINSPECTOR_RELEASE_VERSION (
  echo Could not resolve the release version from Log_checker.py.
  popd
  exit /b 1
)

REM Create venv
if not exist ".venv" (
  python -m venv .venv
)

call .venv\Scripts\activate.bat

REM Refuse to build an updater payload whose manifest hashes are stale.
python tools\verify_update_manifest.py
if errorlevel 1 (
  echo Remote update manifest validation failed. Installer build stopped.
  popd
  exit /b 1
)

REM Install deps
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM Clean old build output (in case files are locked)
if exist "dist\EventInspector" rmdir /s /q "dist\EventInspector"
if exist "build\EventInspector" rmdir /s /q "build\EventInspector"
if exist "build\windows\Output" rmdir /s /q "build\windows\Output"

REM Build EXE
python -m pip install pyinstaller PySide6 qtpy
python -m PyInstaller --noconfirm --clean --windowed --icon assets\\app.ico --name "EventInspector" ^
  --collect-submodules "engineio" ^
  --collect-submodules "socketio" ^
  --collect-submodules "webview" ^
  --collect-submodules "androguard.core" ^
  --collect-data "androguard" ^
  --hidden-import "androguard.core.axml" ^
  --hidden-import "qtpy" ^
  --hidden-import "qtpy.QtCore" ^
  --hidden-import "qtpy.QtGui" ^
  --hidden-import "qtpy.QtWidgets" ^
  --hidden-import "qtpy.QtNetwork" ^
  --hidden-import "qtpy.QtWebChannel" ^
  --hidden-import "qtpy.QtWebEngineCore" ^
  --hidden-import "qtpy.QtWebEngineWidgets" ^
  --hidden-import "PySide6.QtCore" ^
  --hidden-import "PySide6.QtGui" ^
  --hidden-import "PySide6.QtWidgets" ^
  --hidden-import "PySide6.QtNetwork" ^
  --hidden-import "PySide6.QtWebChannel" ^
  --hidden-import "PySide6.QtWebEngineCore" ^
  --hidden-import "PySide6.QtWebEngineWidgets" ^
  --hidden-import "shiboken6" ^
  --add-data "Log_checker.py;." ^
  --add-data "remote_update.py;." ^
  --add-data "Default event + Default Params.xlsx;." ^
  --add-data "remote_update_config_v250.json;." ^
  --add-data "services_checker\app.py;services_checker" ^
  --add-data "services_checker\axml_fallback.py;services_checker" ^
  --add-data "services_checker\bundletool-all-1.18.1.jar;services_checker" ^
  --add-data "services_checker\my-key.keystore;services_checker" ^
  desktop_app.py

if errorlevel 1 (
  echo PyInstaller installer build failed.
  popd
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "build\windows\verify_bundle.ps1" -BundleRoot "dist\EventInspector" -SourcePath "Log_checker.py"
if errorlevel 1 (
  echo Installer bundle validation failed.
  popd
  exit /b 1
)

REM Build installer (requires Inno Setup installed and ISCC on PATH)
if defined EVENTINSPECTOR_RELEASE_VERSION (
  ISCC /DMyAppVersion=%EVENTINSPECTOR_RELEASE_VERSION% build\windows\EventChecker.iss
) else (
  ISCC build\windows\EventChecker.iss
)
if errorlevel 1 (
  echo Inno Setup build failed.
  popd
  exit /b 1
)

endlocal
popd
