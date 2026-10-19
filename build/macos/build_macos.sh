#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/../.."

# Read the release marker from the source that is being packaged.  This keeps
# the visible app version, bundle metadata, and update channel in sync.
APP_VERSION_LABEL="$(grep -Eo 'v[0-9]+\.[0-9]+\.[0-9]+\([0-9]+\)' Log_checker.py | head -n 1 || true)"
if [[ ! "$APP_VERSION_LABEL" =~ ^v[0-9]+\.[0-9]+\.[0-9]+\([0-9]+\)$ ]]; then
  echo "Could not find a release marker in Log_checker.py" >&2
  exit 1
fi
APP_VERSION="${APP_VERSION_LABEL#v}"
APP_SHORT_VERSION="${APP_VERSION%%(*}"
APP_BUILD_NUMBER="${APP_VERSION#*(}"
APP_BUILD_NUMBER="${APP_BUILD_NUMBER%)}"

# Create venv if missing
VENV_DIR="${VENV_DIR:-.venv}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

# Existing clients verify every downloaded payload against this manifest.
# Stop before creating an artifact if a source change was not synchronized.
"$PYTHON_BIN" tools/verify_update_manifest.py

if [ ! -d "$VENV_DIR" ]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

# Build for the host architecture by default.  A universal2 build requires
# universal native Python dependencies; callers can opt into another target.
if [ -z "${MACOS_TARGET_ARCH:-}" ]; then
  case "$(uname -m)" in
    arm64) MACOS_TARGET_ARCH="arm64" ;;
    x86_64) MACOS_TARGET_ARCH="x86_64" ;;
    *) echo "Unsupported macOS host architecture: $(uname -m)" >&2; exit 1 ;;
  esac
else
  MACOS_TARGET_ARCH="$MACOS_TARGET_ARCH"
fi

python -m pip install --upgrade pip
pip install -r requirements.txt

# Clean stale artifacts first so every build target is forced to pick up the
# current source instead of silently reusing old output.
rm -rf "dist/EventInspector.app"
rm -f "dist/EventInspector.dmg"
rm -rf "build/EventInspector"

# Build .icns from PNG if possible (requires macOS sips + iconutil)
PNG_SRC="/Users/truc.bui/Downloads/82690-protective-slitherio-personal-wallpaper-equipment-computer-snake.png"
ICON_DIR="assets/EventInspector.iconset"
ICNS_OUT="assets/app.icns"

if [ -f "$PNG_SRC" ]; then
  rm -rf "$ICON_DIR"
  mkdir -p "$ICON_DIR"

  sips -z 16 16     "$PNG_SRC" --out "$ICON_DIR/icon_16x16.png" >/dev/null
  sips -z 32 32     "$PNG_SRC" --out "$ICON_DIR/icon_16x16@2x.png" >/dev/null
  sips -z 32 32     "$PNG_SRC" --out "$ICON_DIR/icon_32x32.png" >/dev/null
  sips -z 64 64     "$PNG_SRC" --out "$ICON_DIR/icon_32x32@2x.png" >/dev/null
  sips -z 128 128   "$PNG_SRC" --out "$ICON_DIR/icon_128x128.png" >/dev/null
  sips -z 256 256   "$PNG_SRC" --out "$ICON_DIR/icon_128x128@2x.png" >/dev/null
  sips -z 256 256   "$PNG_SRC" --out "$ICON_DIR/icon_256x256.png" >/dev/null
  sips -z 512 512   "$PNG_SRC" --out "$ICON_DIR/icon_256x256@2x.png" >/dev/null
  sips -z 512 512   "$PNG_SRC" --out "$ICON_DIR/icon_512x512.png" >/dev/null
  sips -z 1024 1024 "$PNG_SRC" --out "$ICON_DIR/icon_512x512@2x.png" >/dev/null

  iconutil -c icns "$ICON_DIR" -o "$ICNS_OUT"
  rm -rf "$ICON_DIR"
fi

# Build .app
pip install pyinstaller
# MarkupSafe's optional C speedup is arm64-only in the current venv; the
# pure-Python fallback keeps the universal2 bundle architecture-neutral.
pyinstaller --noconfirm --clean --windowed \
  --target-arch "$MACOS_TARGET_ARCH" \
  --exclude-module "markupsafe._speedups" \
  --hidden-import "engineio.async_drivers.threading" \
  --collect-submodules "androguard.core" \
  --collect-data "androguard" \
  --hidden-import "androguard.core.axml" \
  --add-data "Log_checker.py:." \
  --add-data "Default event + Default Params.xlsx:." \
  --add-data "remote_update_config_v250.json:." \
  --add-data "services_checker/app.py:services_checker" \
  --add-data "services_checker/axml_fallback.py:services_checker" \
  --add-data "services_checker/bundletool-all-1.18.1.jar:services_checker" \
  --add-data "services_checker/my-key.keystore:services_checker" \
  --icon assets/app.icns --name "EventInspector" desktop_app.py

# Create DMG (simple)
APP_PATH="dist/EventInspector.app"
DMG_PATH="dist/EventInspector.dmg"
if [ -d "$APP_PATH" ]; then
  BUNDLETOOL_PATH="$(find "$APP_PATH/Contents" -type f -path "*/services_checker/bundletool-all-1.18.1.jar" -print -quit)"
  KEYSTORE_PATH="$(find "$APP_PATH/Contents" -type f -path "*/services_checker/my-key.keystore" -print -quit)"
  if [ -z "$BUNDLETOOL_PATH" ] || [ ! -s "$BUNDLETOOL_PATH" ]; then
    echo "Build validation failed: bundletool-all-1.18.1.jar is missing from the macOS app bundle." >&2
    exit 1
  fi
  if [ -z "$KEYSTORE_PATH" ] || [ ! -s "$KEYSTORE_PATH" ]; then
    echo "Build validation failed: my-key.keystore is missing from the macOS app bundle." >&2
    exit 1
  fi
  BUNDLETOOL_BYTES="$(stat -f%z "$BUNDLETOOL_PATH")"
  if [ "$BUNDLETOOL_BYTES" -lt 1000000 ]; then
    echo "Build validation failed: bundled bundletool is unexpectedly small (${BUNDLETOOL_BYTES} bytes)." >&2
    exit 1
  fi

  PLIST_PATH="$APP_PATH/Contents/Info.plist"
  /usr/libexec/PlistBuddy -c "Delete :CFBundleShortVersionString" "$PLIST_PATH" 2>/dev/null || true
  /usr/libexec/PlistBuddy -c "Add :CFBundleShortVersionString string $APP_SHORT_VERSION" "$PLIST_PATH"
  /usr/libexec/PlistBuddy -c "Delete :CFBundleVersion" "$PLIST_PATH" 2>/dev/null || true
  /usr/libexec/PlistBuddy -c "Add :CFBundleVersion string $APP_BUILD_NUMBER" "$PLIST_PATH"
  codesign --deep --force --sign - "$APP_PATH" >/dev/null

  ACTUAL_SHORT_VERSION="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$PLIST_PATH")"
  ACTUAL_BUILD_NUMBER="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleVersion' "$PLIST_PATH")"
  if [ "$ACTUAL_SHORT_VERSION" != "$APP_SHORT_VERSION" ] || [ "$ACTUAL_BUILD_NUMBER" != "$APP_BUILD_NUMBER" ]; then
    echo "Bundle metadata mismatch: expected $APP_SHORT_VERSION/$APP_BUILD_NUMBER, got $ACTUAL_SHORT_VERSION/$ACTUAL_BUILD_NUMBER" >&2
    exit 1
  fi

  hdiutil create -volname "Event Inspector $APP_VERSION_LABEL" -srcfolder "$APP_PATH" -ov -format UDZO "$DMG_PATH"
  echo "Built Event Inspector $APP_VERSION_LABEL for $MACOS_TARGET_ARCH: $DMG_PATH"
else
  echo "Build validation failed: $APP_PATH was not created." >&2
  exit 1
fi
