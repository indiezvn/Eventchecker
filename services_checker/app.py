import zipfile
import os
import json
import requests
# Flask import will be attempted below after sys.path printing
import secrets  # For session key
import re  # For version string parsing in APK scan and Gradle scan
import configparser  # To parse .properties files
import subprocess # For running bundletool
import shutil
import logging
import importlib
import importlib.util
import tempfile
import time
import hashlib
import threading
from pathlib import Path # For finding home directory
import xml.etree.ElementTree as ET # Still needed for the object structure
import sys # For printing sys.path for debugging
import traceback # For printing full tracebacks
from concurrent.futures import ThreadPoolExecutor

# --- Enhanced Debugging Output at the Start ---
print("--- Python sys.path (Interpreter's Search Path) ---")
for p in sys.path:
    print(p)
print("----------------------------------------------------")
print(f"--- Current Working Directory: {os.getcwd()} ---")
try:
    print(f"--- Script Location (__file__): {os.path.abspath(__file__)} ---")
except NameError:
    print("--- Script Location (__file__): Not defined (likely running in an interactive session where __file__ is not set) ---")
print("----------------------------------------------------")

# --- Try importing Flask first to ensure basic environment is okay ---
flask_available = False
try:
    from flask import Flask, request, render_template_string, send_file, jsonify, session
    flask_available = True
    print("Successfully imported Flask.")
except ImportError as e_flask:
    print(f"CRITICAL: Failed to import Flask. Error: {e_flask}")
    print("Please ensure Flask is installed in your active Python environment (e.g., pip install flask).")
    print("Exiting due to missing Flask.")
    sys.exit(1) # Exit if Flask is not available
except Exception as e_flask_other:
    print(f"CRITICAL: An unexpected error occurred while importing Flask: {e_flask_other}")
    traceback.print_exc()
    print("Exiting due to Flask import error.")
    sys.exit(1)


# --- Try importing the required library for binary XML (androguard) ---
androguard_axml_available = False
AXMLPrinter_class_from_androguard = None


def _load_fallback_axml_printer():
    """Load the dependency-free parser shipped with remote update payloads."""
    fallback_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "axml_fallback.py")
    if not os.path.isfile(fallback_path):
        return None

    module_name = f"eventinspector_axml_fallback_{os.getpid()}"
    spec = importlib.util.spec_from_file_location(module_name, fallback_path)
    if not spec or not spec.loader:
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    return getattr(module, "AXMLPrinter", None)


def _load_axml_printer():
    """Load AXMLPrinter from source, update, or PyInstaller bundle paths."""
    global androguard_axml_available, AXMLPrinter_class_from_androguard

    roots = [
        os.path.dirname(os.path.abspath(__file__)),
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    ]
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        roots.insert(0, meipass)
    for root in roots:
        if root and os.path.isdir(root) and root not in sys.path:
            sys.path.insert(0, root)

    print("--- Attempting to import AXMLPrinter from androguard.core.axml ---")
    try:
        importlib.invalidate_caches()
        axml_module = importlib.import_module("androguard.core.axml")
        printer = getattr(axml_module, "AXMLPrinter", None)
        if printer is None:
            raise ImportError("androguard.core.axml does not expose AXMLPrinter")
        AXMLPrinter_class_from_androguard = printer
        androguard_axml_available = True
        print("Successfully imported AXMLPrinter from androguard.core.axml.")
        return True
    except Exception as error:
        androguard_axml_available = False
        AXMLPrinter_class_from_androguard = None
        print(f"androguard AXMLPrinter unavailable: {error}")
        try:
            fallback_printer = _load_fallback_axml_printer()
        except Exception as fallback_error:
            print(f"ERROR: Dependency-free AXML fallback failed: {fallback_error}")
            traceback.print_exc()
            return False
        if fallback_printer is None:
            print("ERROR: Neither androguard nor the bundled AXML fallback is available.")
            return False
        AXMLPrinter_class_from_androguard = fallback_printer
        androguard_axml_available = True
        print("Using the bundled dependency-free AXML parser fallback.")
        return True


_load_axml_printer()
print("--------------------------------------------------------------------")
# --- End Import ---


# --- Configuration and Constants ---
UPLOAD_FOLDER_NAME = 'uploads'
APP_ROOT = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER_ABS_PATH = os.path.join(APP_ROOT, UPLOAD_FOLDER_NAME)

# Conversion output is local and single-user. Keep generated files out of the
# way of later bundletool runs, while never touching user-uploaded APK/AAB
# files. A new conversion supersedes the previous generated result, so old
# generated outputs can be removed before bundletool allocates more space.
MIN_CONVERSION_FREE_SPACE_BYTES = 600 * 1024 * 1024

if not os.path.exists(UPLOAD_FOLDER_ABS_PATH):
    os.makedirs(UPLOAD_FOLDER_ABS_PATH)


def _is_generated_conversion_artifact(filename):
    """Return whether *filename* is an APK/APKS created by this converter."""
    filename = str(filename or '')
    return bool(
        re.fullmatch(r'output_convert_[0-9a-f]{16}\.apks', filename, re.IGNORECASE)
        or re.fullmatch(r'temp_aab_convert_[0-9a-f]{16}_.+', filename, re.IGNORECASE)
        or re.fullmatch(r'.+_universal_[0-9a-f]{4}\.apk', filename, re.IGNORECASE)
        or filename == 'universal.apk'
    )


def _cleanup_stale_conversion_artifacts():
    """Remove previous converter artefacts without touching user uploads."""
    upload_root = os.path.abspath(app.config.get('UPLOAD_FOLDER', UPLOAD_FOLDER_ABS_PATH))
    removed = 0

    try:
        entries = os.scandir(upload_root)
    except OSError as error:
        logger.warning(f"Could not scan conversion upload directory '{upload_root}': {error}")
        return removed

    with entries:
        for entry in entries:
            if not entry.is_file(follow_symlinks=False):
                continue
            if not _is_generated_conversion_artifact(entry.name):
                continue

            path = os.path.abspath(entry.path)
            try:
                os.remove(path)
                removed += 1
            except OSError as error:
                logger.warning(f"Could not remove previous conversion artefact '{path}': {error}")

    if removed:
        logger.info(f"Removed {removed} previous AAB conversion artefact(s) before conversion.")
    return removed


def _remove_generated_file(path):
    """Delete one generated conversion file after a successful transfer."""
    if not path:
        return
    path = os.path.abspath(path)
    upload_root = os.path.abspath(app.config.get('UPLOAD_FOLDER', UPLOAD_FOLDER_ABS_PATH))
    try:
        inside_upload_folder = os.path.commonpath([upload_root, path]) == upload_root
    except ValueError:
        inside_upload_folder = False
    if not inside_upload_folder or not _is_generated_conversion_artifact(os.path.basename(path)):
        logger.warning(f"Refusing to remove non-generated conversion file: {path}")
        return
    try:
        if os.path.isfile(path):
            os.remove(path)
            logger.info(f"Removed generated conversion file after transfer: {path}")
    except OSError as error:
        logger.warning(f"Could not remove generated conversion file '{path}': {error}")


def _downloads_directory():
    """Return the current user's platform-native Downloads directory.

    The Services Checker process runs locally inside the desktop app, so the
    user's home directory is the same one used by the native platform.  Using
    ``expanduser`` keeps this path correct for macOS, Windows, and Linux
    without depending on the process working directory.
    """
    return os.path.join(os.path.expanduser('~'), 'Downloads')


def _safe_download_filename(filename, fallback='generated.apk'):
    """Return a single safe filename for a file written to Downloads."""
    candidate = str(filename or '').replace('\\', '/')
    candidate = os.path.basename(candidate).strip()
    # Keep filenames valid on Windows even when the same source AAB is
    # converted on macOS or Linux. This also prevents control characters from
    # reaching the native filesystem APIs.
    candidate = re.sub(r'[<>:"/|?*\x00-\x1f]', '_', candidate).rstrip(' .')
    if not candidate or candidate in {'.', '..'}:
        candidate = fallback
    if candidate.upper().split('.')[0] in {
        'CON', 'PRN', 'AUX', 'NUL',
        *(f'COM{index}' for index in range(1, 10)),
        *(f'LPT{index}' for index in range(1, 10)),
    }:
        candidate = f'_{candidate}'
    return candidate


def _open_unique_download_file(filename):
    """Reserve a new file in Downloads and return ``(path, handle, name)``.

    ``O_EXCL`` makes the collision check atomic across platforms.  The caller
    owns the returned binary handle and must close it; on write failure the
    partially written destination must be removed.
    """
    downloads_dir = _downloads_directory()
    os.makedirs(downloads_dir, exist_ok=True)
    safe_filename = _safe_download_filename(filename)
    stem, extension = os.path.splitext(safe_filename)
    binary_flag = getattr(os, 'O_BINARY', 0)

    for suffix in range(10000):
        if suffix == 0:
            candidate_name = safe_filename
        else:
            candidate_name = f'{stem} ({suffix}){extension}'
        destination = os.path.join(downloads_dir, candidate_name)
        try:
            descriptor = os.open(
                destination,
                os.O_WRONLY | os.O_CREAT | os.O_EXCL | binary_flag,
                0o644,
            )
        except FileExistsError:
            continue
        return destination, os.fdopen(descriptor, 'wb'), candidate_name

    raise FileExistsError(f'Could not reserve a unique filename in {downloads_dir}.')


def _remove_download_file(path):
    """Remove a partially written file owned by the current conversion."""
    if not path:
        return
    path = os.path.abspath(path)
    downloads_root = os.path.abspath(_downloads_directory())
    try:
        inside_downloads = os.path.commonpath([downloads_root, path]) == downloads_root
    except ValueError:
        inside_downloads = False
    if not inside_downloads:
        logger.warning(f"Refusing to remove file outside Downloads: {path}")
        return
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError as error:
        logger.warning(f"Could not remove partial Downloads file '{path}': {error}")


def _write_zip_member_to_downloads(apks_zip, member_name, filename):
    """Write one generated APK member directly to Downloads."""
    destination, output, saved_filename = _open_unique_download_file(filename)
    try:
        with apks_zip.open(member_name, 'r') as source, output:
            shutil.copyfileobj(source, output)
    except Exception:
        _remove_download_file(destination)
        raise
    return destination, saved_filename


def _copy_generated_file_to_downloads(source_path, filename):
    """Copy a generated APK to Downloads and return its final path.

    The converter is single-user, but filenames can collide with a previous
    result that the user kept. Preserve the existing file and choose a
    numbered name instead. The source is validated so this helper cannot be
    used to copy arbitrary uploaded files.
    """
    source_path = os.path.abspath(source_path)
    upload_root = os.path.abspath(app.config.get('UPLOAD_FOLDER', UPLOAD_FOLDER_ABS_PATH))
    try:
        inside_upload_folder = os.path.commonpath([upload_root, source_path]) == upload_root
    except ValueError:
        inside_upload_folder = False
    if not inside_upload_folder or not _is_generated_conversion_artifact(os.path.basename(source_path)):
        raise ValueError('Only generated APK files can be saved to Downloads.')
    if not os.path.isfile(source_path):
        raise FileNotFoundError('The generated APK is no longer available.')

    destination, output, _saved_filename = _open_unique_download_file(filename)
    try:
        with open(source_path, 'rb') as source, output:
            shutil.copyfileobj(source, output)
    except Exception:
        _remove_download_file(destination)
        raise
    logger.info(f"Saved generated file to Downloads: {destination}")
    return destination

def _service_resource_roots():
    """Resolve the platform-neutral resource layouts packaged with the app."""
    configured_root = os.getenv('EVENTINSPECTOR_BUNDLED_SERVICES_ROOT', '').strip()
    meipass = getattr(sys, '_MEIPASS', '') or ''
    executable = getattr(sys, 'executable', '') or ''
    executable_dir = os.path.dirname(os.path.abspath(executable)) if executable else ''
    roots = []
    if configured_root:
        roots.append(os.path.abspath(os.path.expanduser(configured_root)))
    roots.extend([
        APP_ROOT,
        os.path.join(meipass, 'services_checker') if meipass else '',
        meipass,
        os.path.join(executable_dir, 'services_checker') if executable_dir else '',
        os.path.join(executable_dir, '_internal', 'services_checker') if executable_dir else '',
        os.path.join(executable_dir, '..', 'Resources', 'services_checker') if executable_dir else '',
        os.path.join(executable_dir, '..', 'Frameworks', 'services_checker') if executable_dir else '',
    ])
    unique_roots = []
    seen = set()
    for root in roots:
        root = os.path.abspath(root) if root else ''
        if not root or root in seen:
            continue
        seen.add(root)
        unique_roots.append(root)
    return unique_roots


def _service_resource_path(filename):
    """Resolve a resource from the explicit artifact layouts."""
    filename = str(filename or '').strip()
    candidates = [os.path.join(root, filename) for root in _service_resource_roots()]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return candidates[0] if candidates else os.path.join(APP_ROOT, filename)

def _resolve_service_file(filename, env_var=None):
    """Resolve a bundled service file at call time, not only at import time."""
    configured = os.getenv(env_var, '').strip() if env_var else ''
    if configured:
        configured = os.path.expanduser(configured)
        configured_path = os.path.abspath(configured)
        if os.path.isfile(configured_path):
            return configured_path
        if not os.path.isabs(configured):
            configured_resource = _service_resource_path(configured)
            if os.path.isfile(configured_resource):
                return configured_resource
    return _service_resource_path(filename)


def _resolve_bundletool_path():
    """Find bundletool in the active source/update/frozen app resource layout."""
    return _resolve_service_file(
        'bundletool-all-1.18.1.jar',
        'EVENTINSPECTOR_BUNDLETOOL_PATH',
    )


# Kept for compatibility with diagnostics and older integrations. Conversion
# resolves this again at request time through _resolve_bundletool_path().
BUNDLETOOL_PATH = _resolve_bundletool_path()

# Paths to files within APK/AAB that indicate library presence/version
# Keys are paths in the archive, values are human-readable names for display.
TARGET_APK_FILE_PATHS = {
    "firebase-analytics.properties": "Firebase Analytics",
    "app-update.properties": "Google Play App Update",
    "review.properties": "Google Play In-App Review",
    "play-services-appset.properties": "GMS Play Services AppSet",
    "play-services-ads-identifier.properties": "GMS Play Services Ads Identifier",
    "asset-delivery.properties": "Google Play Asset Delivery",
    "user-messaging-platform.properties": "UMP SDK version", # Added as per user request
    "META-INF/androidx.recyclerview_recyclerview.version": "AndroidX RecyclerView",
    "META-INF/androidx.browser_browser.version": "AndroidX Browser",
    "META-INF/androidx.legacy_legacy-support-v4.version": "AndroidX Legacy Support v4",
    "META-INF/com.google.play.assetdelivery.version": "Google Play Asset Delivery (Unity)",
    "META-INF/com.android.installreferrer_installreferrer.version": "Install Referrer Library",
}

ANDROID_MANIFEST_FILENAME = "AndroidManifest.xml"
ANDROID_NAMESPACE_URI = 'http://schemas.android.com/apk/res/android'

# Metadata keys to extract from AndroidManifest.xml, mapped to display names
TARGET_METADATA_KEYS_MAP = {
    "com.facebook.sdk.ApplicationId": "Facebook ID",
    "com.facebook.sdk.ClientToken": "Facebook Client Token",
    "io.appmetrica.analytics.plugin_id": "Appmetrica Unity version"
}

# --- External library mappings ---
# The user-facing Gradle and Podfile mappings live in JSON files so adding or
# removing a library is a data refresh, not a Services Checker code release.
# The bundled JSON files remain the offline fallback when GitHub is unavailable.

# --- Hardcoded Keystore Information ---
# WARNING: Hardcoding passwords in source code is a security risk.
# This is for convenience in a controlled environment.
DEFAULT_KEYSTORE_FILENAME = "my-key.keystore"  # Assumed to be in the APP_ROOT directory
DEFAULT_KEYSTORE_ALIAS = "alias_name"
DEFAULT_KEYSTORE_PASS = "112233"
DEFAULT_KEY_PASS = "112233"
USE_HARDCODED_KEYSTORE = True # This is a non-production test key bundled for AAB conversion.


def _resolve_keystore_path():
    """Find the active bundled keystore, including frozen app resources."""
    return _resolve_service_file(DEFAULT_KEYSTORE_FILENAME, 'EVENTINSPECTOR_KEYSTORE_PATH')

# --- Flask App Initialization ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER_ABS_PATH
app.secret_key = secrets.token_hex(16)

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Log androguard_axml_available availability ---
if not androguard_axml_available:
    logger.warning("AXMLPrinter from 'androguard.core.axml' was not successfully imported. Manifest permission analysis will fail for binary XML manifests.")
    logger.warning("Please check the initial startup logs for detailed import errors and ensure 'androguard' is correctly installed and accessible in the Python environment being used.")
# --- End Log ---


# --- Helper Functions ---

def scan_podfile_for_versions(podfile_content, library_mapping):
    """
    Scans Podfile content for versions of specified libraries.
    Args:
        podfile_content (str): The string content of the Podfile.
        library_mapping (dict): A dictionary mapping user-friendly names to Pod names.
    Returns:
        dict: A dictionary mapping the library's user-friendly name to its found version.
    """
    found_versions = {}
    lines = podfile_content.splitlines()

    for lib_name, pod_name in library_mapping.items():
        # Pattern to find "pod 'PodName', 'version'"
        pattern = re.compile(r"pod\s+['\"]" + re.escape(pod_name) + r"['\"],\s+['\"]([^'\"]+)['\"]")

        for line in lines:
            # Skip commented out lines
            if line.strip().startswith('#'):
                continue

            match = pattern.search(line)
            if match:
                version = match.group(1).strip()
                # Remove CocoaPods operators like '~>' if present
                version = version.lstrip('~> ').strip()
                found_versions[lib_name] = version
                break  # Move to the next library once found

    return found_versions

def scan_gradle_for_versions(gradle_content, library_mapping):
    """
    Scans gradle file content for versions of specified libraries.
    Args:
        gradle_content (str): The string content of the build.gradle file.
        library_mapping (dict): A dictionary mapping user-friendly names to gradle artifact strings.
    Returns:
        dict: A dictionary mapping the library's user-friendly name to its found version.
    """
    found_versions = {}
    lines = gradle_content.splitlines()

    for lib_name, artifact_id in library_mapping.items():
        # Escape special regex characters in the artifact_id
        pattern = re.escape(artifact_id) + r":([^'\"]+)"

        for line in lines:
            # FIX: Check if the line is a comment and skip it if it is.
            if line.strip().startswith('//'):
                continue

            match = re.search(pattern, line)
            if match:
                version = match.group(1).strip()
                # Handle cases where version might be a variable like `versions.core`
                if not re.match(r'^\d', version):
                    # If it's a variable, we try to find its definition, e.g., `core = "1.9.0"`
                    var_name = version.split('.')[-1]
                    var_pattern = re.compile(r'{}\s*=\s*["\']([^"\']+)["\']'.format(re.escape(var_name)))
                    var_match = var_pattern.search(gradle_content)
                    if var_match:
                        version = var_match.group(1)

                found_versions[lib_name] = version
                break # Move to the next library once found

    return found_versions

def get_debug_keystore_path():
    """Attempts to find the default Android debug.keystore path."""
    home = Path.home()
    paths_to_check = [
        home / ".android" / "debug.keystore",
        Path(os.environ.get("USERPROFILE", "")) / ".android" / "debug.keystore" # Windows
    ]
    for path in paths_to_check:
        if path.exists():
            logger.info(f"Found debug.keystore at: {path}")
            return str(path)
    logger.warning("Default debug.keystore not found in common locations.")
    return None

def extract_binary_manifest_from_apk(apk_path, output_binary_filename="AndroidManifest_binary.xml"):
    """Extracts the binary AndroidManifest.xml from an APK."""
    output_binary_path = os.path.join(app.config['UPLOAD_FOLDER'], output_binary_filename)
    try:
        with zipfile.ZipFile(apk_path, 'r') as apk_zip:
            if ANDROID_MANIFEST_FILENAME in apk_zip.namelist():
                with open(output_binary_path, 'wb') as f:
                    f.write(apk_zip.read(ANDROID_MANIFEST_FILENAME))
                return output_binary_path
            else:
                logger.info(f"{ANDROID_MANIFEST_FILENAME} not found in APK for binary extraction.")
                return None
    except Exception as e:
        logger.error(f"Error extracting binary manifest: {e}")
        return None

def scan_archive_for_files(archive_path, target_files):
    """
    Scans a ZIP archive (APK/AAB) for specified target files and extracts their content or version.
    Args:
        archive_path (str): Path to the archive file.
        target_files (dict): Dictionary where keys are file paths within the archive
                             and values are human-readable library names.
    Returns:
        dict: A dictionary mapping library names to found versions or error messages.
    """
    found_versions = {}
    try:
        with zipfile.ZipFile(archive_path, 'r') as archive_zip:
            namelist = archive_zip.namelist()
            # Normalize target file paths (remove leading '/') for matching with zip namelist
            normalized_target_files = {path.lstrip('/'): name for path, name in target_files.items()}

            for target_path, lib_name in normalized_target_files.items():
                actual_path_to_read = None
                # Check if the exact normalized path exists
                if target_path in namelist:
                    actual_path_to_read = target_path
                # Some tools might store paths with a leading slash, zipfile usually doesn't list them like that
                # but we handle it just in case via lstrip above.

                if actual_path_to_read:
                    try:
                        file_content_bytes = archive_zip.read(actual_path_to_read)
                        file_content_str = file_content_bytes.decode('utf-8', errors='ignore').strip()
                        version = "File found, content unclear" # Default if no specific version pattern matches

                        if actual_path_to_read.endswith('.properties'):
                            config = configparser.ConfigParser()
                            # Handle properties files that don't start with a section header
                            if not file_content_str.startswith('['):
                                file_content_str = "[dummy_section]\n" + file_content_str
                            config.read_string(file_content_str)
                            if 'dummy_section' in config:
                                if 'version' in config['dummy_section']:
                                    version = config['dummy_section']['version']
                                elif 'pom.version' in config['dummy_section']: # Maven properties
                                    version = config['dummy_section']['pom.version']
                                elif config['dummy_section']: # Try to find any version-like string
                                    potential_version = None
                                    for key, val in config['dummy_section'].items():
                                        # Regex for common version patterns (e.g., 1.0, 2.3.4, 1.2.3-alpha)
                                        if re.match(r'^\d+\.\d+(\.\d+)?(-\S*)?$', val):
                                            potential_version = val
                                            break
                                    if potential_version: version = potential_version
                                    else: version = f"Properties found: {dict(config['dummy_section'])}" # Show all if no clear version
                            else: version = "Properties file empty"
                        elif actual_path_to_read.startswith('META-INF/') and actual_path_to_read.endswith('.version'):
                            if file_content_str: version = file_content_str
                            else: version = "File found, but empty"
                        # Add more specific parsers for other file types if needed

                        # Only add if we have a meaningful version or non-empty content
                        if version != "File found, content unclear" or file_content_str:
                            found_versions[lib_name] = version
                    except KeyError: # Should not happen if actual_path_to_read was confirmed in namelist
                        logger.warning(f"File '{actual_path_to_read}' listed but not found in archive (unexpected).")
                    except Exception as e:
                        logger.error(f"Error reading or parsing '{actual_path_to_read}' from archive: {e}")
                        found_versions[lib_name] = "Error reading/parsing file"
    except zipfile.BadZipFile:
        logger.error(f"Error: '{archive_path}' is not a valid ZIP archive (APK).")
        return {'error': 'Invalid file format. Not a valid APK.'}
    except FileNotFoundError:
        logger.error(f"Error: Archive file not found at '{archive_path}'")
        return {'error': 'Uploaded file not found on server.'}
    except Exception as e:
        logger.error(f"An unexpected error occurred scanning archive files: {e}")
        return {'error': f'An unexpected error occurred: {str(e)}'}
    return found_versions

def extract_manifest_data(apk_path):
    """
    Extracts permissions, specified metadata, and core manifest attributes from AndroidManifest.xml using androguard.
    Returns: (list_of_permissions, dict_of_metadata, dict_of_manifest_attributes, error_message_string_or_None)
    """
    permissions = []
    metadata = {}
    manifest_attributes = {}
    error_message = None

    if not androguard_axml_available or AXMLPrinter_class_from_androguard is None:
        # A dynamically loaded app.py can run before the frozen import hook is
        # ready. Retry once at request time instead of returning a false
        # "install androguard" error permanently.
        _load_axml_printer()
    if not androguard_axml_available or AXMLPrinter_class_from_androguard is None:
        logger.error("AXMLPrinter from androguard.core.axml was not available/accessible at script startup.")
        return [], {}, {}, "Manifest parsing requires AXMLPrinter from 'androguard', which was not found/accessible. Please install `androguard` and restart the application."

    try:
        with zipfile.ZipFile(apk_path, 'r') as apk_zip:
            if ANDROID_MANIFEST_FILENAME not in apk_zip.namelist():
                logger.error(f"'{ANDROID_MANIFEST_FILENAME}' not found in the APK: {apk_path}")
                return [], {}, {}, f"'{ANDROID_MANIFEST_FILENAME}' not found in the APK."

            try:
                manifest_content = apk_zip.read(ANDROID_MANIFEST_FILENAME)
                logger.debug(f"Read {len(manifest_content)} bytes for {ANDROID_MANIFEST_FILENAME} from {apk_path}")

                # Use the AXMLPrinter class obtained at startup
                printer = AXMLPrinter_class_from_androguard(manifest_content)
                xml_string = printer.get_buff()

                if not xml_string:
                    logger.error(f"androguard.core.axml.AXMLPrinter.get_buff() returned empty or None for {apk_path}.")
                    error_message = "Failed to convert binary XML to string using androguard (empty result)."
                    return permissions, metadata, manifest_attributes, error_message
                else:
                    logger.debug(f"Successfully got XML string from androguard AXMLPrinter for {apk_path} (length: {len(xml_string)})")

                # Parse the XML string using ElementTree
                xml_tree_root = ET.fromstring(xml_string)
                logger.debug(f"Successfully parsed XML string from androguard AXMLPrinter with ET for {apk_path}")

                # Extract manifest attributes
                manifest_attributes['Package Name'] = xml_tree_root.get('package', 'Not Found')
                manifest_attributes['Version Code'] = xml_tree_root.get(f'{{{ANDROID_NAMESPACE_URI}}}versionCode', 'Not Found')
                manifest_attributes['Version Name'] = xml_tree_root.get(f'{{{ANDROID_NAMESPACE_URI}}}versionName', 'Not Found')
                manifest_attributes['Compile SDK Version'] = xml_tree_root.get(f'{{{ANDROID_NAMESPACE_URI}}}compileSdkVersion', 'Not Found')
                manifest_attributes['Compile SDK Version Codename'] = xml_tree_root.get(f'{{{ANDROID_NAMESPACE_URI}}}compileSdkVersionCodename', 'Not Found')
                manifest_attributes['Platform Build Version Code'] = xml_tree_root.get('platformBuildVersionCode', xml_tree_root.get(f'{{{ANDROID_NAMESPACE_URI}}}platformBuildVersionCode','Not Found'))
                manifest_attributes['Platform Build Version Name'] = xml_tree_root.get('platformBuildVersionName', xml_tree_root.get(f'{{{ANDROID_NAMESPACE_URI}}}platformBuildVersionName','Not Found'))

                # Extract SDK versions
                uses_sdk_node = xml_tree_root.find('uses-sdk')
                if uses_sdk_node is not None:
                    manifest_attributes['Min SDK Version'] = uses_sdk_node.get(f'{{{ANDROID_NAMESPACE_URI}}}minSdkVersion', 'Not Found')
                    manifest_attributes['Target SDK Version'] = uses_sdk_node.get(f'{{{ANDROID_NAMESPACE_URI}}}targetSdkVersion', 'Not Found')
                else:
                    manifest_attributes['Min SDK Version'] = 'Not Found (<uses-sdk> missing)'
                    manifest_attributes['Target SDK Version'] = 'Not Found (<uses-sdk> missing)'
                    logger.warning(f"<uses-sdk> tag not found in manifest for {apk_path}.")

                # Extract permissions
                for permission_element in xml_tree_root.findall('.//uses-permission'):
                    permission_name = permission_element.get(f'{{{ANDROID_NAMESPACE_URI}}}name')
                    if permission_name:
                        permissions.append(permission_name)
                    else:
                        logger.warning(f"Found uses-permission tag without '{{{ANDROID_NAMESPACE_URI}}}name' attribute in {apk_path}.")
                logger.info(f"Extracted {len(permissions)} permissions from {apk_path} using androguard.")

                # Extract metadata from <application> tag
                application_node = xml_tree_root.find('application')
                if application_node is not None:
                    for meta_element in application_node.findall('meta-data'):
                        meta_name_actual = meta_element.get(f'{{{ANDROID_NAMESPACE_URI}}}name')
                        meta_value = meta_element.get(f'{{{ANDROID_NAMESPACE_URI}}}value')

                        # Check if this metadata key is one we're targeting
                        if meta_name_actual in TARGET_METADATA_KEYS_MAP:
                            display_name = TARGET_METADATA_KEYS_MAP[meta_name_actual]
                            if meta_value and meta_value.startswith('@'): # Resource ID
                                metadata[display_name] = f"Resource ID: {meta_value} (Actual value requires resource lookup)"
                            else:
                                metadata[display_name] = meta_value if meta_value is not None else "Not Set"
                            logger.info(f"Found metadata: {display_name} (from {meta_name_actual}) = {metadata[display_name]}")
                else:
                    logger.warning(f"<application> tag not found in manifest for {apk_path}.")
                    if not error_message: # Only set if no prior, more specific error occurred
                        error_message = "<application> tag not found in manifest."

            except ET.ParseError as et_pe: # If ElementTree fails to parse the string from AXMLPrinter
                logger.error(f"ET.ParseError after androguard AXMLPrinter for {apk_path}: {et_pe}. XML string snippet (up to 200 chars): '{str(xml_string)[:200]}...'")
                error_message = f"Failed to parse XML string (from androguard AXMLPrinter) with ElementTree: {str(et_pe)}. The manifest might be corrupted."
            except Exception as e: # Other errors during AXMLPrinter or ET processing
                logger.error(f"Error processing manifest with androguard AXMLPrinter/ET for {apk_path}: {e}", exc_info=True)
                error_message = f"An error occurred while parsing the manifest (androguard/ET stage): {str(e)}"

    except zipfile.BadZipFile:
        logger.error(f"BadZipFile: {apk_path} is not a valid APK.")
        error_message = "Invalid APK file format."
    except FileNotFoundError: # Should be caught by Flask if file not uploaded, but good to have
        logger.error(f"APK file not found at {apk_path} for manifest extraction.")
        error_message = "APK file disappeared before manifest extraction."
    except Exception as e:
        logger.error(f"General error in extract_manifest_data for {apk_path}: {e}", exc_info=True)
        # Avoid overwriting a more specific error_message from inner try-except
        if not error_message:
            error_message = f"An unexpected error occurred during manifest extraction: {str(e)}"

    return sorted(list(set(permissions))), metadata, manifest_attributes, error_message


# --- HTML Template ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Android App Tools</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
        .container h2.text-xl { font-size: 1.1rem !important; line-height: 1.35; }
        .container h3.text-lg { font-size: 1rem !important; line-height: 1.35; }
        .container > .mb-6 .tab-button { font-size: 0.8rem; padding: 0.45rem 0.85rem; }
        .container .error-message, .container .info-message, .container .warning-message { font-size: 0.9rem; }
        .container { max-width: 900px; margin: 2rem auto; padding: 1.5rem; }
        .card { background-color: white; border-radius: 0.5rem; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06); padding: 1.5rem; margin-bottom: 1.5rem; }
        .button { background-color: #3b82f6; color: white; padding: 0.75rem 1.5rem; border-radius: 0.375rem; border: none; cursor: pointer; font-weight: 500; transition: background-color 0.2s; display: inline-flex; align-items: center; justify-content: center; }
        .button:hover { background-color: #2563eb; }
        .button-secondary { background-color: #6b7280; }
        .button-secondary:hover { background-color: #4b5563; }
        .file-input { border: 1px solid #d1d5db; border-radius: 0.375rem; padding: 0.5rem; width: 100%; }
        .text-input { border: 1px solid #d1d5db; border-radius: 0.375rem; padding: 0.5rem; width: 100%; }
        .textarea { border: 1px solid #d1d5db; border-radius: 0.375rem; padding: 0.5rem; width: 100%; min-height: 100px; font-family: monospace; font-size: 0.9em; }
        .results-section h3, .results-section h4 { font-size: 1.25rem; font-weight: 600; margin-bottom: 0.75rem; color: #1f2937; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.5rem;}
        .results-section h4 { font-size: 1.1rem; border-bottom: none; margin-top: 1rem;}
        .results-section p, .results-section li { color: #374151; margin-bottom: 0.25rem; line-height: 1.6; }
        .results-section ul { list-style-type: disc; padding-left: 1.5rem; }
        .results-section .data-list ul { list-style-type: none; padding-left: 0; } /* Specific for metadata and attributes */
        .results-section .code { background-color: #f3f4f6; padding: 0.25rem 0.5rem; border-radius: 0.25rem; font-family: monospace; }
        #apk-scan-results-section > h2,
        #gradle-scan-results-section > h2,
        #podfile-scan-results-section > h2,
        #manifest-results-section > h2,
        #aab-conversion-results-section > h2 { font-size: 0.95rem !important; line-height: 1.25; }
        .results-section h3, .results-section h4 { font-size: 0.78rem !important; line-height: 1.25; padding-bottom: 0.25rem; margin-bottom: 0.35rem; }
        .results-section h5 { font-size: 0.7rem !important; line-height: 1.25; }
        .results-section p, .results-section li, .results-section .comparison-item { font-size: 0.68rem; line-height: 1.3; }
        .results-section .comparison-item strong { font-size: 0.72rem; }
        .results-section .code { font-size: 0.64rem; padding: 0.1rem 0.25rem; }
        .results-section .manifest-permission-item { font-size: 0.78rem; line-height: 1.4; }
        .results-section .manifest-permission-item strong { font-size: 0.82rem; }
        .results-section .manifest-permission-item .code { font-size: 0.74rem; }
        .error-message { color: #ef4444; background-color: #fee2e2; padding: 0.75rem; border-radius: 0.375rem; margin-bottom: 1rem; border: 1px solid #fca5a5; }
        .info-message { color: #059669; background-color: #d1fae5; padding: 0.75rem; border-radius: 0.375rem; margin-bottom: 1rem; border: 1px solid #6ee7b7;}
        .warning-message { color: #f59e0b; background-color: #fef3c7; padding: 0.75rem; border-radius: 0.375rem; margin-bottom: 1rem; border: 1px solid #fcd34d;}
        .status-passed { color: #16a34a; font-weight: bold; }
        .status-failed { color: #dc2626; font-weight: bold; }
        .status-strange { color: #d97706; font-weight: bold; }
        .status-info { color: #6b7280; font-weight: bold; }
        .comparison-item { margin-bottom: 0.5rem; padding-bottom: 0.5rem; border-bottom: 1px solid #e5e7eb; }
        .comparison-item:last-child { border-bottom: none; }
        .spinner { border: 4px solid rgba(0, 0, 0, 0.1); width: 36px; height: 36px; border-radius: 50%; border-left-color: #3b82f6; animation: spin 1s ease infinite; margin: 1rem auto; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .hidden { display: none; }
        .tab-button { padding: 0.5rem 1rem; margin-right: 0.5rem; border-radius: 0.375rem 0.375rem 0 0; cursor: pointer; background-color: #e5e7eb; color: #4b5563; border: 1px solid #d1d5db; border-bottom: none; font-size: 0.875rem; }
        .tab-button.active { background-color: white; color: #3b82f6; border-bottom: 1px solid white; position: relative; top: 1px; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .keystore-info-hardcoded { background-color: #eef2ff; border: 1px solid #c7d2fe; padding: 0.75rem; border-radius: 0.375rem; margin-bottom: 1rem; font-size: 0.9em; color: #4338ca; }
    </style>
</head>
<body class="bg-gray-100">
    <div class="container">
        <div class="mb-6 border-b border-gray-300">
            <button id="tab-apk-analyzer" class="tab-button active">Google Services & Meta-INF</button>
            <button id="tab-gradle-verifier" class="tab-button">Gradle Version Verifier</button>
            <button id="tab-podfile-verifier" class="tab-button">Podfile Version Verifier</button>
            <button id="tab-manifest-analyzer" class="tab-button">Permission & Manifest Data</button>
            <button id="tab-aab-converter" class="tab-button">AAB to APK Converter</button>
        </div>

        <div id="message-area"></div>

        <div id="apk-analyzer-content" class="tab-content active">
            <div class="card" id="apk-upload-section">
                <div class="flex items-center justify-between gap-4 flex-wrap mb-4">
                    <h2 class="text-xl font-semibold text-gray-700 mb-0">APK Library Version Analysis</h2>
                    <div class="flex items-center gap-3 flex-wrap">
                        <label for="apk-build-check-preset" class="text-sm font-semibold text-gray-700">Saved Check List:</label>
                        <select id="apk-build-check-preset" class="border border-gray-300 rounded-md px-3 py-2 text-sm bg-white min-w-[220px]">
                            <option value="">Manual input</option>
                        </select>
                        <button id="reload-apk-build-check-presets" type="button" class="button button-secondary text-sm py-2 px-3">Reload</button>
                    </div>
                </div>
                <form id="apk-upload-form" method="post" action="/upload_apk_for_scan" enctype="multipart/form-data">
                    <div class="mb-4">
                        <label for="apk_file_scan" class="block text-sm font-medium text-gray-700 mb-1">Select your APK file:</label>
                        <input type="file" name="apk_file" id="apk_file_scan" class="file-input" accept=".apk" required>
                    </div>
                    <div class="mb-4">
                        <label for="apk_expected_versions" class="block text-sm font-medium text-gray-700 mb-1">Input Expected Library Versions (Optional):</label>
                        <textarea id="apk_expected_versions" name="expected_versions" class="textarea" rows="3" placeholder="AndroidX RecyclerView      1.3.2&#10;Google Play App Update    2.1.0&#10;UMP SDK version           2.1.0&#10;... (Library Name <tabs/spaces> Version)"></textarea>
                    </div>
                    <div class="mb-4">
                        <a id="download-binary-manifest" href="#" class="button button-secondary mr-2 text-sm py-2 px-3 hidden" download="AndroidManifest_binary.xml">Download Binary Manifest</a>
                        <span id="no-binary-link" class="text-sm text-gray-500 hidden">Binary manifest could not be extracted.</span>
                    </div>
                    <button type="submit" class="button w-full">Analyze APK Libraries</button>
                </form>
                <div id="apk-spinner" class="spinner hidden"></div>
            </div>
            <div class="card hidden" id="apk-scan-results-section">
                <h2 class="text-xl font-semibold text-gray-700 mb-4">APK Scan & Comparison Results</h2>
                <div id="apk-scan-output"></div>
                <div id="apk-version-comparison-output" class="mt-4"></div>
                <button id="reset-apk-analyzer" class="button mt-6 w-full">Analyze Another APK (Libraries)</button>
            </div>
        </div>

        <div id="gradle-verifier-content" class="tab-content">
            <div class="card" id="gradle-upload-section">
                <div class="flex items-center justify-between gap-4 flex-wrap mb-4">
                    <h2 class="text-xl font-semibold text-gray-700 mb-0">Gradle Version Verifier</h2>
                    <div class="flex items-center gap-3 flex-wrap">
                        <label for="gradle-build-check-preset" class="text-sm font-semibold text-gray-700">Saved Check List:</label>
                        <select id="gradle-build-check-preset" class="border border-gray-300 rounded-md px-3 py-2 text-sm bg-white min-w-[220px]">
                            <option value="">Manual input</option>
                        </select>
                        <button id="reload-gradle-build-check-presets" type="button" class="button button-secondary text-sm py-2 px-3">Reload</button>
                    </div>
                </div>
                <form id="gradle-upload-form" method="post" action="/analyze_gradle" enctype="multipart/form-data">
                    <div class="mb-4">
                        <label for="gradle_file_scan" class="block text-sm font-medium text-gray-700 mb-1">Select your build.gradle file:</label>
                        <input type="file" name="gradle_file" id="gradle_file_scan" class="file-input" accept=".gradle" required>
                    </div>
                    <div class="mb-4">
                        <label for="gradle_expected_versions" class="block text-sm font-medium text-gray-700 mb-1">Input Expected Library Versions (Optional):</label>
                        <textarea id="gradle_expected_versions" name="expected_versions" class="textarea" rows="5" placeholder="Kidoz adapter    1.3.0&#10;Kidoz native     9.1.2&#10;... (Library Name <tabs/spaces> Version)"></textarea>
                    </div>
                    <button type="submit" class="button w-full">Analyze Gradle Versions</button>
                </form>
                <div id="gradle-spinner" class="spinner hidden"></div>
            </div>
            <div class="card hidden" id="gradle-scan-results-section">
                <h2 class="text-xl font-semibold text-gray-700 mb-4">Gradle Analysis Results</h2>
                <div id="gradle-version-comparison-output"></div>
                <button id="reset-gradle-verifier" class="button mt-6 w-full">Analyze Another Gradle File</button>
            </div>
        </div>

        <div id="podfile-verifier-content" class="tab-content">
            <div class="card" id="podfile-upload-section">
                <div class="flex items-center justify-between gap-4 flex-wrap mb-4">
                    <h2 class="text-xl font-semibold text-gray-700 mb-0">Podfile Version Verifier</h2>
                    <div class="flex items-center gap-3 flex-wrap">
                        <label for="podfile-build-check-preset" class="text-sm font-semibold text-gray-700">Saved Check List:</label>
                        <select id="podfile-build-check-preset" class="border border-gray-300 rounded-md px-3 py-2 text-sm bg-white min-w-[220px]">
                            <option value="">Manual input</option>
                        </select>
                        <button id="reload-podfile-build-check-presets" type="button" class="button button-secondary text-sm py-2 px-3">Reload</button>
                    </div>
                </div>
                <form id="podfile-upload-form" method="post" action="/analyze_podfile" enctype="multipart/form-data">
                    <div class="mb-4">
                        <label for="podfile_file_scan" class="block text-sm font-medium text-gray-700 mb-1">Select your Podfile:</label>
                        <input type="file" name="podfile_file" id="podfile_file_scan" class="file-input" required>
                    </div>
                    <div class="mb-4">
                        <label for="podfile_expected_versions" class="block text-sm font-medium text-gray-700 mb-1">Input Expected Library Versions (Optional):</label>
                        <textarea id="podfile_expected_versions" name="expected_versions" class="textarea" rows="5" placeholder="Kidoz Adapter    1.3.2&#10;Kidoz SDK     9.1.5&#10;... (Library Name <tabs/spaces> Version)"></textarea>
                    </div>
                    <button type="submit" class="button w-full">Analyze Podfile Versions</button>
                </form>
                <div id="podfile-spinner" class="spinner hidden"></div>
            </div>
            <div class="card hidden" id="podfile-scan-results-section">
                <h2 class="text-xl font-semibold text-gray-700 mb-4">Podfile Analysis Results</h2>
                <div id="podfile-version-comparison-output"></div>
                <button id="reset-podfile-verifier" class="button mt-6 w-full">Analyze Another Podfile</button>
            </div>
        </div>

        <div id="manifest-analyzer-content" class="tab-content">
            <div class="card" id="manifest-upload-section">
                <div class="flex items-center justify-between gap-4 flex-wrap mb-4">
                    <h2 class="text-xl font-semibold text-gray-700 mb-0">APK Manifest Data Analysis</h2>
                    <div class="flex items-center gap-3 flex-wrap">
                        <label for="manifest-check-preset" class="text-sm font-semibold text-gray-700">Saved Check List:</label>
                        <select id="manifest-check-preset" class="border border-gray-300 rounded-md px-3 py-2 text-sm bg-white min-w-[220px]">
                            <option value="">Manual input</option>
                        </select>
                        <button id="reload-manifest-check-presets" type="button" class="button button-secondary text-sm py-2 px-3">Reload</button>
                    </div>
                </div>
                <form id="manifest-upload-form" method="post" action="/analyze_manifest_data" enctype="multipart/form-data">
                    <div class="mb-4">
                        <label for="manifest_apk_file" class="block text-sm font-medium text-gray-700 mb-1">Select your APK file:</label>
                        <input type="file" name="apk_file" id="manifest_apk_file" class="file-input" accept=".apk" required>
                    </div>
                    <button type="submit" class="button w-full">Analyze Manifest Data</button>
                </form>
                <div id="manifest-spinner" class="spinner hidden"></div>
            </div>
            <div class="card hidden" id="manifest-results-section">
                <h2 class="text-xl font-semibold text-gray-700 mb-4">Manifest Analysis Results</h2>
                <div id="manifest-comparison-output" class="mb-6"></div>
                <div id="manifest-attributes-output" class="mb-6"></div>
                <div id="manifest-permissions-output"></div>
                <div id="manifest-metadata-output" class="mt-6"></div>
                <button id="reset-manifest-analyzer" class="button mt-6 w-full">Analyze Another Manifest</button>
            </div>
        </div>

        <div id="aab-converter-content" class="tab-content">
            <div class="card" id="aab-converter-upload-section">
                <h2 class="text-xl font-semibold text-gray-700 mb-4">AAB to APK Converter</h2>
                <form id="aab-converter-upload-form" method="post" action="/convert_aab_to_apk" enctype="multipart/form-data">
                    <div class="mb-4">
                        <label for="aab_convert_file" class="block text-sm font-medium text-gray-700 mb-1">Select your AAB file (.aab):</label>
                        <input type="file" name="aab_file" id="aab_convert_file" class="file-input" accept=".aab" required>
                    </div>
                     <p class="text-xs text-gray-500 mb-3">
                        APK signing will use hardcoded keystore information if enabled in the script.
                        Otherwise, it will attempt to use a debug key if found, or may be unsigned if a debug keystore is not found in default locations.
                    </p>
                    <div id="keystore-input-area">
                        <hr class="my-4">
                        <h3 class="text-lg font-medium text-gray-700 mb-3">Optional: Custom Keystore for Signing</h3>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                            <div>
                                <label for="keystore_path" class="block text-sm font-medium text-gray-700 mb-1">Keystore Path or Filename:</label>
                                <input type="text" name="keystore_path" id="keystore_path" class="text-input" placeholder="my-key.keystore OR /abs/path/to/key.jks">
                                 <p class="text-xs text-gray-500 mt-1">If only filename, ensure it's in the app's root directory.</p>
                            </div>
                            <div>
                                <label for="keystore_alias" class="block text-sm font-medium text-gray-700 mb-1">Key Alias:</label>
                                <input type="text" name="keystore_alias" id="keystore_alias" class="text-input" placeholder="mykeyalias">
                            </div>
                            <div>
                                <label for="keystore_pass" class="block text-sm font-medium text-gray-700 mb-1">Keystore Password:</label>
                                <input type="password" name="keystore_pass" id="keystore_pass" class="text-input">
                            </div>
                            <div>
                                <label for="key_pass" class="block text-sm font-medium text-gray-700 mb-1">Key Password (if different):</label>
                                <input type="password" name="key_pass" id="key_pass" class="text-input">
                            </div>
                        </div>
                    </div>
                     <div id="hardcoded-keystore-info" class="keystore-info-hardcoded hidden">
                        Using hardcoded keystore information: <br>
                        Filename: <span id="hc-ks-filename"></span> <br>
                        Alias: <span id="hc-ks-alias"></span>
                        <p class="text-xs mt-1">Password is not displayed. To use different keystore, modify app.py and set USE_HARDCODED_KEYSTORE to False.</p>
                    </div>
                    <button type="submit" class="button w-full mt-4">Convert AAB to APK</button>
                </form>
                <div id="aab-converter-spinner" class="spinner hidden"></div>
            </div>
            <div class="card hidden" id="aab-conversion-results-section">
                <h2 class="text-xl font-semibold text-gray-700 mb-4">AAB Conversion Result</h2>
                <div id="aab-conversion-output">
                </div>
                <button id="reset-aab-converter" class="button mt-6 w-full">Convert Another AAB</button>
            </div>
        </div>
    </div>

    <script>
        // --- DOM Elements ---
        const apkUploadForm = document.getElementById('apk-upload-form');
        const apkUploadSection = document.getElementById('apk-upload-section');
        const apkScanResultsSection = document.getElementById('apk-scan-results-section');
        const apkScanOutputDiv = document.getElementById('apk-scan-output');
        const apkVersionComparisonOutputDiv = document.getElementById('apk-version-comparison-output');
        const downloadBinaryLink = document.getElementById('download-binary-manifest');
        const noBinaryLink = document.getElementById('no-binary-link');
        const apkSpinner = document.getElementById('apk-spinner');

        const gradleUploadForm = document.getElementById('gradle-upload-form');
        const gradleUploadSection = document.getElementById('gradle-upload-section');
        const gradleScanResultsSection = document.getElementById('gradle-scan-results-section');
        const gradleVersionComparisonOutputDiv = document.getElementById('gradle-version-comparison-output');
        const gradleSpinner = document.getElementById('gradle-spinner');

        const podfileUploadForm = document.getElementById('podfile-upload-form');
        const podfileUploadSection = document.getElementById('podfile-upload-section');
        const podfileScanResultsSection = document.getElementById('podfile-scan-results-section');
        const podfileVersionComparisonOutputDiv = document.getElementById('podfile-version-comparison-output');
        const podfileSpinner = document.getElementById('podfile-spinner');

        const manifestUploadForm = document.getElementById('manifest-upload-form');
        const manifestUploadSection = document.getElementById('manifest-upload-section');
        const manifestResultsSection = document.getElementById('manifest-results-section');
        const manifestComparisonOutputDiv = document.getElementById('manifest-comparison-output');
        const manifestAttributesOutputDiv = document.getElementById('manifest-attributes-output');
        const manifestPermissionsOutputDiv = document.getElementById('manifest-permissions-output');
        const manifestMetadataOutputDiv = document.getElementById('manifest-metadata-output');
        const manifestSpinner = document.getElementById('manifest-spinner');

        const aabConverterUploadForm = document.getElementById('aab-converter-upload-form');
        const aabConverterUploadSection = document.getElementById('aab-converter-upload-section');
        const aabConversionResultsSection = document.getElementById('aab-conversion-results-section');
        const aabConversionOutputDiv = document.getElementById('aab-conversion-output');
        const aabConverterSpinner = document.getElementById('aab-converter-spinner');

        const messageArea = document.getElementById('message-area');
        const buildCheckPresetControls = [
            { selectId: 'apk-build-check-preset', reloadId: 'reload-apk-build-check-presets', presetKey: 'presets', platform: 'android', inputId: 'apk_expected_versions' },
            { selectId: 'gradle-build-check-preset', reloadId: 'reload-gradle-build-check-presets', presetKey: 'gradle_presets', platform: 'android', inputId: 'gradle_expected_versions' },
            { selectId: 'podfile-build-check-preset', reloadId: 'reload-podfile-build-check-presets', presetKey: 'podfile_presets', platform: 'ios', inputId: 'podfile_expected_versions' },
            { selectId: 'manifest-check-preset', reloadId: 'reload-manifest-check-presets', presetKey: 'manifest_presets', platform: 'android', inputId: null },
        ];

        const tabApkAnalyzer = document.getElementById('tab-apk-analyzer');
        const tabGradleVerifier = document.getElementById('tab-gradle-verifier');
        const tabPodfileVerifier = document.getElementById('tab-podfile-verifier');
        const tabManifestAnalyzer = document.getElementById('tab-manifest-analyzer');
        const tabAabConverter = document.getElementById('tab-aab-converter');

        const resetApkAnalyzerButton = document.getElementById('reset-apk-analyzer');
        const resetGradleVerifierButton = document.getElementById('reset-gradle-verifier');
        const resetPodfileVerifierButton = document.getElementById('reset-podfile-verifier');
        const resetManifestAnalyzerButton = document.getElementById('reset-manifest-analyzer');
        const resetAabConverterButton = document.getElementById('reset-aab-converter');

        const keystoreInputArea = document.getElementById('keystore-input-area');
        const hardcodedKeystoreInfoDiv = document.getElementById('hardcoded-keystore-info');
        const hcKsFilenameSpan = document.getElementById('hc-ks-filename');
        const hcKsAliasSpan = document.getElementById('hc-ks-alias');

        // --- Utility Functions ---
        function showSpinner(spinnerElement) { spinnerElement.classList.remove('hidden'); }
        function hideSpinner(spinnerElement) { spinnerElement.classList.add('hidden'); }
        function showMessage(text, type = 'info', area = messageArea) {
            const alertClass = type === 'error' ? 'error-message' : (type === 'warning' ? 'warning-message' : 'info-message');
            const messageDiv = document.createElement('div');
            messageDiv.className = alertClass;
            messageDiv.innerHTML = text; // Use innerHTML to allow basic HTML like <br> or <pre>
            // Keep one current status message. Repeated reloads must not grow a
            // stack of identical banners that pushes the active preset offscreen.
            area.replaceChildren(messageDiv);
        }
        function clearMessages(area = messageArea) { area.innerHTML = ''; }
        function escapeHtml(unsafe) {
            if (unsafe === null || typeof unsafe === 'undefined') return '';
            return String(unsafe)
                 .replace(/&/g, "&amp;")
                 .replace(/</g, "&lt;")
                 .replace(/>/g, "&gt;")
                 .replace(/"/g, "&quot;")
                 .replace(/'/g, "&#039;")
                 .replace(/`/g, "&#96;");
        }
        function normalizeVersion(versionStr) {
            // Extracts the core version string (e.g., "1.2.3" from "1.2.3-alpha")
            if (!versionStr) return '';
            const match = versionStr.match(/\\d+\\.\\d+(\\.\\d+)?/); // Basic semver pattern, without suffix
            return match ? match[0] : versionStr.trim();
        }

        function presetLines(value) {
            if (Array.isArray(value)) return value.join('\\n');
            return typeof value === 'string' ? value : '';
        }

        function presetMatchesControl(control, preset) {
            const platform = String((preset || {}).platform || control.platform).trim().toLowerCase();
            return platform === control.platform;
        }

        function applyBuildCheckPreset(control, preset) {
            if (!control.inputId) return;
            const input = document.getElementById(control.inputId);
            if (input) input.value = presetLines((preset || {}).lines);
        }

        function syncBuildCheckPreset(control, select, clearManual = false) {
            const option = select && select.options[select.selectedIndex];
            if (!option || !option.value) {
                if (clearManual) applyBuildCheckPreset(control, { lines: [] });
                return;
            }
            try {
                const preset = option.dataset.preset ? JSON.parse(option.dataset.preset) : {};
                applyBuildCheckPreset(control, presetMatchesControl(control, preset) ? preset : { lines: [] });
            } catch (error) {
                console.warn('Unable to apply build check preset:', error);
                applyBuildCheckPreset(control, { lines: [] });
            }
        }

        function restoreSelectedBuildCheckPreset(selectId) {
            const control = buildCheckPresetControls.find(item => item.selectId === selectId);
            const select = document.getElementById(selectId);
            if (control && select && select.value) syncBuildCheckPreset(control, select);
        }

        let buildCheckPresetRequestSerial = 0;

        async function loadBuildCheckPresets(forceRemote = false) {
            const requestSerial = ++buildCheckPresetRequestSerial;
            const selectedValues = Object.fromEntries(buildCheckPresetControls.map(control => {
                const select = document.getElementById(control.selectId);
                return [control.selectId, select ? select.value : ''];
            }));
            try {
                const refreshQuery = forceRemote ? '&refresh=1' : '';
                const response = await fetch('/api/build-check-presets?ts=' + Date.now() + refreshQuery, { cache: 'no-store' });
                if (!response.ok) throw new Error('preset_request_failed:' + response.status);
                const data = await response.json();
                // A slower request started before this one must never restore an
                // older Git payload after the latest Reload has completed.
                if (requestSerial !== buildCheckPresetRequestSerial) return;
                if (data.refresh_errors && data.refresh_errors.length) {
                    console.warn('Some Services Checker presets could not be refreshed:', data.refresh_errors);
                    if (forceRemote) {
                        showMessage('Preset reload did not fetch every file from GitHub. The previous local copy was kept.', 'error');
                    }
                } else if (forceRemote) {
                    const branches = [...new Set(Object.values(data.refreshed_sources || {}))].filter(Boolean);
                    const sourceLabel = branches.length ? ` (GitHub ${escapeHtml(branches.join(', '))})` : '';
                    showMessage(`Service Checker presets reloaded from GitHub${sourceLabel}.`, 'info');
                }
                buildCheckPresetControls.forEach(control => {
                    const select = document.getElementById(control.selectId);
                    if (!select) return;
                    const presets = data[control.presetKey] || {};
                    while (select.options.length > 1) select.remove(1);
                    Object.entries(presets)
                        .filter(([, preset]) => presetMatchesControl(control, preset))
                        .forEach(([name, preset]) => {
                            const option = document.createElement('option');
                            option.value = name;
                            option.textContent = name;
                            option.dataset.preset = JSON.stringify(preset || {});
                            select.appendChild(option);
                        });
                    const selectedValue = selectedValues[control.selectId];
                    if (selectedValue && presets[selectedValue] && presetMatchesControl(control, presets[selectedValue])) {
                        select.value = selectedValue;
                    } else if (selectedValue) {
                        select.value = '';
                    }
                    // Re-apply the selected option after rebuilding the list. This
                    // also handles a preset selected by the browser before the
                    // async GitHub response arrives.
                    syncBuildCheckPreset(control, select);
                });
            } catch (error) {
                if (forceRemote) showMessage('Unable to reload Service Checker presets from GitHub.', 'error');
                console.warn('Unable to load build check presets:', error);
            }
        }

        buildCheckPresetControls.forEach(control => {
            const select = document.getElementById(control.selectId);
            const reloadButton = document.getElementById(control.reloadId);
            if (select) select.onchange = () => syncBuildCheckPreset(control, select, true);
            if (reloadButton) reloadButton.onclick = async () => {
                reloadButton.disabled = true;
                reloadButton.textContent = 'Reloading...';
                try {
                    await loadBuildCheckPresets(true);
                } finally {
                    reloadButton.disabled = false;
                    reloadButton.textContent = 'Reload';
                }
            };
        });

        // --- Tab Switching ---
        function switchTab(tabId) {
            clearMessages();
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            document.querySelectorAll('.tab-button').forEach(button => button.classList.remove('active'));

            const contentToShow = document.getElementById(tabId + '-content');
            const buttonToActivate = document.getElementById('tab-' + tabId);

            if (contentToShow) contentToShow.classList.add('active');
            if (buttonToActivate) buttonToActivate.classList.add('active');

            resetApp(tabId); // Reset the specific tab's form and results
        }

        // --- Event Handlers ---
        if (apkUploadForm) {
            apkUploadForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                clearMessages();
                showSpinner(apkSpinner);
                apkScanResultsSection.classList.add('hidden');
                apkScanOutputDiv.innerHTML = '';
                apkVersionComparisonOutputDiv.innerHTML = '';
                downloadBinaryLink.classList.add('hidden');
                noBinaryLink.classList.add('hidden');

                const formData = new FormData(apkUploadForm);
                const expectedVersionsInput = document.getElementById('apk_expected_versions').value;

                try {
                    const response = await fetch('/upload_apk_for_scan', { method: 'POST', body: formData });
                    const data = await response.json();
                    hideSpinner(apkSpinner);

                    if (data.success) {
                        const foundVersions = data.found_file_versions || {};
                        displayArchiveScanResults(foundVersions, apkScanOutputDiv, "APK");

                        if (expectedVersionsInput.trim()) {
                            displayVersionComparison(foundVersions, expectedVersionsInput, apkVersionComparisonOutputDiv, "APK");
                        } else {
                            apkVersionComparisonOutputDiv.innerHTML = ''; // Clear previous
                            const title = document.createElement('h4');
                            title.textContent = 'APK Version Comparison Results';
                            apkVersionComparisonOutputDiv.appendChild(title);
                            const p = document.createElement('p');
                            p.className = 'text-sm text-gray-500';
                            p.textContent = 'No expected versions were entered. Comparison not performed.';
                            apkVersionComparisonOutputDiv.appendChild(p);
                        }

                        apkScanResultsSection.classList.remove('hidden');
                        showMessage('APK file scan complete.', 'info');
                        if(data.warning_message) showMessage(data.warning_message, 'warning');
                        if (data.binary_manifest_url) {
                            downloadBinaryLink.href = data.binary_manifest_url;
                            downloadBinaryLink.download = data.binary_manifest_filename;
                            downloadBinaryLink.classList.remove('hidden');
                            noBinaryLink.classList.add('hidden');
                        } else {
                            downloadBinaryLink.classList.add('hidden');
                            noBinaryLink.classList.remove('hidden');
                        }
                        apkUploadSection.classList.add('hidden');
                    } else {
                        showMessage(data.error || 'Failed to process APK.', 'error');
                        apkUploadSection.classList.remove('hidden'); // Keep form visible on error
                    }
                } catch (error) {
                    hideSpinner(apkSpinner);
                    showMessage('An error occurred during APK processing: ' + error.toString(), 'error');
                    apkUploadSection.classList.remove('hidden');
                }
            });
        }

        if (gradleUploadForm) {
            gradleUploadForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                clearMessages();
                showSpinner(gradleSpinner);
                gradleScanResultsSection.classList.add('hidden');
                gradleVersionComparisonOutputDiv.innerHTML = '';

                const formData = new FormData(gradleUploadForm);

                try {
                    const response = await fetch('/analyze_gradle', { method: 'POST', body: formData });
                    const data = await response.json();
                    hideSpinner(gradleSpinner);

                    if(data.success) {
                        if (data.expected_versions_input.trim()) {
                            displayVersionComparison(data.found_versions, data.expected_versions_input, gradleVersionComparisonOutputDiv, "Gradle");
                        } else {
                            displayArchiveScanResults(data.found_versions, gradleVersionComparisonOutputDiv, "Gradle");
                        }
                        gradleScanResultsSection.classList.remove('hidden');
                        gradleUploadSection.classList.add('hidden');
                        showMessage('Gradle file analysis complete.', 'info');
                    } else {
                        showMessage(data.error || 'Failed to process Gradle file.', 'error');
                        gradleUploadSection.classList.remove('hidden');
                    }
                } catch (error) {
                    hideSpinner(gradleSpinner);
                    showMessage('An error occurred during Gradle processing: ' + error.toString(), 'error');
                    gradleUploadSection.classList.remove('hidden');
                }
            });
        }

        if (podfileUploadForm) {
            podfileUploadForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                clearMessages();
                showSpinner(podfileSpinner);
                podfileScanResultsSection.classList.add('hidden');
                podfileVersionComparisonOutputDiv.innerHTML = '';

                const formData = new FormData(podfileUploadForm);

                try {
                    const response = await fetch('/analyze_podfile', { method: 'POST', body: formData });
                    const data = await response.json();
                    hideSpinner(podfileSpinner);

                    if(data.success) {
                        if (data.expected_versions_input.trim()) {
                            displayVersionComparison(data.found_versions, data.expected_versions_input, podfileVersionComparisonOutputDiv, "Podfile");
                        } else {
                            displayArchiveScanResults(data.found_versions, podfileVersionComparisonOutputDiv, "Podfile");
                        }
                        podfileScanResultsSection.classList.remove('hidden');
                        podfileUploadSection.classList.add('hidden');
                        showMessage('Podfile analysis complete.', 'info');
                    } else {
                        showMessage(data.error || 'Failed to process Podfile.', 'error');
                        podfileUploadSection.classList.remove('hidden');
                    }
                } catch (error) {
                    hideSpinner(podfileSpinner);
                    showMessage('An error occurred during Podfile processing: ' + error.toString(), 'error');
                    podfileUploadSection.classList.remove('hidden');
                }
            });
        }

        if (manifestUploadForm) {
            manifestUploadForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                clearMessages();
                showSpinner(manifestSpinner);
                manifestResultsSection.classList.add('hidden');
                manifestComparisonOutputDiv.innerHTML = '';
                manifestAttributesOutputDiv.innerHTML = '';
                manifestPermissionsOutputDiv.innerHTML = '';
                manifestMetadataOutputDiv.innerHTML = '';

                const formData = new FormData(manifestUploadForm);
                const manifestPresetSelect = document.getElementById('manifest-check-preset');
                if (manifestPresetSelect && manifestPresetSelect.value) {
                    formData.append('preset_name', manifestPresetSelect.value);
                }
                try {
                    const response = await fetch('/analyze_manifest_data', { method: 'POST', body: formData });
                    const data = await response.json();
                    hideSpinner(manifestSpinner);

                    if (data.success) {
                        displayManifestData(
                            data.manifest_attributes,
                            data.permissions,
                            data.metadata,
                            data.apk_filename,
                            data.manifest_comparison,
                        );

                        // Handle messages from backend (e.g., if androguard was missing or manifest not found)
                        if (data.error_message_from_backend) {
                            let msgType = 'warning'; // Default to warning for backend messages that aren't outright failures
                            // Escalate to error if critical components like androguard are missing
                            if (data.error_message_from_backend.toLowerCase().includes("not found") &&
                                (data.error_message_from_backend.toLowerCase().includes("androguard")) ) {
                                msgType = 'error';
                            } else if ((!data.permissions || data.permissions.length === 0) &&
                                       (!data.metadata || Object.keys(data.metadata).length === 0) &&
                                       (!data.manifest_attributes || Object.keys(data.manifest_attributes).length === 0) ) {
                                // If no data AND an error message, it's likely a parsing issue.
                                msgType = 'warning';
                            }
                            showMessage(escapeHtml(data.error_message_from_backend), msgType);
                        } else if ((data.permissions && data.permissions.length > 0) ||
                                   (data.metadata && Object.keys(data.metadata).length > 0) ||
                                   (data.manifest_attributes && Object.keys(data.manifest_attributes).length > 0) ) {
                             showMessage(`Manifest data for ${escapeHtml(data.apk_filename)} displayed.`, 'info');
                        } else {
                             showMessage(`Manifest processed for ${escapeHtml(data.apk_filename)}, but no relevant data was found.`, 'info');
                        }
                        manifestResultsSection.classList.remove('hidden');
                        manifestUploadSection.classList.add('hidden');
                    } else {
                        showMessage(data.error || 'Failed to analyze manifest data.', 'error');
                        manifestUploadSection.classList.remove('hidden');
                    }
                } catch (error) {
                    hideSpinner(manifestSpinner);
                    showMessage('An error occurred: ' + error.toString(), 'error');
                    manifestUploadSection.classList.remove('hidden');
                }
            });
        }

        if (aabConverterUploadForm) {
            aabConverterUploadForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                clearMessages();
                showSpinner(aabConverterSpinner);
                aabConversionResultsSection.classList.add('hidden');
                aabConversionOutputDiv.innerHTML = '';

                const formData = new FormData(aabConverterUploadForm);

                try {
                    const response = await fetch('/convert_aab_to_apk', { method: 'POST', body: formData });
                    const data = await response.json();
                    hideSpinner(aabConverterSpinner);

                    if (data.success) {
                        aabConversionOutputDiv.innerHTML = ''; // Clear previous

                        const successP = document.createElement('p');
                        successP.className = 'info-message';
                        successP.textContent = 'Successfully converted AAB to APK!';
                        aabConversionOutputDiv.appendChild(successP);

                        if(data.warning_message){
                            const warningP = document.createElement('p');
                            warningP.className = 'warning-message';
                            warningP.textContent = data.warning_message;
                            aabConversionOutputDiv.appendChild(warningP);
                        }

                        const apkFilename = String(data.apk_filename || 'generated.apk');
                        if (data.saved_to_downloads !== true) {
                            throw new Error(data.error || 'The converted APK was not saved to Downloads.');
                        }

                        const savedP = document.createElement('p');
                        savedP.className = 'info-message';
                        savedP.textContent = 'APK saved directly to Downloads: ' + String(data.saved_filename || apkFilename);
                        aabConversionOutputDiv.appendChild(savedP);

                        aabConversionResultsSection.classList.remove('hidden');
                        aabConverterUploadSection.classList.add('hidden');
                    } else {
                        showMessage(data.error || 'Failed to convert AAB to APK.', 'error');
                        if(data.bundletool_stderr){
                             showMessage('Bundletool STDERR (see console for full output):<br><pre class="text-xs whitespace-pre-wrap">' + escapeHtml(data.bundletool_stderr.substring(0, 500)) + (data.bundletool_stderr.length > 500 ? '...' : '') + '</pre>', 'error');
                        }
                        aabConverterUploadSection.classList.remove('hidden');
                    }
                } catch (error) {
                    hideSpinner(aabConverterSpinner);
                    showMessage('An error occurred during AAB conversion: ' + error.toString(), 'error');
                    aabConverterUploadSection.classList.remove('hidden');
                }
            });
        }

        // --- Display Functions ---
        function displayArchiveScanResults(versions, outputDiv, archiveType = "Archive") {
            outputDiv.innerHTML = '';
            const resultsSectionDiv = document.createElement('div');
            resultsSectionDiv.className = 'results-section';

            const h4 = document.createElement('h4');
            h4.textContent = 'Found Library Versions in ' + archiveType;
            resultsSectionDiv.appendChild(h4);

            if (versions && Object.keys(versions).length > 0) {
                const p = document.createElement('p');
                p.className = 'text-sm text-gray-500 mb-2';
                p.textContent = 'Versions extracted from the ' + archiveType + ' file.';
                resultsSectionDiv.appendChild(p);

                const ul = document.createElement('ul');
                ul.style.listStyleType = 'none'; // No bullets for this list
                ul.style.paddingLeft = '0';
                for (const [library, version] of Object.entries(versions).sort()) { // Sort by library name
                    const li = document.createElement('li');
                    li.className = 'mb-1';
                    const strong = document.createElement('strong');
                    strong.textContent = library + ': ';
                    li.appendChild(strong);
                    li.appendChild(document.createTextNode(version));
                    ul.appendChild(li);
                }
                resultsSectionDiv.appendChild(ul);
            } else {
                const p = document.createElement('p');
                p.className = 'text-gray-600';
                p.textContent = 'No targeted library versions found in the ' + archiveType + '.';
                resultsSectionDiv.appendChild(p);
            }
            outputDiv.appendChild(resultsSectionDiv);
        }

        function displayVersionComparison(foundVersions, expectedVersionsInput, comparisonDiv, archiveType = "Archive") {
            comparisonDiv.innerHTML = '';
            const title = document.createElement('h4');
            title.textContent = archiveType + ' Version Comparison Results';
            comparisonDiv.appendChild(title);

            if (!expectedVersionsInput || expectedVersionsInput.trim() === "") {
                const p = document.createElement('p');
                p.className = 'text-sm text-gray-500';
                p.textContent = 'No expected versions were provided for ' + archiveType + '. Comparison not performed.';
                comparisonDiv.appendChild(p);
                return;
            }

            const expected = {};
            const lines = expectedVersionsInput.split('\\n');
            let successfullyParsedCount = 0;
            let parsingAttempts = 0;

            lines.forEach((line, index) => {
                const originalLine = line;
                line = line.trim();
                if (!line) return; // Skip empty lines

                parsingAttempts++;
                let libName = '', version = '';

                // Try parsing with multiple spaces or a tab as delimiter first
                const multiSpaceOrTabMatch = line.match(/^(.+?)(?:\\s{2,}|\\t)(.+)$/);
                if (multiSpaceOrTabMatch && multiSpaceOrTabMatch[1] && multiSpaceOrTabMatch[2]) {
                    libName = multiSpaceOrTabMatch[1].trim();
                    version = multiSpaceOrTabMatch[2].trim();
                } else {
                    // Fallback: split by the last space if not matched by tab/multiple spaces
                    const lastSpaceIndex = line.lastIndexOf(' ');
                    if (lastSpaceIndex > 0 && lastSpaceIndex < line.length - 1) {
                        libName = line.substring(0, lastSpaceIndex).trim();
                        version = line.substring(lastSpaceIndex + 1).trim();
                    }
                }

                if (libName && version) {
                    expected[libName] = version;
                    successfullyParsedCount++;
                }
            });

            if (successfullyParsedCount === 0 && parsingAttempts > 0) {
                const warningDiv = document.createElement('div');
                warningDiv.className = 'warning-message p-3 rounded-md'; // Use warning-message style

                const pSemibold = document.createElement('p');
                pSemibold.className = 'font-semibold';
                pSemibold.textContent = 'Input Parsing Issue for ' + archiveType + ' Comparison:';
                warningDiv.appendChild(pSemibold);

                const pTextSm = document.createElement('p');
                pTextSm.className = 'text-sm mt-1';

                pTextSm.appendChild(document.createTextNode("No valid 'Library Name & Version' pairs were parsed from your input. Please ensure each line follows the format:"));
                pTextSm.appendChild(document.createElement('br'));
                const code1 = document.createElement('code');
                code1.textContent = "Library Name<multiple spaces or a tab>Version";
                pTextSm.appendChild(code1);
                pTextSm.appendChild(document.createElement('br'));
                pTextSm.appendChild(document.createTextNode("For example: "));
                const code2 = document.createElement('code');
                code2.className = "block mt-1"; // Make example standout
                code2.innerHTML = "AndroidX Browser&nbsp;&nbsp;&nbsp;&nbsp;1.8.0"; // Using &nbsp; for visual spacing
                pTextSm.appendChild(code2);
                pTextSm.appendChild(document.createElement('br'));
                pTextSm.appendChild(document.createTextNode("Or using a single space as a fallback: "));
                const code3 = document.createElement('code');
                code3.className = "block mt-1";
                code3.textContent = "Firebase Analytics 21.0.0";
                pTextSm.appendChild(code3);

                warningDiv.appendChild(pTextSm);
                comparisonDiv.appendChild(warningDiv);
                return;
            } else if (successfullyParsedCount === 0 && parsingAttempts === 0 && expectedVersionsInput.trim() !== "") {
                // Input was not empty but resulted in zero parsing attempts (e.g., only whitespace lines)
                const p = document.createElement('p');
                p.className = 'text-sm text-gray-500';
                p.textContent = 'The expected versions input for ' + archiveType + ' contained only whitespace. No comparison performed.';
                comparisonDiv.appendChild(p);
                return;
            } else if (successfullyParsedCount === 0 && parsingAttempts === 0 && expectedVersionsInput.trim() === "") {
                // Input was genuinely empty, handled by the first check, but this is a safeguard.
                return;
            }

            const ul = document.createElement('ul');
            ul.style.listStyleType = 'none'; // No bullets
            ul.style.paddingLeft = '0';
            const comparedLibs = new Set();

            // Iterate through expected versions first
            for (const [libName, expectedVersion] of Object.entries(expected)) {
                comparedLibs.add(libName);
                const li = document.createElement('li');
                li.className = 'comparison-item';

                const strong = document.createElement('strong');
                strong.textContent = libName;
                li.appendChild(strong);
                li.appendChild(document.createElement('br'));

                let actualVersionText = 'Not Found in ' + archiveType;
                let statusClass = 'status-failed';
                let statusText = 'FAILED';
                let icon = '❓'; // Default icon

                if (foundVersions.hasOwnProperty(libName)) {
                    const foundVersionRaw = foundVersions[libName];
                    const foundVersionNormalized = normalizeVersion(foundVersionRaw);
                    const expectedVersionNormalized = normalizeVersion(expectedVersion);
                    actualVersionText = foundVersionRaw;
                    if (foundVersionNormalized === expectedVersionNormalized) {
                        statusClass = 'status-passed'; statusText = 'PASSED'; icon = '✅';
                    } else { icon = '❌'; }
                } else if (String(expectedVersion).trim().toLowerCase() === 'removed') {
                    actualVersionText = 'NOT FOUND';
                    statusClass = 'status-passed'; statusText = 'PASSED'; icon = '✅';
                }

                const spanActual = document.createElement('span');
                spanActual.className = 'text-sm';
                spanActual.textContent = 'Actual: ' + actualVersionText;
                li.appendChild(spanActual);
                li.appendChild(document.createElement('br'));

                const spanExpected = document.createElement('span');
                spanExpected.className = 'text-sm';
                spanExpected.textContent = 'Expected: ' + expectedVersion;
                li.appendChild(spanExpected);
                li.appendChild(document.createElement('br'));

                const spanStatusOuter = document.createElement('span');
                spanStatusOuter.className = 'text-sm';
                spanStatusOuter.textContent = 'Status: ';

                const spanStatusInner = document.createElement('span');
                spanStatusInner.className = statusClass;
                spanStatusInner.textContent = icon + ' ' + statusText;
                spanStatusOuter.appendChild(spanStatusInner);
                li.appendChild(spanStatusOuter);
                li.dataset.resultStatus = statusText;
                ul.appendChild(li);
            }

            // The Gradle preset is the visible allow-list for this comparison.
            // Keep scanning every mapped artifact, but do not render mapped
            // dependencies that are intentionally absent from the selected
            // Gradle preset. APK and Podfile comparisons retain the existing
            // informational rows for unlisted findings.
            if (archiveType !== "Gradle") {
              // Add any found versions that were not in the expected list
              for (const [libName, foundVersionRaw] of Object.entries(foundVersions)) {
                 if (!comparedLibs.has(libName)) { // Only if not already processed
                    const li = document.createElement('li');
                    li.className = 'comparison-item';

                    const strong = document.createElement('strong');
                    strong.textContent = libName;
                    li.appendChild(strong);
                    li.appendChild(document.createElement('br'));

                    const spanActual = document.createElement('span');
                    spanActual.className = 'text-sm';
                    spanActual.textContent = 'Actual: ' + foundVersionRaw;
                    li.appendChild(spanActual);
                    li.appendChild(document.createElement('br'));

                    const spanExpected = document.createElement('span');
                    spanExpected.className = 'text-sm';
                    spanExpected.textContent = 'Expected: N/A (Not in your list)';
                    li.appendChild(spanExpected);
                    li.appendChild(document.createElement('br'));

                    const spanStatusOuter = document.createElement('span');
                    spanStatusOuter.className = 'text-sm';
                    spanStatusOuter.textContent = 'Status: ';

                    const spanStatusInner = document.createElement('span');
                    spanStatusInner.className = 'status-info';
                    spanStatusInner.textContent = 'ℹ️ INFO (Found in ' + archiveType + ', not in expected list)';
                    spanStatusOuter.appendChild(spanStatusInner);
                    li.appendChild(spanStatusOuter);
                    li.dataset.resultStatus = 'INFO';
                    ul.appendChild(li);
                }
              }
            }
            const resultRank = { FAILED: 0, INFO: 1, PASSED: 2 };
            Array.from(ul.children)
                .sort((left, right) => (resultRank[left.dataset.resultStatus] ?? 9) - (resultRank[right.dataset.resultStatus] ?? 9))
                .forEach(item => ul.appendChild(item));
            comparisonDiv.appendChild(ul);
        }

        function appendManifestComparisonRow(container, row, extraClass = '') {
            const item = document.createElement('div');
            item.className = `comparison-item ${extraClass}`.trim();
            const statusClass = row.status === 'PASSED'
                ? 'status-passed'
                : (row.status === 'STRANGE' ? 'status-strange' : 'status-failed');
            item.innerHTML = `
                <strong>${escapeHtml(row.name)}</strong><br>
                Actual: <span class="code">${escapeHtml(row.actual)}</span><br>
                Expected: <span class="code">${escapeHtml(row.expected)}</span><br>
                Status: <span class="${statusClass}">${escapeHtml(row.status)}</span>`;
            container.appendChild(item);
        }

        function displayManifestComparison(comparison) {
            if (!manifestComparisonOutputDiv || !comparison) return;
            const section = document.createElement('div');
            section.className = 'results-section';
            const heading = document.createElement('h4');
            heading.textContent = 'Preset Comparison Results';
            section.appendChild(heading);

            const groups = [
                ['Core Manifest', comparison.core || []],
                ['Permissions', comparison.permissions || []],
                ['Appmetrica Unity', comparison.metadata || []],
                ['Strange Permissions', comparison.strange_permissions || []],
            ];
            groups.forEach(([title, rows]) => {
                if (!rows.length) return;
                const groupTitle = document.createElement('h5');
                groupTitle.className = 'mt-4 mb-2 font-semibold text-gray-700';
                groupTitle.textContent = title;
                section.appendChild(groupTitle);
                const rowClass = title === 'Permissions' ? 'manifest-permission-item' : '';
                rows.forEach(row => appendManifestComparisonRow(section, row, rowClass));
            });
            manifestComparisonOutputDiv.appendChild(section);
        }

        function displayManifestData(manifestAttributes, permissions, metadata, apkFilename, comparison) {
            manifestComparisonOutputDiv.innerHTML = '';
            manifestAttributesOutputDiv.innerHTML = '';
            manifestPermissionsOutputDiv.innerHTML = '';
            manifestMetadataOutputDiv.innerHTML = '';

            displayManifestComparison(comparison);

            // Display Manifest Attributes
            const attrSectionDiv = document.createElement('div');
            attrSectionDiv.className = 'results-section data-list';
            const attrH4 = document.createElement('h4');
            attrH4.textContent = `Core Manifest Attributes for ${escapeHtml(apkFilename)}`;
            attrSectionDiv.appendChild(attrH4);

            if (manifestAttributes && Object.keys(manifestAttributes).length > 0) {
                const attrUl = document.createElement('ul');
                for (const [key, value] of Object.entries(manifestAttributes)) {
                    const li = document.createElement('li');
                    li.className = 'mb-1';
                    const strong = document.createElement('strong');
                    strong.textContent = escapeHtml(key) + ': ';
                    li.appendChild(strong);
                    li.appendChild(document.createTextNode(escapeHtml(value !== null && typeof value !== 'undefined' ? value : 'Not Set')));
                    attrUl.appendChild(li);
                }
                attrSectionDiv.appendChild(attrUl);
            } else {
                const attrP = document.createElement('p');
                attrP.className = 'text-gray-600';
                attrP.textContent = 'No core manifest attributes found or extracted.';
                attrSectionDiv.appendChild(attrP);
            }
            manifestAttributesOutputDiv.appendChild(attrSectionDiv);


            // Display Permissions
            const permsSectionDiv = document.createElement('div');
            permsSectionDiv.className = 'results-section';
            const permsH4 = document.createElement('h4');
            // Display total permission count
            permsH4.textContent = `Permissions (Total: ${permissions ? permissions.length : 0})`;
            permsSectionDiv.appendChild(permsH4);

            if (permissions && permissions.length > 0) {
                const permsUl = document.createElement('ul');
                permissions.forEach(permission => {
                    const li = document.createElement('li');
                    li.className = 'text-sm font-mono'; // Good for permission strings
                    li.textContent = escapeHtml(permission);
                    permsUl.appendChild(li);
                });
                permsSectionDiv.appendChild(permsUl);
            } else {
                const permsP = document.createElement('p');
                permsP.className = 'text-gray-600';
                permsP.textContent = 'No permissions found in the manifest.';
                permsSectionDiv.appendChild(permsP);
            }
            manifestPermissionsOutputDiv.appendChild(permsSectionDiv);

            // Display Metadata
            const metaSectionDiv = document.createElement('div');
            metaSectionDiv.className = 'results-section data-list';
            const metaH4 = document.createElement('h4');
            metaH4.textContent = `Targeted Metadata for ${escapeHtml(apkFilename)}`;
            metaSectionDiv.appendChild(metaH4);

            if (metadata && Object.keys(metadata).length > 0) {
                const metaUl = document.createElement('ul');
                for (const [key, value] of Object.entries(metadata).sort()) { // Sort by key for consistent order
                    const li = document.createElement('li');
                    li.className = 'mb-1';
                    const strong = document.createElement('strong');
                    strong.textContent = escapeHtml(key) + ': '; // Key is already the display name
                    li.appendChild(strong);
                    li.appendChild(document.createTextNode(escapeHtml(value)));
                    metaUl.appendChild(li);
                }
                metaSectionDiv.appendChild(metaUl);
            } else {
                const metaP = document.createElement('p');
                metaP.className = 'text-gray-600';
                metaP.textContent = 'No targeted metadata found in the manifest.';
                metaSectionDiv.appendChild(metaP);
            }
            manifestMetadataOutputDiv.appendChild(metaSectionDiv);
        }

        // --- Reset Function ---
        function resetApp(mode) {
            // General cleanup call to backend, if needed (e.g., for server-side session files not tied to a specific download)
            fetch('/cleanup', { method: 'POST' }); // Call general cleanup

            if (mode === 'apk-analyzer') {
                if(apkUploadForm) apkUploadForm.reset();
                if(document.getElementById('apk_expected_versions')) document.getElementById('apk_expected_versions').value = '';
                restoreSelectedBuildCheckPreset('apk-build-check-preset');
                if(apkUploadSection) apkUploadSection.classList.remove('hidden');
                if(apkScanResultsSection) apkScanResultsSection.classList.add('hidden');
                if(apkScanOutputDiv) apkScanOutputDiv.innerHTML = '';
                if(apkVersionComparisonOutputDiv) apkVersionComparisonOutputDiv.innerHTML = '';
                if(downloadBinaryLink) downloadBinaryLink.classList.add('hidden');
                if(noBinaryLink) noBinaryLink.classList.add('hidden');
            } else if (mode === 'gradle-verifier') {
                if(gradleUploadForm) gradleUploadForm.reset();
                restoreSelectedBuildCheckPreset('gradle-build-check-preset');
                if(gradleUploadSection) gradleUploadSection.classList.remove('hidden');
                if(gradleScanResultsSection) gradleScanResultsSection.classList.add('hidden');
                if(gradleVersionComparisonOutputDiv) gradleVersionComparisonOutputDiv.innerHTML = '';
            } else if (mode === 'podfile-verifier') {
                if(podfileUploadForm) podfileUploadForm.reset();
                restoreSelectedBuildCheckPreset('podfile-build-check-preset');
                if(podfileUploadSection) podfileUploadSection.classList.remove('hidden');
                if(podfileScanResultsSection) podfileScanResultsSection.classList.add('hidden');
                if(podfileVersionComparisonOutputDiv) podfileVersionComparisonOutputDiv.innerHTML = '';
            } else if (mode === 'manifest-analyzer') {
                if(manifestUploadForm) manifestUploadForm.reset();
                if(manifestUploadSection) manifestUploadSection.classList.remove('hidden');
                if(manifestResultsSection) manifestResultsSection.classList.add('hidden');
                if(manifestComparisonOutputDiv) manifestComparisonOutputDiv.innerHTML = '';
                if(manifestAttributesOutputDiv) manifestAttributesOutputDiv.innerHTML = '';
                if(manifestPermissionsOutputDiv) manifestPermissionsOutputDiv.innerHTML = '';
                if(manifestMetadataOutputDiv) manifestMetadataOutputDiv.innerHTML = '';
            } else if (mode === 'aab-converter') {
                if(aabConverterUploadForm) aabConverterUploadForm.reset();
                if(aabConverterUploadSection) aabConverterUploadSection.classList.remove('hidden');
                if(aabConversionResultsSection) aabConversionResultsSection.classList.add('hidden');
                if(aabConversionOutputDiv) aabConversionOutputDiv.innerHTML = '';
            }
        }

        // --- Initialize Event Listeners for Tabs and Resets ---
        document.addEventListener('DOMContentLoaded', () => {
            if (tabApkAnalyzer) tabApkAnalyzer.addEventListener('click', () => switchTab('apk-analyzer'));
            if (tabGradleVerifier) tabGradleVerifier.addEventListener('click', () => switchTab('gradle-verifier'));
            if (tabPodfileVerifier) tabPodfileVerifier.addEventListener('click', () => switchTab('podfile-verifier'));
            if (tabManifestAnalyzer) tabManifestAnalyzer.addEventListener('click', () => switchTab('manifest-analyzer'));
            if (tabAabConverter) tabAabConverter.addEventListener('click', () => switchTab('aab-converter'));

            if (resetApkAnalyzerButton) resetApkAnalyzerButton.addEventListener('click', () => { clearMessages(); resetApp('apk-analyzer'); });
            if (resetGradleVerifierButton) resetGradleVerifierButton.addEventListener('click', () => { clearMessages(); resetApp('gradle-verifier'); });
            if (resetPodfileVerifierButton) resetPodfileVerifierButton.addEventListener('click', () => { clearMessages(); resetApp('podfile-verifier'); });
            if (resetManifestAnalyzerButton) resetManifestAnalyzerButton.addEventListener('click', () => { clearMessages(); resetApp('manifest-analyzer'); });
            if (resetAabConverterButton) resetAabConverterButton.addEventListener('click', () => { clearMessages(); resetApp('aab-converter'); });

            // Initialize Keystore UI
            fetch('/get_keystore_config')
                .then(response => response.json())
                .then(data => {
                    if (data.use_hardcoded) {
                        if(keystoreInputArea) keystoreInputArea.classList.add('hidden');
                        if(hardcodedKeystoreInfoDiv) hardcodedKeystoreInfoDiv.classList.remove('hidden');
                        if(hcKsFilenameSpan) hcKsFilenameSpan.textContent = data.filename;
                        if(hcKsAliasSpan) hcKsAliasSpan.textContent = data.alias;
                    } else {
                        if(keystoreInputArea) keystoreInputArea.classList.remove('hidden');
                        if(hardcodedKeystoreInfoDiv) hardcodedKeystoreInfoDiv.classList.add('hidden');
                    }
                });

            // Refresh on every Service Checker start. Previously only the
            // per-control Reload buttons used refresh=1, so a restarted app
            // could keep showing an old per-user preset cache indefinitely.
            loadBuildCheckPresets(true);
            switchTab('apk-analyzer'); // Default to the first tab
        });

    </script>
</body>
</html>
"""

# --- Flask Routes ---
APK_CHECK_PRESETS_FILENAME = "apk_check_presets.json"
GRADLE_CHECK_PRESETS_FILENAME = "gradle_check_presets.json"
PODFILE_CHECK_PRESETS_FILENAME = "podfile_check_presets.json"
MANIFEST_CHECK_PRESETS_FILENAME = "manifest_check_presets.json"
GRADLE_LIB_MAPPING_FILENAME = "gradle_lib_mapping.json"
PODFILE_LIB_MAPPING_FILENAME = "podfile_lib_mapping.json"

# Presets are live data. The Service Checker always refreshes these files when
# its page starts, and the Reload buttons repeat the same operation. `main` is
# the single editable source users see on GitHub. Do not let an inherited
# release-build environment redirect this data to an older branch: that made
# Reload report success while returning stale presets.
SERVICES_CHECKER_PRESET_BRANCH = "main"
SERVICES_CHECKER_PRESET_BRANCHES = (SERVICES_CHECKER_PRESET_BRANCH,)
SERVICES_CHECKER_GITHUB_REPO = "trucbm/Eventchecker"
SERVICES_CHECKER_GITHUB_API_URL = (
    f"https://api.github.com/repos/{SERVICES_CHECKER_GITHUB_REPO}/commits/"
    f"{SERVICES_CHECKER_PRESET_BRANCH}"
)
SERVICES_CHECKER_COMMIT_SHA_PATTERN = re.compile(r"^[0-9a-f]{40}$", re.IGNORECASE)
SERVICES_CHECKER_PRESET_FILENAMES = (
    APK_CHECK_PRESETS_FILENAME,
    GRADLE_CHECK_PRESETS_FILENAME,
    PODFILE_CHECK_PRESETS_FILENAME,
    MANIFEST_CHECK_PRESETS_FILENAME,
)
SERVICES_CHECKER_REMOTE_DATA_FILENAMES = (
    *SERVICES_CHECKER_PRESET_FILENAMES,
    GRADLE_LIB_MAPPING_FILENAME,
    PODFILE_LIB_MAPPING_FILENAME,
)

# Keep the payload returned by the most recent successful Git refresh in
# memory. The disk cache remains the cross-session fallback, but the same
# request must use the bytes it just downloaded instead of re-reading an
# older copy from a different bundled/portable path.
_live_remote_preset_payloads = {}
_remote_preset_refresh_lock = threading.Lock()


def _services_checker_preset_cache_dir():
    """Return a writable per-user cache shared by installed and portable builds."""
    configured = os.getenv("EVENTINSPECTOR_PRESET_CACHE_DIR", "").strip()
    if configured:
        cache_dir = os.path.expandvars(os.path.expanduser(configured))
    elif os.name == "nt":
        cache_dir = os.path.join(
            os.getenv("LOCALAPPDATA") or os.path.expanduser("~"),
            "EventInspector",
            "services_checker_presets",
        )
    elif sys.platform == "darwin":
        cache_dir = os.path.expanduser(
            "~/Library/Application Support/EventInspector/services_checker_presets"
        )
    else:
        cache_dir = os.path.expanduser("~/.eventinspector/services_checker_presets")
    try:
        os.makedirs(cache_dir, exist_ok=True)
    except OSError:
        # A bundled app may be launched from a read-only location. Keep the
        # live preset cache writable without touching the installed payload.
        cache_dir = os.path.join(tempfile.gettempdir(), "EventInspector", "services_checker_presets")
        os.makedirs(cache_dir, exist_ok=True)
    return cache_dir


def _preset_file_candidates(filename):
    """Prefer the latest writable Git cache, then an offline source fallback."""
    return (
        os.path.join(_services_checker_preset_cache_dir(), filename),
        os.path.join(APP_ROOT, filename),
    )


def _fetch_services_checker_revision():
    """Resolve the current main commit before downloading live preset files.

    GitHub Raw caches branch paths for several minutes. A commit SHA is
    immutable, so using it for the file paths makes Reload reflect the latest
    main commit immediately while retaining the branch URL fallback below.
    """
    cache_bust = time.time_ns()
    try:
        response = requests.get(
            f"{SERVICES_CHECKER_GITHUB_API_URL}?eventinspector_refresh={cache_bust}",
            headers={
                "Accept": "application/vnd.github+json",
                "Cache-Control": "no-cache, no-store, max-age=0",
                "Pragma": "no-cache",
                "User-Agent": "EventInspector-Service-Checker",
            },
            timeout=8,
        )
        response.raise_for_status()
        revision = str((response.json() or {}).get("sha") or "").strip()
        if not SERVICES_CHECKER_COMMIT_SHA_PATTERN.fullmatch(revision):
            raise ValueError("github_main_commit_sha_invalid")
        return revision.lower()
    except Exception as exc:
        logger.warning(
            "Could not resolve the current GitHub main commit; using branch fallback: %s",
            exc,
        )
        return None


def _remote_preset_urls(filename, revision=None):
    for branch in SERVICES_CHECKER_PRESET_BRANCHES:
        if revision:
            # These URLs are immutable and therefore do not inherit the
            # multi-minute cache attached to the moving main branch.
            yield branch, (
                f"https://raw.githubusercontent.com/{SERVICES_CHECKER_GITHUB_REPO}/{revision}/"
                f"services_checker/{filename}"
            )
            yield branch, (
                f"https://github.com/{SERVICES_CHECKER_GITHUB_REPO}/raw/{revision}/"
                f"services_checker/{filename}"
            )
            yield branch, (
                f"https://cdn.jsdelivr.net/gh/{SERVICES_CHECKER_GITHUB_REPO}@{revision}/"
                f"services_checker/{filename}"
            )
        yield branch, (
            f"https://raw.githubusercontent.com/{SERVICES_CHECKER_GITHUB_REPO}/{branch}/"
            f"services_checker/{filename}"
        )
        yield branch, (
            f"https://github.com/{SERVICES_CHECKER_GITHUB_REPO}/raw/{branch}/"
            f"services_checker/{filename}"
        )
        yield branch, (
            f"https://cdn.jsdelivr.net/gh/{SERVICES_CHECKER_GITHUB_REPO}@{branch}/"
            f"services_checker/{filename}"
        )


def _fetch_remote_preset(filename, revision=None):
    """Fetch and validate one preset file, bypassing intermediary caches."""
    last_error = None
    cache_bust = time.time_ns()
    for branch, url in _remote_preset_urls(filename, revision):
        separator = "&" if "?" in url else "?"
        try:
            response = requests.get(
                f"{url}{separator}eventinspector_refresh={cache_bust}",
                headers={
                    "Accept": "application/json",
                    "Cache-Control": "no-cache, no-store, max-age=0",
                    "Pragma": "no-cache",
                    "Accept-Encoding": "identity",
                },
                timeout=8,
            )
            response.raise_for_status()
            payload = response.content
            decoded = response.json()
            if not isinstance(decoded, dict) or not decoded:
                raise ValueError("preset_root_must_be_object")
            if filename == GRADLE_LIB_MAPPING_FILENAME and not _normalize_gradle_lib_mapping(decoded):
                raise ValueError("gradle_mapping_must_contain_group_artifact_entries")
            if filename == PODFILE_LIB_MAPPING_FILENAME and not _normalize_podfile_lib_mapping(decoded):
                raise ValueError("podfile_mapping_must_contain_pod_entries")
            return payload, branch
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"remote_preset_unavailable:{filename}:{last_error}")


def _refresh_remote_preset_files():
    """Serialize Git refreshes so an older request cannot overwrite a newer one."""
    with _remote_preset_refresh_lock:
        return _refresh_remote_preset_files_unlocked()


def _refresh_remote_preset_files_unlocked():
    """Refresh every preset atomically; keep the previous copy on failures."""
    cache_dir = _services_checker_preset_cache_dir()
    refreshed = []
    errors = []
    refreshed_sources = {}
    refreshed_digests = {}
    revision = _fetch_services_checker_revision()

    def fetch_one(filename):
        try:
            return filename, _fetch_remote_preset(filename, revision), None
        except Exception as exc:
            return filename, None, exc

    # Each file has the same GitHub fallback chain. Fetching them concurrently
    # keeps one slow/unreachable mirror from multiplying the total reload time.
    with ThreadPoolExecutor(max_workers=len(SERVICES_CHECKER_REMOTE_DATA_FILENAMES)) as executor:
        results = list(executor.map(fetch_one, SERVICES_CHECKER_REMOTE_DATA_FILENAMES))

    for filename, result, error in results:
        if error is not None:
            errors.append(str(error))
            continue

        payload, branch = result
        refreshed_digests[filename] = hashlib.sha256(payload).hexdigest()
        temp_path = os.path.join(cache_dir, f".{filename}.{secrets.token_hex(4)}.tmp")
        try:
            with open(temp_path, "wb") as handle:
                handle.write(payload)
            os.replace(temp_path, os.path.join(cache_dir, filename))
            _live_remote_preset_payloads[filename] = payload
            refreshed.append(filename)
            refreshed_sources[filename] = branch
        except Exception as exc:
            errors.append(f"{filename}:{exc}")
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except OSError:
                pass
    return refreshed, errors, refreshed_sources, refreshed_digests


def _preset_file_metadata():
    """Describe the exact preset copy currently visible to the client."""
    cache_dir = _services_checker_preset_cache_dir()
    metadata = {}
    for filename in SERVICES_CHECKER_REMOTE_DATA_FILENAMES:
        live_payload = _live_remote_preset_payloads.get(filename)
        if live_payload is not None:
            metadata[filename] = {
                "source": "github-live",
                "sha256": hashlib.sha256(live_payload).hexdigest(),
                "size": len(live_payload),
            }
            continue
        cache_path = os.path.join(cache_dir, filename)
        bundled_path = os.path.join(APP_ROOT, filename)
        for path, source in ((cache_path, "github-cache"), (bundled_path, "offline-source")):
            try:
                with open(path, "rb") as handle:
                    payload = handle.read()
            except OSError:
                continue
            metadata[filename] = {
                "source": source,
                "sha256": hashlib.sha256(payload).hexdigest(),
                "size": len(payload),
            }
            break
    return metadata


def _preset_revision(metadata):
    digest_input = {
        filename: (metadata.get(filename) or {}).get("sha256", "")
        for filename in SERVICES_CHECKER_REMOTE_DATA_FILENAMES
    }
    return hashlib.sha256(
        json.dumps(digest_input, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()


def _load_preset_file(filename, default_platform):
    payload_candidates = []
    live_payload = _live_remote_preset_payloads.get(filename)
    if live_payload is not None:
        try:
            payload_candidates.append(("<live-remote>", json.loads(live_payload.decode("utf-8"))))
        except (UnicodeDecodeError, json.JSONDecodeError):
            # Fall back to the atomic disk cache if a malformed payload ever
            # reaches this point; _fetch_remote_preset already validates JSON.
            pass

    for path in _preset_file_candidates(filename):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                payload_candidates.append((path, json.load(handle)))
        except Exception as exc:
            logger.warning("Preset file unavailable (%s): %s", path, exc)
    for source, payload in payload_candidates:
        if not isinstance(payload, dict):
            continue
        presets = {}
        for name, preset in payload.items():
            if isinstance(preset, list):
                preset = {"lines": preset}
            if not isinstance(preset, dict):
                continue
            normalized = dict(preset)
            normalized.setdefault("platform", default_platform)
            normalized.setdefault("lines", [])
            presets[str(name)] = normalized
        if presets:
            return presets
    return {}


def _load_build_check_presets():
    return _load_preset_file(APK_CHECK_PRESETS_FILENAME, "android")


def _load_gradle_check_presets():
    return _load_preset_file(GRADLE_CHECK_PRESETS_FILENAME, "android")


def _load_podfile_check_presets():
    return _load_preset_file(PODFILE_CHECK_PRESETS_FILENAME, "ios")


def _load_manifest_check_presets():
    return _load_preset_file(MANIFEST_CHECK_PRESETS_FILENAME, "android")


def _normalize_gradle_lib_mapping(payload):
    """Return a safe display-name -> ``group:artifact`` mapping from JSON."""
    if not isinstance(payload, dict):
        return {}

    mapping = {}
    for raw_name, raw_artifact in payload.items():
        name = str(raw_name or "").strip()
        artifact = str(raw_artifact or "").strip()
        if not name or not artifact or ":" not in artifact:
            continue
        mapping[name] = artifact
    return mapping


def _load_gradle_lib_mapping():
    """Load the latest remote mapping, with cache and offline fallback."""
    payload_candidates = []
    live_payload = _live_remote_preset_payloads.get(GRADLE_LIB_MAPPING_FILENAME)
    if live_payload is not None:
        try:
            payload_candidates.append(("<live-remote>", json.loads(live_payload.decode("utf-8"))))
        except (UnicodeDecodeError, json.JSONDecodeError):
            pass

    for path in _preset_file_candidates(GRADLE_LIB_MAPPING_FILENAME):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                payload_candidates.append((path, json.load(handle)))
        except Exception as exc:
            logger.warning("Gradle mapping file unavailable (%s): %s", path, exc)

    for _source, payload in payload_candidates:
        mapping = _normalize_gradle_lib_mapping(payload)
        if mapping:
            return mapping
    return {}


def _normalize_podfile_lib_mapping(payload):
    """Return a safe display-name -> CocoaPods name mapping from JSON."""
    if not isinstance(payload, dict):
        return {}

    mapping = {}
    for raw_name, raw_pod in payload.items():
        name = str(raw_name or "").strip()
        pod_name = str(raw_pod or "").strip()
        if not name or not pod_name:
            continue
        mapping[name] = pod_name
    return mapping


def _load_podfile_lib_mapping():
    """Load the latest remote Podfile mapping, with cache and offline fallback."""
    payload_candidates = []
    live_payload = _live_remote_preset_payloads.get(PODFILE_LIB_MAPPING_FILENAME)
    if live_payload is not None:
        try:
            payload_candidates.append(("<live-remote>", json.loads(live_payload.decode("utf-8"))))
        except (UnicodeDecodeError, json.JSONDecodeError):
            pass

    for path in _preset_file_candidates(PODFILE_LIB_MAPPING_FILENAME):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                payload_candidates.append((path, json.load(handle)))
        except Exception as exc:
            logger.warning("Podfile mapping file unavailable (%s): %s", path, exc)

    for _source, payload in payload_candidates:
        mapping = _normalize_podfile_lib_mapping(payload)
        if mapping:
            return mapping
    return {}


def _resolve_manifest_permission(permission_name, actual_package):
    """Resolve package-scoped permission templates against the package found in the APK."""
    template = str(permission_name or "").strip()
    package_name = actual_package if actual_package and actual_package != "Not Found" else ""
    if "{{PACKAGE_NAME}}" in template:
        return template.replace("{{PACKAGE_NAME}}", package_name or "")
    return template


def _manifest_compare_status(actual, expected):
    actual_value = str(actual or "").strip()
    expected_value = str(expected or "").strip()
    if not actual_value or actual_value.lower().startswith("not found"):
        return "FAILED"
    return "PASSED" if actual_value == expected_value else "FAILED"


def compare_manifest_data(manifest_attributes, permissions, metadata, preset):
    """Compare extracted APK manifest data with one saved Android manifest preset."""
    if not isinstance(preset, dict):
        return None

    manifest_attributes = manifest_attributes or {}
    metadata = metadata or {}
    actual_package = str(manifest_attributes.get("Package Name") or "Not Found").strip()

    core_rows = []
    for name, expected in (preset.get("manifest") or {}).items():
        actual = manifest_attributes.get(name, "Not Found")
        core_rows.append({
            "name": str(name),
            "actual": str(actual),
            "expected": str(expected),
            "status": _manifest_compare_status(actual, expected),
        })

    actual_permissions = {
        str(permission).strip()
        for permission in (permissions or [])
        if str(permission).strip()
    }
    expected_permissions = []
    expected_permission_names = set()
    for permission_template in (preset.get("permissions") or []):
        resolved_permission = _resolve_manifest_permission(
            permission_template,
            actual_package,
        )
        expected_permission_names.add(resolved_permission)
        found = resolved_permission in actual_permissions
        expected_permissions.append({
            "name": resolved_permission,
            "actual": resolved_permission if found else "NOT FOUND",
            "expected": resolved_permission,
            "status": "PASSED" if found else "FAILED",
        })

    strange_permissions = [
        {
            "name": permission,
            "actual": permission,
            "expected": "Not in preset",
            "status": "STRANGE",
        }
        for permission in sorted(actual_permissions - expected_permission_names)
    ]

    expected_appmetrica = str(preset.get("appmetrica_unity_version") or "").strip()
    metadata_actual = metadata.get("Appmetrica Unity version", "Not Found")
    metadata_rows = []
    if expected_appmetrica:
        metadata_rows.append({
            "name": "Appmetrica Unity version",
            "actual": str(metadata_actual),
            "expected": expected_appmetrica,
            "status": _manifest_compare_status(metadata_actual, expected_appmetrica),
        })

    return {
        "actual_package_name": actual_package,
        "core": core_rows,
        "permissions": expected_permissions,
        "strange_permissions": strange_permissions,
        "metadata": metadata_rows,
    }


@app.route('/')
def index():
    """Renders the main HTML page and clears session."""
    cleanup_session_files() # Clean up any old files when the main page is loaded
    return render_template_string(HTML_TEMPLATE)


@app.get('/api/build-check-presets')
def get_build_check_presets_final():
    explicit_refresh = request.args.get("refresh", "").strip().lower() in {"1", "true", "yes"}
    # Always refresh on the server before loading the response. The frontend
    # still sends refresh=1 for its Reload buttons, but making this endpoint
    # authoritative also fixes stale presets after an app restart or when an
    # older embedded HTML page omits the query parameter.
    refresh_requested = True
    refreshed = []
    refresh_errors = []
    refreshed_sources = {}
    refreshed_digests = {}
    if refresh_requested:
        refreshed, refresh_errors, refreshed_sources, refreshed_digests = _refresh_remote_preset_files()
    presets = _load_build_check_presets()
    gradle_presets = _load_gradle_check_presets()
    podfile_presets = _load_podfile_check_presets()
    manifest_presets = _load_manifest_check_presets()
    loaded_preset_files = _preset_file_metadata()
    if len(refreshed) == len(SERVICES_CHECKER_REMOTE_DATA_FILENAMES):
        source = 'github'
    elif refreshed:
        source = 'github (partial)'
    else:
        source = 'local preset files'
    return jsonify({
        'success': bool(presets or gradle_presets or podfile_presets or manifest_presets),
        'source': source,
        'refreshed_files': refreshed,
        'refreshed_sources': refreshed_sources,
        'refreshed_digests': refreshed_digests,
        'loaded_preset_files': loaded_preset_files,
        'preset_revision': _preset_revision(loaded_preset_files),
        'refresh_errors': refresh_errors,
        'refresh_requested': explicit_refresh,
        'presets': presets,
        'gradle_presets': gradle_presets,
        'podfile_presets': podfile_presets,
        'manifest_presets': manifest_presets,
    })

@app.route('/get_keystore_config') # New endpoint
def get_keystore_config():
    """Provides frontend with keystore configuration status."""
    return jsonify({
        'use_hardcoded': USE_HARDCODED_KEYSTORE,
        'filename': DEFAULT_KEYSTORE_FILENAME if USE_HARDCODED_KEYSTORE else '',
        'alias': DEFAULT_KEYSTORE_ALIAS if USE_HARDCODED_KEYSTORE else ''
    })

@app.route('/upload_apk_for_scan', methods=['POST'])
def upload_apk_for_scan():
    """Handles APK file upload for library analysis."""
    if 'apk_file' not in request.files:
        return jsonify({'success': False, 'error': 'No APK file part'})
    file = request.files['apk_file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No selected APK file'})

    if file and file.filename.endswith('.apk'):
        original_apk_name = file.filename
        # Use a unique name for the temporary file to avoid conflicts
        temp_apk_scan_filename = f"temp_apk_scan_{secrets.token_hex(8)}_{original_apk_name}"
        apk_savelocation = os.path.join(app.config['UPLOAD_FOLDER'], temp_apk_scan_filename)
        session['temp_apk_for_scan_path'] = apk_savelocation # Track for cleanup

        try:
            file.save(apk_savelocation)
        except Exception as e:
            logger.error(f"Error saving uploaded APK for scan {apk_savelocation}: {e}")
            if os.path.exists(apk_savelocation): os.remove(apk_savelocation) # Clean up partially saved file
            session.pop('temp_apk_for_scan_path', None)
            return jsonify({'success': False, 'error': f'Error saving uploaded file: {e}'})

        binary_manifest_url = None
        binary_manifest_filename_ondisk = None
        found_file_versions = {}
        warning_message_from_backend = None

        try:
            # Extract binary manifest (optional feature)
            binary_manifest_filename_gen = f"AndroidManifest_binary_{secrets.token_hex(4)}.xml"
            extracted_binary_path = extract_binary_manifest_from_apk(apk_savelocation, binary_manifest_filename_gen)
            if extracted_binary_path:
                binary_manifest_url = f"/download/{binary_manifest_filename_gen}"
                binary_manifest_filename_ondisk = binary_manifest_filename_gen
                # Add to session's downloadable files list
                session.setdefault('downloadable_files', {})[binary_manifest_filename_gen] = extracted_binary_path

            # Scan for library versions
            found_file_versions = scan_archive_for_files(apk_savelocation, TARGET_APK_FILE_PATHS)
            if 'error' in found_file_versions: # If scan_archive_for_files returned an error object
                return jsonify({'success': False, 'error': found_file_versions['error']})

        except Exception as e:
            logger.error(f"Error during APK processing after save {apk_savelocation}: {e}")
            return jsonify({'success': False, 'error': f'An unexpected error occurred during APK processing: {str(e)}'})
        finally:
            # Clean up the uploaded APK for scan immediately after processing
            apk_path_to_clean = session.pop('temp_apk_for_scan_path', None)
            if apk_path_to_clean and os.path.exists(apk_path_to_clean):
                try:
                    os.remove(apk_path_to_clean)
                    logger.info(f"Cleaned up temp APK for scan: {apk_path_to_clean}")
                except OSError as e:
                    logger.error(f"Error removing temp APK for scan {apk_path_to_clean}: {e}")

        response_data = {
            'success': True,
            'binary_manifest_url': binary_manifest_url,
            'binary_manifest_filename': binary_manifest_filename_ondisk,
            'original_apk_name': original_apk_name,
            'found_file_versions': found_file_versions
        }
        if warning_message_from_backend: # Though not set in this specific path currently
            response_data['warning_message'] = warning_message_from_backend
        return jsonify(response_data)
    else:
        return jsonify({'success': False, 'error': 'Invalid file type. Please upload an .apk file.'})


@app.route('/analyze_gradle', methods=['POST'])
def analyze_gradle():
    """Handles Gradle file upload for library version analysis."""
    if 'gradle_file' not in request.files:
        return jsonify({'success': False, 'error': 'No gradle file part'})
    file = request.files['gradle_file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No selected gradle file'})

    expected_versions_input = request.form.get('expected_versions', '')

    if file and file.filename.endswith('.gradle'):
        try:
            gradle_content = file.read().decode('utf-8')

            # Load the latest mapping so library additions/removals do not
            # require changing or rebuilding the Services Checker code.
            gradle_mapping = _load_gradle_lib_mapping()
            found_versions = scan_gradle_for_versions(gradle_content, gradle_mapping)

            return jsonify({
                'success': True,
                'found_versions': found_versions,
                'expected_versions_input': expected_versions_input
            })

        except Exception as e:
            logger.error(f"Error processing gradle file: {e}", exc_info=True)
            return jsonify({'success': False, 'error': f'An error occurred processing the gradle file: {str(e)}'})
    else:
        return jsonify({'success': False, 'error': 'Invalid file type. Please upload a .gradle file.'})


@app.route('/analyze_podfile', methods=['POST'])
def analyze_podfile():
    """Handles Podfile upload for library version analysis."""
    if 'podfile_file' not in request.files:
        return jsonify({'success': False, 'error': 'No Podfile part'})
    file = request.files['podfile_file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No selected Podfile'})

    expected_versions_input = request.form.get('expected_versions', '')

    try:
        podfile_content = file.read().decode('utf-8')

        # Load the latest Git mapping so Pod additions/removals do not require
        # changing or rebuilding the Services Checker code.
        podfile_mapping = _load_podfile_lib_mapping()
        found_versions = scan_podfile_for_versions(podfile_content, podfile_mapping)

        return jsonify({
            'success': True,
            'found_versions': found_versions,
            'expected_versions_input': expected_versions_input
        })

    except Exception as e:
        logger.error(f"Error processing Podfile: {e}", exc_info=True)
        return jsonify({'success': False, 'error': f'An error occurred processing the Podfile: {str(e)}'})


@app.route('/analyze_manifest_data', methods=['POST'])
def analyze_manifest_data_route():
    """Handles APK upload for manifest permission and metadata analysis."""
    if 'apk_file' not in request.files:
        return jsonify({'success': False, 'error': 'No APK file part for manifest analysis.'})
    file = request.files['apk_file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No selected APK file for manifest analysis.'})

    if file and file.filename.endswith('.apk'):
        original_apk_name = file.filename
        manifest_preset_name = request.form.get('preset_name', '').strip()
        manifest_preset = _load_manifest_check_presets().get(manifest_preset_name)
        temp_apk_manifest_filename = f"temp_apk_manifest_{secrets.token_hex(8)}_{original_apk_name}"
        apk_savelocation = os.path.join(app.config['UPLOAD_FOLDER'], temp_apk_manifest_filename)
        session['temp_apk_for_manifest_path'] = apk_savelocation # Track for cleanup

        try:
            file.save(apk_savelocation)
            logger.info(f"APK for manifest analysis saved to: {apk_savelocation}")
        except Exception as e:
            logger.error(f"Error saving uploaded APK for manifest analysis {apk_savelocation}: {e}")
            if os.path.exists(apk_savelocation): os.remove(apk_savelocation)
            session.pop('temp_apk_for_manifest_path', None)
            return jsonify({'success': False, 'error': f'Error saving uploaded file: {str(e)}'})

        permissions = []
        metadata = {}
        manifest_attributes = {}
        error_from_extraction = None
        try:
            permissions, metadata, manifest_attributes, error_from_extraction = extract_manifest_data(apk_savelocation)

            # If androguard was not available, error_from_extraction would be set by extract_manifest_data
            if not androguard_axml_available:
                # Ensure the frontend gets the specific error about androguard missing
                final_error = error_from_extraction if error_from_extraction else "AXMLPrinter from androguard.core.axml was not available at script startup."
                # Return success:false so frontend JS handles it as an error display
                return jsonify({'success': False, 'error': final_error, 'apk_filename': original_apk_name})

            # If manifest was not found, it's an error for this specific operation
            if error_from_extraction and not permissions and not metadata and not manifest_attributes and "not found in the APK" in error_from_extraction:
                 return jsonify({'success': False, 'error': error_from_extraction, 'apk_filename': original_apk_name})

        except Exception as e:
            logger.error(f"Unexpected error in /analyze_manifest_data route for {apk_savelocation}: {e}", exc_info=True)
            return jsonify({'success': False, 'error': f'An unexpected server error occurred: {str(e)}', 'apk_filename': original_apk_name})
        finally:
            # Clean up the uploaded APK for manifest analysis immediately
            apk_path_to_clean = session.pop('temp_apk_for_manifest_path', None)
            if apk_path_to_clean and os.path.exists(apk_path_to_clean):
                try:
                    os.remove(apk_path_to_clean)
                    logger.info(f"Cleaned up temp APK for manifest: {apk_path_to_clean}")
                except OSError as e:
                    logger.error(f"Error removing temp APK for manifest {apk_path_to_clean}: {e}")

        # If we reach here, the process itself succeeded, even if no data was found or a non-critical error occurred during extraction
        return jsonify({
            'success': True, # Indicates the route itself processed successfully
            'permissions': permissions,
            'metadata': metadata,
            'manifest_attributes': manifest_attributes,
            'apk_filename': original_apk_name,
            'manifest_preset_name': manifest_preset_name,
            'manifest_comparison': compare_manifest_data(
                manifest_attributes,
                permissions,
                metadata,
                manifest_preset,
            ),
            'error_message_from_backend': error_from_extraction # Pass any non-fatal errors from extraction to frontend
        })
    else:
        return jsonify({'success': False, 'error': 'Invalid file type. Please upload an .apk file for manifest analysis.'})


@app.route('/convert_aab_to_apk', methods=['POST'])
def convert_aab_to_apk():
    """Handles AAB file upload and converts it to a universal APK."""
    if 'aab_file' not in request.files:
        return jsonify({'success': False, 'error': 'No AAB file part'})
    file = request.files['aab_file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No selected AAB file'})

    _cleanup_stale_conversion_artifacts()

    # Resolve resources for this request. The service can be loaded from a
    # downloaded update directory while its bundled assets remain in the
    # frozen application, so import-time paths are not reliable here.
    bundletool_path = _resolve_bundletool_path()
    if not os.path.isfile(bundletool_path):
        logger.error(f"Bundletool not found at {bundletool_path}")
        return jsonify({
            'success': False,
            'error': (
                'Server configuration error: bundletool is missing from the '
                'installed app resources. Please update/reinstall the app.'
            ),
        })

    try:
        free_space = shutil.disk_usage(app.config['UPLOAD_FOLDER']).free
    except OSError as error:
        logger.error(f"Could not check free disk space before conversion: {error}")
        free_space = None
    if free_space is not None and free_space < MIN_CONVERSION_FREE_SPACE_BYTES:
        free_space_mb = free_space // (1024 * 1024)
        logger.error(f"Not enough free disk space for AAB conversion: {free_space_mb} MB")
        return jsonify({
            'success': False,
            'error': (
                f'Not enough free disk space to convert this AAB ({free_space_mb} MB free). '
                'Remove old generated APKs and try again.'
            ),
        })

    if file and str(file.filename).lower().endswith('.aab'):
        original_aab_name = _safe_download_filename(file.filename, 'uploaded.aab')
        token = secrets.token_hex(8) # Unique token for this conversion

        # Define temporary file names using the token
        temp_aab_filename = f"temp_aab_convert_{token}_{original_aab_name}"
        aab_savelocation = os.path.join(app.config['UPLOAD_FOLDER'], temp_aab_filename)

        apks_filename = f"output_convert_{token}.apks" # Intermediate .apks file
        apks_savelocation = os.path.join(app.config['UPLOAD_FOLDER'], apks_filename)

        # Define the final APK filename. The APK itself is reserved and
        # written in Downloads after bundletool succeeds; it is never staged
        # as a generated file in UPLOAD_FOLDER.
        base_aab_name = os.path.splitext(original_aab_name)[0] or 'converted'
        base_aab_name = _safe_download_filename(base_aab_name, 'converted')
        final_apk_filename = f"{base_aab_name}_universal_{token[:4]}.apk" # Shortened token for readability

        # Track files for cleanup in session (though some are cleaned immediately in finally)
        session['temp_aab_for_convert_path'] = aab_savelocation
        session['temp_apks_for_convert_path'] = apks_savelocation

        bundletool_stderr_output = ""
        warning_message_from_backend = None
        download_apk_path = None
        saved_filename = ''
        download_succeeded = False

        # Keystore details
        ks_path_input = None
        ks_pass = None
        ks_alias = None
        key_pass = None

        if USE_HARDCODED_KEYSTORE:
            ks_path_input = DEFAULT_KEYSTORE_FILENAME # Assumed relative to APP_ROOT
            ks_pass = DEFAULT_KEYSTORE_PASS
            ks_alias = DEFAULT_KEYSTORE_ALIAS
            key_pass = DEFAULT_KEY_PASS if DEFAULT_KEY_PASS else ks_pass # Use keystore_pass if key_pass is not specifically set
            logger.info("Using hardcoded keystore information for AAB conversion.")
        else:
            ks_path_input = request.form.get('keystore_path')
            ks_pass = request.form.get('keystore_pass')
            ks_alias = request.form.get('keystore_alias')
            key_pass = request.form.get('key_pass', ks_pass) # Default to ks_pass if key_pass is empty
            logger.info(f"Received keystore_path input from form: '{ks_path_input}'")
            logger.info(f"Received keystore_alias input from form: '{ks_alias}'")


        final_ks_path = None
        if ks_path_input: # Only proceed if ks_path_input is not None or empty
            if os.path.isabs(ks_path_input):
                final_ks_path = ks_path_input
            elif os.path.basename(os.path.normpath(ks_path_input)) == DEFAULT_KEYSTORE_FILENAME:
                final_ks_path = _resolve_keystore_path()
            else:
                # Custom relative keystores remain relative to the active
                # service app directory.
                final_ks_path = os.path.join(APP_ROOT, ks_path_input)
            logger.info(f"Resolved final_ks_path: '{final_ks_path}'")


        try:
            file.save(aab_savelocation)
            logger.info(f"AAB file for conversion saved to: {aab_savelocation}")

            # Construct bundletool command
            cmd = [
                'java', '-jar', bundletool_path,
                'build-apks',
                '--bundle=' + aab_savelocation,
                '--output=' + apks_savelocation,
                '--mode=universal', # Generate a universal APK
                '--overwrite' # Overwrite output if it exists
            ]

            # Signing logic
            # Check if we have enough details for custom/hardcoded signing
            if final_ks_path and ks_pass and ks_alias:
                if not os.path.exists(final_ks_path):
                    logger.error(f"Custom/Hardcoded keystore not found at resolved path: {final_ks_path}")
                    return jsonify({'success': False, 'error': f'Keystore not found at the resolved path: {final_ks_path}. Ensure the path/filename is correct and accessible.'})

                logger.info(f"Attempting to use keystore: {final_ks_path} with alias: {ks_alias}")
                cmd.extend([
                    '--ks=' + final_ks_path,
                    '--ks-key-alias=' + ks_alias,
                    '--ks-pass=pass:' + ks_pass,
                    '--key-pass=pass:' + (key_pass if key_pass else ks_pass) # Use ks_pass if key_pass is empty
                ])
            else: # Fallback to debug keystore if custom/hardcoded details are insufficient or not provided
                logger.info("Keystore details (custom/hardcoded) incomplete or not provided. Attempting to use default debug keystore.")
                debug_keystore = get_debug_keystore_path()
                if debug_keystore:
                    logger.info(f"Using debug keystore: {debug_keystore}")
                    cmd.extend([
                        '--ks=' + debug_keystore,
                        '--ks-key-alias=androiddebugkey',
                        '--ks-pass=pass:android',
                        '--key-pass=pass:android'
                    ])
                else:
                    logger.warning("Neither custom/hardcoded keystore details fully provided nor default debug.keystore found. Bundletool will attempt default signing, which may result in an unsigned/uninstallable APK.")
                    warning_message_from_backend = "No custom/hardcoded keystore provided (or details incomplete) and standard debug.keystore not found. The generated APK might not be signed correctly and may fail to install. Check server logs for bundletool output."


            logger.info(f"Running bundletool command (first few parts): {' '.join(cmd[:7])} ...") # Log part of command, excluding passwords
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate(timeout=180) # 3 minute timeout

            if stdout:
                logger.info(f"Bundletool stdout: {stdout.decode('utf-8', errors='ignore')}")
            if stderr:
                bundletool_stderr_output = stderr.decode('utf-8', errors='ignore')
                if process.returncode != 0: # Log as error only if bundletool failed
                    logger.error(f"Bundletool stderr: {bundletool_stderr_output}")
                else: # Log as warning if bundletool succeeded but still had stderr output
                    logger.warning(f"Bundletool stderr (return code 0): {bundletool_stderr_output}")

            if process.returncode != 0:
                error_message = f"Bundletool failed. Exit code: {process.returncode}. STDERR: {bundletool_stderr_output if bundletool_stderr_output else 'N/A'}"
                logger.error(error_message)
                return jsonify({'success': False, 'error': "Bundletool conversion failed. Please check server logs for details.", 'bundletool_stderr': bundletool_stderr_output})

            logger.info(f"APKS file generated successfully: {apks_savelocation}")

            # Extract universal.apk directly into the user's Downloads folder.
            # The .aab upload and .apks archive remain short-lived conversion
            # intermediates, but the final APK never touches UPLOAD_FOLDER or
            # the Flask session download cache.
            with zipfile.ZipFile(apks_savelocation, 'r') as apks_zip:
                if 'universal.apk' in apks_zip.namelist():
                    download_apk_path, saved_filename = _write_zip_member_to_downloads(
                        apks_zip,
                        'universal.apk',
                        final_apk_filename,
                    )
                    logger.info(f"Universal APK written directly to Downloads: {download_apk_path}")
                else:
                    logger.error(f"'universal.apk' not found in {apks_savelocation}")
                    return jsonify({'success': False, 'error': "'universal.apk' not found in the generated APKS file."})

            download_succeeded = True
            response_data = {
                'success': True,
                'apk_filename': final_apk_filename,
                'saved_to_downloads': True,
                'saved_filename': saved_filename,
                'saved_directory': os.path.dirname(download_apk_path),
            }
            if warning_message_from_backend:
                response_data['warning_message'] = warning_message_from_backend
            return jsonify(response_data)

        except subprocess.TimeoutExpired:
            logger.error("Bundletool command timed out.")
            return jsonify({'success': False, 'error': 'AAB to APK conversion timed out. The AAB might be too large or complex.'})
        except (OSError, ValueError) as error:
            logger.error(f"Could not save the converted APK directly to Downloads: {error}", exc_info=True)
            return jsonify({'success': False, 'error': f'Could not save the converted APK directly to Downloads: {error}'})
        except Exception as e:
            logger.error(f"Error during AAB to APK conversion: {e}", exc_info=True)
            return jsonify({'success': False, 'error': f'An unexpected error occurred: {str(e)}'})
        finally:
            if not download_succeeded:
                _remove_download_file(download_apk_path)
            # Clean up temporary AAB and APKS files immediately after conversion attempt
            aab_path_to_clean = session.pop('temp_aab_for_convert_path', None)
            apks_path_to_clean = session.pop('temp_apks_for_convert_path', None)
            files_to_remove_immediately = [aab_path_to_clean, apks_path_to_clean]
            for f_path in files_to_remove_immediately:
                if f_path and os.path.exists(f_path):
                    try:
                        os.remove(f_path)
                        logger.info(f"Cleaned up AAB/APKS temp file: {f_path}")
                    except OSError as e_os:
                        logger.error(f"Error removing immediate temp file {f_path}: {e_os}")
    else:
        return jsonify({'success': False, 'error': 'Invalid file type. Please upload an .aab file.'})


@app.route('/download/<filename>')
def download_file(filename):
    """Serves files marked as downloadable in the session."""
    logger.info(f"Download request received for filename: '{filename}'")

    # Retrieve the dictionary of downloadable files from the session
    downloadable_files = session.get('downloadable_files', {})
    logger.debug(f"Current session['downloadable_files']: {downloadable_files}")

    # Get the actual file path on disk using the filename from the URL as a key
    file_path_on_disk = downloadable_files.get(filename)
    logger.debug(f"Retrieved file_path_on_disk from session for '{filename}': '{file_path_on_disk}'")

    # Basic security check for filename
    if '..' in filename or filename.startswith('/'):
        logger.warning(f"Invalid filename attempt: {filename}")
        return "Invalid filename.", 400

    if file_path_on_disk:
        # Security check: ensure the path from session is within the UPLOAD_FOLDER
        if not os.path.abspath(file_path_on_disk).startswith(os.path.abspath(app.config['UPLOAD_FOLDER'])):
             logger.error(f"Security Error: Path from session '{file_path_on_disk}' is outside UPLOAD_FOLDER '{app.config['UPLOAD_FOLDER']}'. Denying download.")
             return "Access denied.", 403

        logger.info(f"Attempting to serve file. Path from session: '{file_path_on_disk}'")
        logger.info(f"Does file exist at this path? {os.path.exists(file_path_on_disk)}")

        actual_filename_to_serve = os.path.basename(file_path_on_disk)

        logger.debug(f"Actual filename to serve: '{actual_filename_to_serve}'")

        # It's good practice that the filename from URL matches the basename of the path from session
        if actual_filename_to_serve != filename:
             logger.warning(f"Filename mismatch warning: URL filename ('{filename}') differs from session path basename ('{actual_filename_to_serve}'). Serving based on session path.")

        if os.path.exists(file_path_on_disk):
            logger.info(f"File confirmed to exist at: '{file_path_on_disk}' using path from session.")
            try:
                response = send_file(
                    file_path_on_disk,
                    mimetype=(
                        'application/vnd.android.package-archive'
                        if actual_filename_to_serve.lower().endswith('.apk')
                        else None
                    ),
                    as_attachment=True,
                    download_name=actual_filename_to_serve,
                    conditional=True,
                )
                # The native WebView path uses /save_download, but regular
                # browsers use this route. Remove the generated APK after the
                # response is sent so repeated conversions cannot fill disk.
                response.call_on_close(lambda: _remove_generated_file(file_path_on_disk))
                downloadable_files = dict(session.get('downloadable_files', {}))
                downloadable_files.pop(filename, None)
                session['downloadable_files'] = downloadable_files
                return response
            except Exception as e:
                logger.error(f"Error during send_file: {e}", exc_info=True)
                return "Error serving file.", 500
        else:
            logger.error(f"File not found on disk at path from session: '{file_path_on_disk}' (for requested filename '{filename}')")
            return "File not found on server (it may have been cleaned up).", 404
    else:
        logger.warning(f"Filename '{filename}' not found as a key in session's downloadable_files.")
        return "File not found or access denied (invalid link or session expired).", 404


@app.route('/save_download/<filename>', methods=['POST'])
def save_download_file(filename):
    """Copies a generated file to the current user's Downloads directory.

    This avoids relying on native WebView download support, which differs
    between WebKit on macOS and Qt WebEngine on Windows.
    """
    logger.info(f"Save-to-Downloads request received for filename: '{filename}'")

    if '..' in filename or filename.startswith('/') or os.path.basename(filename) != filename:
        return jsonify({'success': False, 'error': 'Invalid filename.'}), 400

    downloadable_files = session.get('downloadable_files', {})
    file_path_on_disk = downloadable_files.get(filename)
    if not file_path_on_disk:
        return jsonify({'success': False, 'error': 'File not found or the conversion session expired.'}), 404

    upload_root = os.path.abspath(app.config['UPLOAD_FOLDER'])
    source_path = os.path.abspath(file_path_on_disk)
    try:
        inside_upload_folder = os.path.commonpath([upload_root, source_path]) == upload_root
    except ValueError:
        inside_upload_folder = False
    if not inside_upload_folder:
        logger.error(f"Security Error: refusing to copy file outside UPLOAD_FOLDER: {source_path}")
        return jsonify({'success': False, 'error': 'Invalid generated file path.'}), 403
    if not os.path.isfile(source_path):
        return jsonify({'success': False, 'error': 'Generated APK is no longer available on the server.'}), 404

    try:
        destination = _copy_generated_file_to_downloads(source_path, filename)
    except (OSError, ValueError) as error:
        logger.error(f"Failed to save generated file to Downloads: {error}", exc_info=True)
        return jsonify({'success': False, 'error': f'Could not save APK to Downloads: {error}'}), 500

    _remove_generated_file(source_path)
    downloadable_files = dict(session.get('downloadable_files', {}))
    downloadable_files.pop(filename, None)
    session['downloadable_files'] = downloadable_files
    return jsonify({
        'success': True,
        'filename': os.path.basename(destination),
        'directory': os.path.dirname(destination),
    })

def cleanup_session_files():
    """Cleans up all files tracked via session keys."""
    logger.info("Running cleanup_session_files...")
    count = 0

    # Clean up files from 'downloadable_files' (e.g., generated APKs, binary manifests)
    downloadable_files = session.pop('downloadable_files', {})
    for filename, file_path in downloadable_files.items():
        if file_path and os.path.exists(file_path):
            # Security check: only delete if within UPLOAD_FOLDER
            if os.path.abspath(file_path).startswith(os.path.abspath(app.config['UPLOAD_FOLDER'])):
                try:
                    os.remove(file_path)
                    logger.info(f"Cleaned up session downloadable file: {file_path}")
                    count +=1
                except OSError as e:
                    logger.error(f"Error during session cleanup of downloadable {file_path}: {e}")
            else:
                 logger.warning(f"Skipping cleanup of downloadable file outside UPLOAD_FOLDER: {file_path}")

    # Clean up other temporary files tracked by specific session keys (uploaded APKs/AABs that were processed)
    other_temp_paths_keys = [
        'temp_apk_for_scan_path',
        'temp_apk_for_manifest_path',
        'temp_aab_for_convert_path',  # Original uploaded AAB
        'temp_apks_for_convert_path'  # Intermediate .apks file
    ]
    for key in other_temp_paths_keys:
        path_to_clean = session.pop(key, None)
        if path_to_clean and os.path.exists(path_to_clean):
            if os.path.abspath(path_to_clean).startswith(os.path.abspath(app.config['UPLOAD_FOLDER'])):
                try:
                    os.remove(path_to_clean)
                    logger.info(f"Cleaned up session temp file by key {key}: {path_to_clean}")
                    count +=1
                except OSError as e:
                    logger.error(f"Error during session key cleanup of {path_to_clean}: {e}")
            else:
                 logger.warning(f"Skipping cleanup of temp file outside UPLOAD_FOLDER (key: {key}): {path_to_clean}")

    # Deprecated 'temp_files_to_clean' list, but keep for backward compatibility if old sessions exist
    temp_files_list = session.pop('temp_files_to_clean', [])
    for file_path in temp_files_list:
        if file_path and os.path.exists(file_path):
            if os.path.abspath(file_path).startswith(os.path.abspath(app.config['UPLOAD_FOLDER'])):
                try:
                    os.remove(file_path)
                    logger.info(f"Cleaned up temp_files_to_clean item: {file_path}")
                    count += 1
                except OSError as e:
                    logger.error(f"Error cleaning temp_files_to_clean item {file_path}: {e}")
            else:
                logger.warning(f"Skipping cleanup of temp_files_to_clean item outside UPLOAD_FOLDER: {file_path}")

    # Clear other session variables if they exist
    session.pop('original_apk_filename', None)
    session.pop('original_aab_name', None)
    logger.info(f"Cleanup complete. Removed {count} files.")


@app.route('/cleanup', methods=['POST'])
def cleanup_files_endpoint():
    """ Endpoint to clean up session files, typically called on reset/page load. """
    cleanup_session_files()
    return jsonify({'success': True, 'message': 'Cleanup attempt complete.'})


@app.get('/health')
def health_check():
    return jsonify({'ok': True, 'service': 'services-checker'})


# --- Main Execution ---
if __name__ == '__main__':
    # Ensure upload folder exists
    if not os.path.exists(UPLOAD_FOLDER_ABS_PATH):
        os.makedirs(UPLOAD_FOLDER_ABS_PATH)
    # Resolve again at startup so a frozen/update bundle cannot retain a stale
    # path calculated during an earlier import.
    runtime_bundletool_path = _resolve_bundletool_path()
    if not os.path.exists(runtime_bundletool_path):
        logger.critical(
            f"CRITICAL: bundletool.jar not found at '{runtime_bundletool_path}'. "
            "AAB to APK conversion will FAIL."
        )
        logger.critical(
            "Please place bundletool-all-1.18.1.jar in the bundled "
            "services_checker resource directory or set "
            "EVENTINSPECTOR_BUNDLETOOL_PATH."
        )

    # Reminder about androguard if it wasn't imported
    if not androguard_axml_available:
        logger.warning("Reminder: AXMLPrinter from 'androguard.core.axml' was not imported successfully. Manifest analysis may fail for binary XML.")

    service_host = os.getenv('SERVICES_CHECKER_HOST', '127.0.0.1')
    service_port = int(os.getenv('SERVICES_CHECKER_PORT', '5010'))
    app.run(debug=False, host=service_host, port=service_port, use_reloader=False)
