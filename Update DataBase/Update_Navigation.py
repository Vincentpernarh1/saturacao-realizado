import dotenv
import os
import sys
import getpass
import json
import re
import time
from playwright.sync_api import sync_playwright
from datetime import datetime
import shutil

# PyInstaller-compatible path resolution
if getattr(sys, 'frozen', False):
    # Running from PyInstaller .exe
    script_dir = sys._MEIPASS
else:
    # Running from source
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Load environment variables using absolute path
env_path = os.path.join(script_dir, "BD", ".env")
if not os.path.exists(env_path):
    # Fallback for source code structure
    env_path = os.path.join(script_dir, "..", "BD", ".env")
    env_path = os.path.abspath(env_path)

dotenv.load_dotenv(dotenv_path=env_path)
sharepoint_url = os.getenv("SHAREPOINT_URL")

if not sharepoint_url:
    raise ValueError(f"SHAREPOINT_URL not found in .env file at {env_path}")

print(f"SharePoint URL: {sharepoint_url}")

# Create the download folder path - BD folder should be alongside the .exe or in source structure
current_date = datetime.now().strftime("%Y-%m-%d")
if getattr(sys, 'frozen', False):
    # When running from .exe, BD folder is alongside the executable
    download_folder = os.path.join(os.path.dirname(sys.executable), "BD")
else:
    # When running from source
    download_folder = os.path.join(script_dir, "..", "BD")
    download_folder = os.path.abspath(download_folder)

# Ensure the download folder exists
os.makedirs(download_folder, exist_ok=True)

print(f"Download folder: {download_folder}")

# Use AppData for automation profile — one profile per Windows user so SSO sessions are preserved
# LOCALAPPDATA works in all Windows languages (Portuguese, English, etc.) - it's a system variable
# Profile path: %LOCALAPPDATA%\Viajante\edge_automation_profile\<username>

# Get Windows username - try environment variable first, then getpass as fallback
windows_username = os.getenv('USERNAME') or getpass.getuser() or 'default'

# LOCALAPPDATA is safe - it exists in all Windows versions/languages (C:\Users\<user>\AppData\Local)
local_appdata = os.getenv('LOCALAPPDATA')
if not local_appdata:
    # Fallback: construct manually if LOCALAPPDATA is somehow missing
    user_profile = os.getenv('USERPROFILE') or os.path.expanduser('~')
    local_appdata = os.path.join(user_profile, 'AppData', 'Local')

automation_profile = os.path.join(local_appdata, 'Viajante', 'edge_automation_profile', windows_username)
os.makedirs(automation_profile, exist_ok=True)

# Source Edge profile (already logged-in) used to seed temp profile for SSO reuse
source_edge_user_data = os.path.join(local_appdata, 'Microsoft', 'Edge', 'User Data')

# Temp profile always uses Default directory inside the automation user data dir
automation_profile_directory = 'Default'
automation_profile_path = os.path.join(automation_profile, automation_profile_directory)

# If true, refresh temp profile from source Edge profile every run
edge_refresh_temp_profile = os.getenv('EDGE_REFRESH_TEMP_PROFILE', '1').strip().lower() in ('1', 'true', 'yes')


def _detect_corporate_edge_profile(user_data_dir, sharepoint_url):
    """Scan all Edge profiles and return the directory name of the one
    signed in with the corporate account matching the SharePoint tenant.
    Falls back to EDGE_PROFILE_DIRECTORY env var, then 'Default'.
    """
    # Derive expected tenant domain from SharePoint URL
    # e.g. https://shiftup.sharepoint.com -> shiftup
    tenant_match = re.search(r'https?://([^.]+)\.sharepoint\.com', sharepoint_url or '')
    tenant = tenant_match.group(1).lower() if tenant_match else None

    if not os.path.isdir(user_data_dir):
        return os.getenv('EDGE_PROFILE_DIRECTORY', 'Default').strip() or 'Default'

    candidates = []
    for entry in os.listdir(user_data_dir):
        if entry == 'Default' or re.match(r'^Profile \d+$', entry):
            prefs_path = os.path.join(user_data_dir, entry, 'Preferences')
            if not os.path.isfile(prefs_path):
                continue
            try:
                with open(prefs_path, 'r', encoding='utf-8', errors='ignore') as f:
                    prefs = json.load(f)
                accounts = (
                    prefs.get('account_info', [])
                    or prefs.get('signin', {}).get('allowed_usernames', [])
                )
                if not isinstance(accounts, list):
                    accounts = []
                for acc in accounts:
                    email = ''
                    if isinstance(acc, dict):
                        email = (acc.get('email') or acc.get('full_name') or '').lower()
                    elif isinstance(acc, str):
                        email = acc.lower()
                    if tenant and tenant in email:
                        return entry  # exact tenant match — use immediately
                    if email:
                        candidates.append((entry, email))
            except Exception:
                continue

    # No exact tenant match — prefer a profile with an @<org> account over personal ones
    personal_domains = {'gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com', 'live.com'}
    for profile_dir, email in candidates:
        domain = email.split('@')[-1] if '@' in email else ''
        if domain and domain not in personal_domains:
            return profile_dir

    # Fall back to env var or Default
    return os.getenv('EDGE_PROFILE_DIRECTORY', 'Default').strip() or 'Default'


source_edge_profile_directory = _detect_corporate_edge_profile(source_edge_user_data, sharepoint_url)
source_edge_profile_path = os.path.join(source_edge_user_data, source_edge_profile_directory)

print(f"Edge source profile: {source_edge_profile_path} (profile: {source_edge_profile_directory})")
print(f"Edge automation profile folder: {automation_profile_path}")

# Files to download
FILES_TO_DOWNLOAD = [
    "BD_CADASTRO_PN",
    "BD_CADASTRO_MDR"
]


def _copy_file_tolerant(src, dst):
    """Copy a single file, silently skipping if locked or missing."""
    try:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
    except Exception:
        pass


def _copy_dir_tolerant(src_dir, dst_dir):
    """Shallow-copy all files in src_dir to dst_dir, skipping locked ones."""
    if not os.path.isdir(src_dir):
        return
    os.makedirs(dst_dir, exist_ok=True)
    for name in os.listdir(src_dir):
        src = os.path.join(src_dir, name)
        dst = os.path.join(dst_dir, name)
        if os.path.isfile(src):
            _copy_file_tolerant(src, dst)
        elif os.path.isdir(src):
            _copy_dir_tolerant(src, dst)


def seed_automation_profile_from_edge(silent=False):
    """Seed temp profile with only the auth-relevant files from the logged-in Edge profile.

    Skips large cache directories so startup is fast even on slow machines.
    """
    if not os.path.isdir(source_edge_user_data):
        if not silent:
            print(f"Edge source user data not found: {source_edge_user_data}")
        return False

    if not os.path.isdir(source_edge_profile_path):
        if not silent:
            print(f"Edge source profile not found: {source_edge_profile_path}")
        return False

    need_seed = edge_refresh_temp_profile or (not os.path.isdir(automation_profile_path))
    if not need_seed:
        return True

    if not silent:
        print("Seeding temporary profile from logged-in Edge profile...")

    if edge_refresh_temp_profile and os.path.isdir(automation_profile):
        try:
            shutil.rmtree(automation_profile, ignore_errors=True)
        except Exception:
            pass

    os.makedirs(automation_profile, exist_ok=True)
    os.makedirs(automation_profile_path, exist_ok=True)

    # ── User Data root ──────────────────────────────────────────────────────
    # Local State is required by Chromium for profile internals.
    _copy_file_tolerant(
        os.path.join(source_edge_user_data, 'Local State'),
        os.path.join(automation_profile, 'Local State'),
    )

    # ── Profile-level auth files (skip large cache/GPU directories) ─────────
    src = source_edge_profile_path
    dst = automation_profile_path

    # Single files that carry session cookies and auth data
    for fname in ('Cookies', 'Login Data', 'Login Data For Account',
                  'Web Data', 'Preferences', 'Secure Preferences'):
        _copy_file_tolerant(os.path.join(src, fname), os.path.join(dst, fname))

    # Network sub-folder (newer Chromium stores cookies here too)
    _copy_dir_tolerant(os.path.join(src, 'Network'), os.path.join(dst, 'Network'))

    # Local Storage and IndexedDB carry MSAL/OAuth tokens for Microsoft SSO
    _copy_dir_tolerant(os.path.join(src, 'Local Storage'), os.path.join(dst, 'Local Storage'))
    _copy_dir_tolerant(os.path.join(src, 'IndexedDB'),     os.path.join(dst, 'IndexedDB'))
    _copy_dir_tolerant(os.path.join(src, 'Session Storage'), os.path.join(dst, 'Session Storage'))

    if not silent:
        print("Temporary profile is ready.")

    return True


def wait_for_sharepoint_ready(page, silent=False, progress_callback=None, timeout_seconds=180):
    """Wait until SharePoint library is available or login timeout is reached."""
    if not silent:
        print("Waiting for SharePoint authentication/library readiness...")

    if progress_callback:
        progress_callback("Aguardando autenticacao no SharePoint...")

    login_hint_printed = False
    start = time.time()

    while (time.time() - start) < timeout_seconds:
        try:
            # If one of the target files is visible, the library is ready.
            for filename in FILES_TO_DOWNLOAD:
                if page.locator(f"text={filename}").first.is_visible(timeout=500):
                    if not silent:
                        print("SharePoint library is ready.")
                    return True

            current_url = (page.url or "").lower()
            looks_like_login = (
                "login.microsoftonline.com" in current_url
                or "microsoftonline.com" in current_url
                or page.locator("text=Sign in").first.is_visible(timeout=300)
                or page.locator("text=Entrar").first.is_visible(timeout=300)
                or page.locator("text=Use another account").first.is_visible(timeout=300)
                or page.locator("text=Usar outra conta").first.is_visible(timeout=300)
            )

            if looks_like_login and not login_hint_printed:
                if not silent:
                    print("Authentication page detected. Please complete Microsoft login in the opened Edge window.")
                if progress_callback:
                    progress_callback("Finalize o login Microsoft na janela do Edge para continuar...")
                login_hint_printed = True

            # Also treat the page as ready when we can identify the documents area.
            if (
                page.locator("text=Shared Documents").first.is_visible(timeout=300)
                or page.locator("text=Documentos Compartilhados").first.is_visible(timeout=300)
                or page.locator("text=Files").first.is_visible(timeout=300)
                or page.locator("text=Arquivos").first.is_visible(timeout=300)
            ):
                if not silent:
                    print("SharePoint documents page detected.")
                return True

        except Exception:
            pass

        page.wait_for_timeout(1000)

    if not silent:
        print(f"Timeout waiting for SharePoint readiness ({timeout_seconds}s).")
    if progress_callback:
        progress_callback("Tempo esgotado aguardando autenticacao SharePoint.")
    return False


def navigate_to_sharepoint(context, url, silent=False, max_attempts=3):
    """Navigate with retries to handle transient ERR_ABORTED/closed-page cases."""
    last_error = None
    page = None

    for attempt in range(1, max_attempts + 1):
        page = context.new_page()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            return page
        except Exception as e:
            last_error = e
            err_text = str(e)
            if not silent:
                print(f"Navigation attempt {attempt}/{max_attempts} failed: {err_text}")

            # Close failed page before retrying.
            try:
                page.close()
            except Exception:
                pass

            should_retry = (
                "ERR_ABORTED" in err_text
                or "Target page, context or browser has been closed" in err_text
            )
            if not should_retry or attempt == max_attempts:
                raise

            # Short pause before retrying with a clean page.
            time.sleep(1.5)

    # Defensive: if loop ends unexpectedly, raise the last error.
    if last_error:
        raise last_error
    raise RuntimeError("Unknown navigation failure")

def cleanup_old_versions(filename, current_file, silent=False):
    """Delete old versions of the file, keeping only the newly downloaded one"""
    import glob
    import re
    
    # Find all files in the download folder
    all_files = os.listdir(download_folder)
    
    # Pattern to match: {filename}_YYYY-MM-DD.extension (dated backups only)
    date_pattern = re.compile(rf"^{re.escape(filename)}_\d{{4}}-\d{{2}}-\d{{2}}\..+$")
    
    deleted_count = 0
    for file in all_files:
        # Check if it matches our dated backup pattern
        if date_pattern.match(file):
            file_path = os.path.join(download_folder, file)
            # Skip the file we just downloaded
            if os.path.abspath(file_path) != os.path.abspath(current_file):
                try:
                    os.remove(file_path)
                    if not silent:
                        print(f"  🗑️  Deleted old version: {file}")
                    deleted_count += 1
                except Exception as e:
                    if not silent:
                        print(f"  ⚠️  Could not delete {file}: {e}")
    
    if deleted_count == 0 and not silent:
        print(f"  ℹ️  No old versions to clean up")
    
    return deleted_count

def download_file_from_sharepoint(page, filename, silent=False, progress_callback=None):
    """Download a specific file from the SharePoint page"""
    if not silent:
        print(f"\nLooking for {filename} file...")
    
    if progress_callback:
        progress_callback(f"Procurando arquivo {filename}...")
    
    # Wait a moment for page to be ready
    page.wait_for_timeout(2000)
    
    # Try to find and click the file
    file_found = False
    
    # Try different approaches to find the file
    selectors_to_try = [
        f"button[name*='{filename}']",
        f"a[title*='{filename}']",
        f"[aria-label*='{filename}']",
        f"text={filename}.xlsx",
        f"text={filename}",
    ]
    
    for selector in selectors_to_try:
        try:
            element = page.locator(selector).first
            if element.is_visible(timeout=2000):
                if not silent:
                    print(f"  ✓ Found file using selector: {selector}")
                
                if progress_callback:
                    progress_callback(f"Baixando {filename}...")
                
                # Right-click to open context menu
                element.click(button="right")
                page.wait_for_timeout(1000)
                
                # Look for download option in context menu (EN/PT)
                download_option = page.locator("text=Download").first
                if not download_option.is_visible(timeout=1000):
                    download_option = page.locator("text=Baixar").first
                
                # Click download and wait for the download to start
                with page.expect_download(timeout=30000) as download_info:
                    download_option.click()
                
                download = download_info.value
                file_found = True
                
                # Save the file with a timestamp
                original_filename = download.suggested_filename
                file_extension = original_filename.split('.')[-1] if '.' in original_filename else 'xlsx'
                saved_filename = f"{filename}_{current_date}.{file_extension}"
                save_path = os.path.join(download_folder, saved_filename)
                download.save_as(save_path)
                
                if not silent:
                    print(f"  ✓ Downloaded successfully to: {saved_filename}")
                
                if progress_callback:
                    progress_callback(f"✓ {filename} baixado com sucesso!")
                
                # Clean up old versions
                cleanup_old_versions(filename, save_path, silent=silent)
                
                return True
        except Exception as e:
            continue
    
    if not file_found:
        if not silent:
            print(f"  ✗ Could not find {filename} file")
        return False
    
    return file_found


def download_sharepoint_files(headless=False, silent=False, auto_close=False, progress_callback=None):
    """Main function to download all required files from SharePoint
    
    Args:
        headless: If True, run browser in headless mode (no window)
        silent: If True, suppress print messages
        auto_close: If True, close browser automatically without waiting for input
        progress_callback: Optional function(message) to report progress to GUI
    """
    
    
    with sync_playwright() as p:
        if not seed_automation_profile_from_edge(silent=silent):
            return {f: False for f in FILES_TO_DOWNLOAD}

        # Launch Edge with automation profile
        if not silent:
            print("Launching Edge browser for automation...")
        
        if progress_callback:
            progress_callback("Abrindo navegador Edge...")
        
        # Always use the installed Microsoft Edge so SSO sessions work correctly
        launch_args = [
            f"--profile-directory={automation_profile_directory}",
            "--disable-blink-features=AutomationControlled",
        ]
        if not headless:
            launch_args.append("--start-maximized")

        context = p.chromium.launch_persistent_context(
            user_data_dir=automation_profile,
            headless=headless,
            channel="msedge",
            accept_downloads=True,
            ignore_default_args=["--enable-automation"],
            args=launch_args,
        )
        
        try:
            # Navigate to SharePoint
            if not silent:
                print(f"Navigating to SharePoint...")
                print("\nUsing temporary profile seeded from your Edge SSO session.")
                print()
            
            if progress_callback:
                progress_callback("Conectando ao SharePoint...")

            page = navigate_to_sharepoint(context, sharepoint_url, silent=silent, max_attempts=3)
            if not silent:
                print("Page loaded!")
            
            if progress_callback:
                progress_callback("SharePoint carregado - iniciando downloads...")
            
            # Wait for authentication and library readiness before trying to find files.
            page.wait_for_timeout(2000)
            ready = wait_for_sharepoint_ready(
                page,
                silent=silent,
                progress_callback=progress_callback,
                timeout_seconds=180
            )

            if not ready:
                if not silent:
                    print("Could not confirm authenticated SharePoint access. Aborting downloads.")
                return {f: False for f in FILES_TO_DOWNLOAD}
            
            # Download each file
            results = {}
            for filename in FILES_TO_DOWNLOAD:
                success = download_file_from_sharepoint(page, filename, silent=silent, progress_callback=progress_callback)
                results[filename] = success
            
            # Summary
            if not silent:
                print("\n" + "="*70)
                print("Download Summary:")
                print("="*70)
                for filename, success in results.items():
                    status = "✓ SUCCESS" if success else "✗ FAILED"
                    print(f"{status}: {filename}")
            
            return results  # Return results for programmatic use
            
        except Exception as e:
            if not silent:
                print(f"\n✗ Error: {e}")
                print("Taking a screenshot for debugging...")
            screenshot_path = os.path.join(script_dir, "..", "debug_screenshot.png")
            page.screenshot(path=screenshot_path)
            if not silent:
                print(f"Screenshot saved to: {screenshot_path}")
            return {f: False for f in FILES_TO_DOWNLOAD}
        
        finally:
            if not silent and not auto_close:
                input("\nPress Enter to close the browser...")
            context.close()

if __name__ == "__main__":
    print("="*70)
    print("SharePoint Files Downloader")
    print("="*70)
    print(f"Files to download: {', '.join(FILES_TO_DOWNLOAD)}")
    print(f"Source profile location: {source_edge_profile_path}")
    print(f"Automation profile location: {automation_profile_path}")
    print()
    download_sharepoint_files()
    print("\nDone!")
