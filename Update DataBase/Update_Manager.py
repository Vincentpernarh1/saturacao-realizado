"""
Update Manager - Checks file ages and triggers SharePoint downloads when needed
"""
import os
import glob
import re
from datetime import datetime, timedelta
import sys
import argparse
    

# PyInstaller-compatible path resolution
if getattr(sys, 'frozen', False):
    # Running from PyInstaller .exe
    script_dir = sys._MEIPASS
else:
    # Running from source
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Add current directory to path for imports
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

def get_file_age_days(filepath):
    """Get the age of a file in days based on the date in the filename
    
    Expects filename format: FILENAME_YYYY-MM-DD.extension
    Falls back to file modification time if date cannot be parsed from filename.
    """
    if not os.path.exists(filepath):
        return None
    
    # Try to extract date from filename (format: FILENAME_YYYY-MM-DD.extension)
    filename = os.path.basename(filepath)
    date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
    
    if date_match:
        try:
            # Parse date from filename
            year, month, day = map(int, date_match.groups())
            file_date = datetime(year, month, day)
            age = datetime.now() - file_date
            return age.days
        except (ValueError, TypeError):
            pass  # Fall back to modification time
    
    # Fallback: use file modification time
    modified_time = os.path.getmtime(filepath)
    modified_date = datetime.fromtimestamp(modified_time)
    age = datetime.now() - modified_date
    return age.days

def get_latest_file(pattern):
    """Get the most recent file matching the pattern based on date in filename
    
    Expects filename format: FILENAME_YYYY-MM-DD.extension
    Falls back to file modification time if date cannot be parsed.
    """
    files = glob.glob(pattern)
    if not files:
        return None
    
    def extract_date_from_filename(filepath):
        """Extract date from filename for sorting"""
        filename = os.path.basename(filepath)
        date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
        if date_match:
            try:
                year, month, day = map(int, date_match.groups())
                return datetime(year, month, day)
            except (ValueError, TypeError):
                pass
        # Fallback to modification time
        return datetime.fromtimestamp(os.path.getmtime(filepath))
    
    # Sort by date extracted from filename, newest first
    files.sort(key=extract_date_from_filename, reverse=True)
    return files[0]

def needs_update(file_pattern, max_age_days=5):
    """Check if files matching the pattern need updating
    
    Args:
        file_pattern: Glob pattern to match files (e.g., "BD/BD_CADASTRO_PN_*.xlsx")
        max_age_days: Maximum age in days before update is needed
        
    Returns:
        tuple: (needs_update: bool, reason: str, latest_file: str or None)
    """
    latest_file = get_latest_file(file_pattern)
    
    if latest_file is None:
        return True, "File not found", None
    
    age_days = get_file_age_days(latest_file)
    
    if age_days is None:
        return True, "Cannot determine file age", latest_file
    
    if age_days > max_age_days:
        return True, f"File is {age_days} days old (max: {max_age_days})", latest_file
    
    return False, f"File is current ({age_days} days old)", latest_file

def check_and_update_files(max_age_days=5, force_update=False, silent=False, progress_callback=None):
    """Check file ages and update if needed
    
    Args:
        max_age_days: Maximum file age in days before triggering update
        force_update: If True, update regardless of file age
        silent: If True, suppress console output (for automated runs)
        progress_callback: Optional function(message) to report progress to GUI
        
    Returns:
        dict: Results of the update process
    """
    # Get the BD folder path (different for PyInstaller .exe vs source)
    if getattr(sys, 'frozen', False):
        # Running from .exe - BD folder is alongside the executable
        bd_folder = os.path.join(os.path.dirname(sys.executable), "BD")
    else:
        # Running from source - BD folder is in parent directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(script_dir)
        bd_folder = os.path.join(project_root, "BD")
    
    files_to_check = {
        "BD_CADASTRO_PN": os.path.join(bd_folder, "BD_CADASTRO_PN_*.xlsx"),
        "BD_CADASTRO_MDR": os.path.join(bd_folder, "BD_CADASTRO_MDR_*.xlsx")
    }
    
    update_needed = False
    reasons = []
    
    if not silent:
        print("="*70)
        print("Database Update Check")
        print("="*70)
    
    # Check each file
    for file_name, pattern in files_to_check.items():
        needs, reason, latest = needs_update(pattern, max_age_days)
        
        if not silent:
            status = "⚠️ UPDATE NEEDED" if needs else "✓ CURRENT"
            print(f"{status}: {file_name}")
            print(f"  Reason: {reason}")
            if latest:
                print(f"  File: {os.path.basename(latest)}")
        
        if progress_callback:
            if needs:
                progress_callback(f"⚠️ {file_name}: {reason}")
            else:
                progress_callback(f"✓ {file_name} atualizado")
        
        if needs or force_update:
            update_needed = True
            reasons.append(f"{file_name}: {reason}")
    
    # Perform update if needed
    if update_needed or force_update:
        if not silent:
            print("\n" + "="*70)
            print("Downloading latest files from SharePoint...")
            print("="*70)
        
        if progress_callback:
            progress_callback("Baixando arquivos do SharePoint...")
        
        try:
            # Import and run the download function
            from Update_Navigation import download_sharepoint_files
            
            # For now, always run with browser visible to avoid profile conflicts
            # TODO: Implement proper headless mode detection after profile initialization
            results = download_sharepoint_files(
                headless=False,  # Force headless=False to avoid profile issues
                silent=silent,
                auto_close=True,  # Auto-close when called from application
                progress_callback=progress_callback
            )
            
            if not silent:
                print("\n" + "="*70)
                print("Update Complete")
                print("="*70)
                all_success = all(results.values())
                if all_success:
                    print("✓ All files updated successfully")
                else:
                    print("⚠️ Some files failed to update")
                    for filename, success in results.items():
                        if not success:
                            print(f"  ✗ {filename}")
            
            return {
                "updated": True,
                "success": all(results.values()),
                "results": results,
                "reasons": reasons
            }
            
        except Exception as e:
            if not silent:
                print(f"\n✗ Error during update: {e}")
            return {
                "updated": False,
                "success": False,
                "error": str(e),
                "reasons": reasons
            }
    else:
        if not silent:
            print("\n✓ All files are up to date. No download needed.")
        return {
            "updated": False,
            "success": True,
            "message": "Files are current"
        }

if __name__ == "__main__":
    # When run directly, perform update check
   
    parser = argparse.ArgumentParser(description="Check and update database files from SharePoint")
    parser.add_argument("--max-age", type=int, default=5, help="Maximum file age in days (default: 5)")
    parser.add_argument("--force", action="store_true", help="Force update regardless of file age")
    parser.add_argument("--silent", action="store_true", help="Suppress console output")
    
    args = parser.parse_args()
    
    result = check_and_update_files(
        max_age_days=args.max_age,
        force_update=args.force,
        silent=args.silent
    )
    
    # Exit with appropriate code
    sys.exit(0 if result["success"] else 1)
