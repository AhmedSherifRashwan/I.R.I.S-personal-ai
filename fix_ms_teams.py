import os
import glob
import winreg
import sqlite3
from pathlib import Path

def comprehensive_teams_search():
    """Comprehensive search for Microsoft Teams anywhere on the system"""
    print("🔍 Comprehensive Microsoft Teams Search")
    print("=" * 60)
    
    found_teams = []
    
    # 1. Search entire AppData directory
    print("📁 Searching AppData directory...")
    appdata_paths = [
        os.path.expanduser("~\\AppData\\Local"),
        os.path.expanduser("~\\AppData\\Roaming"),
        os.path.expanduser("~\\AppData\\LocalLow")
    ]
    
    for appdata_path in appdata_paths:
        if os.path.exists(appdata_path):
            print(f"   Scanning: {appdata_path}")
            # Look for Teams executables
            teams_patterns = [
                "**/Teams.exe",
                "**/teams.exe", 
                "**/ms-teams*.exe",
                "**/Microsoft Teams*.exe",
                "**/msteams*.exe"
            ]
            
            for pattern in teams_patterns:
                try:
                    matches = glob.glob(os.path.join(appdata_path, pattern), recursive=True)
                    for match in matches:
                        if match not in found_teams:
                            print(f"   ✅ Found: {match}")
                            found_teams.append(match)
                except Exception as e:
                    continue
    
    # 2. Search Program Files
    print("\n📁 Searching Program Files...")
    program_dirs = [
        "C:\\Program Files",
        "C:\\Program Files (x86)"
    ]
    
    for prog_dir in program_dirs:
        if os.path.exists(prog_dir):
            print(f"   Scanning: {prog_dir}")
            try:
                teams_patterns = [
                    "**/Teams.exe",
                    "**/Microsoft Teams*.exe",
                    "**/teams.exe"
                ]
                
                for pattern in teams_patterns:
                    matches = glob.glob(os.path.join(prog_dir, pattern), recursive=True)
                    for match in matches[:5]:  # Limit to first 5 matches to avoid too many results
                        if match not in found_teams:
                            print(f"   ✅ Found: {match}")
                            found_teams.append(match)
            except Exception as e:
                print(f"   ⚠️ Error scanning {prog_dir}: {e}")
    
    # 3. Check Windows Store apps (WindowsApps)
    print("\n📱 Searching Windows Store apps...")
    windowsapps_paths = [
        "C:\\Program Files\\WindowsApps",
        os.path.expanduser("~\\AppData\\Local\\Packages")
    ]
    
    for wa_path in windowsapps_paths:
        if os.path.exists(wa_path):
            print(f"   Scanning: {wa_path}")
            try:
                for item in os.listdir(wa_path):
                    if "teams" in item.lower() or "microsoft" in item.lower():
                        item_path = os.path.join(wa_path, item)
                        if os.path.isdir(item_path):
                            print(f"   📁 Found Teams-related folder: {item}")
                            # Look for executables in this folder
                            exe_files = glob.glob(os.path.join(item_path, "**/*.exe"), recursive=True)
                            for exe in exe_files:
                                if "teams" in os.path.basename(exe).lower():
                                    print(f"   ✅ Found Teams exe: {exe}")
                                    if exe not in found_teams:
                                        found_teams.append(exe)
            except PermissionError:
                print(f"   ⚠️ Permission denied accessing {wa_path}")
            except Exception as e:
                print(f"   ⚠️ Error scanning {wa_path}: {e}")
    
    # 4. Check Windows Registry
    print("\n💾 Checking Windows Registry for Teams...")
    registry_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    
    for reg_path in registry_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, subkey_name) as subkey:
                            try:
                                display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                if "teams" in display_name.lower() and "microsoft" in display_name.lower():
                                    print(f"   📝 Found in registry: {display_name}")
                                    try:
                                        install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                        if install_location:
                                            print(f"      📍 Install location: {install_location}")
                                            # Look for Teams exe in install location
                                            teams_exe = os.path.join(install_location, "Teams.exe")
                                            if os.path.exists(teams_exe) and teams_exe not in found_teams:
                                                found_teams.append(teams_exe)
                                    except FileNotFoundError:
                                        pass
                            except FileNotFoundError:
                                pass
                        i += 1
                    except OSError:
                        break
        except Exception as e:
            print(f"   ⚠️ Error checking registry: {e}")
    
    # 5. Check Start Menu
    print("\n🗂️ Checking Start Menu for Teams shortcuts...")
    start_menu_paths = [
        os.path.expanduser(r"~\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"),
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    ]
    
    for start_path in start_menu_paths:
        if os.path.exists(start_path):
            print(f"   Scanning: {start_path}")
            for root, dirs, files in os.walk(start_path):
                for file in files:
                    if ("teams" in file.lower() and 
                        ("microsoft" in file.lower() or file.lower().endswith('.lnk'))):
                        shortcut_path = os.path.join(root, file)
                        print(f"   🔗 Found shortcut: {file}")
                        print(f"      📍 Path: {shortcut_path}")
                        
                        # Try to resolve the shortcut
                        target = resolve_shortcut(shortcut_path)
                        if target and target not in found_teams:
                            print(f"      🎯 Target: {target}")
                            found_teams.append(target)
    
    # 6. Check Desktop
    print("\n🖥️ Checking Desktop for Teams shortcuts...")
    desktop_path = os.path.expanduser("~/Desktop")
    if os.path.exists(desktop_path):
        for file in os.listdir(desktop_path):
            if "teams" in file.lower() and ("microsoft" in file.lower() or file.lower().endswith('.lnk')):
                shortcut_path = os.path.join(desktop_path, file)
                print(f"   🔗 Found on desktop: {file}")
                target = resolve_shortcut(shortcut_path)
                if target and target not in found_teams:
                    print(f"      🎯 Target: {target}")
                    found_teams.append(target)
    
    return found_teams

def resolve_shortcut(shortcut_path):
    """Resolve Windows shortcut (.lnk) to actual target"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        return shortcut.Targetpath
    except Exception:
        return None

def search_by_process():
    """Search for Teams by checking running processes"""
    print("\n🔄 Checking running processes for Teams...")
    try:
        import psutil
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                if proc.info['name'] and 'teams' in proc.info['name'].lower():
                    print(f"   🏃 Running Teams process: {proc.info['name']}")
                    if proc.info['exe']:
                        print(f"      📍 Executable: {proc.info['exe']}")
                        return proc.info['exe']
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    except ImportError:
        print("   ⚠️ psutil not available, skipping process check")
        print("   💡 Install psutil with: pip install psutil")
    
    return None

def manual_teams_setup():
    """Allow user to manually specify Teams location"""
    print("\n" + "=" * 60)
    print("🔧 Manual Teams Setup")
    print("=" * 60)
    
    print("If you know where Microsoft Teams is installed, you can:")
    print("1. Right-click on Teams icon/shortcut")
    print("2. Select 'Open file location' or 'Properties'")
    print("3. Copy the path to the Teams.exe file")
    print()
    
    manual_path = input("Enter the full path to Teams.exe (or press Enter to skip): ").strip()
    
    if manual_path and os.path.exists(manual_path):
        print(f"✅ Valid path provided: {manual_path}")
        return [manual_path]
    elif manual_path:
        print(f"❌ Path not found: {manual_path}")
    
    return []

def add_teams_from_web_version():
    """Add Teams web version as an alternative"""
    print("\n🌐 Would you like to add Teams web version as an alternative?")
    print("This will open Teams in your default browser.")
    
    choice = input("Add Teams web version? (y/n): ").lower()
    if choice == 'y':
        # Add a web version launcher
        web_launcher_script = '''
import webbrowser
import sys

def open_teams_web():
    """Open Microsoft Teams in web browser"""
    teams_url = "https://teams.microsoft.com"
    webbrowser.open(teams_url)
    print("Opening Microsoft Teams in web browser...")

if __name__ == "__main__":
    open_teams_web()
'''
        
        script_path = "teams_web_launcher.py"
        with open(script_path, "w") as f:
            f.write(web_launcher_script)
        
        print(f"✅ Created Teams web launcher: {script_path}")
        
        # Add to database
        db_path = "iris_voice_learning.db"
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        python_path = os.sys.executable
        launcher_command = f'"{python_path}" "{os.path.abspath(script_path)}"'
        
        cursor.execute('''
            INSERT OR REPLACE INTO detected_programs (program_name, program_path, category, platform)
            VALUES (?, ?, ?, ?)
        ''', ("Microsoft Teams (Web)", launcher_command, "Application", "Web Browser"))
        
        # Add aliases
        teams_aliases = ["teams", "microsoft teams", "teams web", "ms teams"]
        for alias in teams_aliases:
            cursor.execute('''
                INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
                VALUES (?, ?, ?)
            ''', ("Microsoft Teams (Web)", alias, 1.5))
        
        conn.commit()
        conn.close()
        
        print("✅ Added Microsoft Teams web version to database!")
        return True
    
    return False

if __name__ == "__main__":
    print("🔍 Comprehensive Microsoft Teams Search")
    print("=" * 60)
    
    # First, try comprehensive search
    teams_installations = comprehensive_teams_search()
    
    # Check running processes
    running_teams = search_by_process()
    if running_teams and running_teams not in teams_installations:
        teams_installations.append(running_teams)
    
    if teams_installations:
        print(f"\n✅ Found {len(teams_installations)} Teams installations:")
        for i, path in enumerate(teams_installations, 1):
            print(f"   {i}. {path}")
        
        # Add the first one to database
        from fix_teams_detection import add_teams_to_database
        add_teams_to_database(teams_installations[:1])
    else:
        print("\n❌ No Microsoft Teams installations found through automatic search")
        
        # Try manual setup
        manual_teams = manual_teams_setup()
        if manual_teams:
            from fix_teams_detection import add_teams_to_database
            add_teams_to_database(manual_teams)
        else:
            # Offer web version
            if add_teams_from_web_version():
                print("\n✅ Teams web version added as alternative!")
            else:
                print("\n💡 Suggestions:")
                print("   1. Install Microsoft Teams from https://teams.microsoft.com")
                print("   2. Check if Teams is installed as a Windows Store app")
                print("   3. Look for Teams in your browser bookmarks")
                print("   4. Try the Teams web version in your browser")