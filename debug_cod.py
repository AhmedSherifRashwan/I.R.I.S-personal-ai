import os
import sqlite3
import winreg
import glob
from pathlib import Path

def debug_call_of_duty_detection():
    """Debug script to find Call of Duty Black Ops 6"""
    print("🔍 Debugging Call of Duty Black Ops 6 Detection")
    print("=" * 60)
    
    # Check common Xbox Game Pass locations
    print("📱 Checking Xbox Game Pass locations...")
    xbox_paths = [
        "C:\\Program Files\\WindowsApps",
        "C:\\XboxGames",
        os.path.expanduser("~\\AppData\\Local\\Packages")
    ]
    
    for xbox_path in xbox_paths:
        if os.path.exists(xbox_path):
            print(f"✅ Found directory: {xbox_path}")
            try:
                for item in os.listdir(xbox_path):
                    if "call" in item.lower() or "cod" in item.lower() or "activision" in item.lower():
                        item_path = os.path.join(xbox_path, item)
                        print(f"🎮 Potential COD folder: {item}")
                        if os.path.isdir(item_path):
                            # Look for exe files
                            exe_files = glob.glob(os.path.join(item_path, "**/*.exe"), recursive=True)
                            for exe in exe_files:
                                print(f"   📁 Found exe: {exe}")
            except PermissionError:
                print(f"⚠️ Permission denied accessing {xbox_path}")
            except Exception as e:
                print(f"❌ Error scanning {xbox_path}: {e}")
        else:
            print(f"❌ Directory not found: {xbox_path}")
    
    print("\n🎮 Checking Steam locations...")
    steam_paths = [
        "C:\\Program Files (x86)\\Steam\\steamapps\\common",
        "C:\\Program Files\\Steam\\steamapps\\common",
        os.path.expanduser("~\\Steam\\steamapps\\common")
    ]
    
    for steam_path in steam_paths:
        if os.path.exists(steam_path):
            print(f"✅ Found Steam directory: {steam_path}")
            for game_folder in os.listdir(steam_path):
                if "call" in game_folder.lower() or "cod" in game_folder.lower():
                    game_path = os.path.join(steam_path, game_folder)
                    print(f"🎮 Found COD game folder: {game_folder}")
                    exe_files = glob.glob(os.path.join(game_path, "*.exe"))
                    for exe in exe_files:
                        print(f"   📁 Found exe: {exe}")
        else:
            print(f"❌ Steam directory not found: {steam_path}")
    
    print("\n💾 Checking Windows Registry...")
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
                                if ("call of duty" in display_name.lower() or 
                                    "cod" in display_name.lower() or 
                                    "black ops" in display_name.lower()):
                                    print(f"🎮 Found in registry: {display_name}")
                                    try:
                                        install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                        print(f"   📍 Install location: {install_location}")
                                    except FileNotFoundError:
                                        print(f"   ⚠️ No install location found")
                            except FileNotFoundError:
                                pass
                        i += 1
                    except OSError:
                        break
        except Exception as e:
            print(f"❌ Error checking registry: {e}")
    
    print("\n🗂️ Checking Start Menu...")
    start_menu_paths = [
        os.path.expanduser(r"~\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"),
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    ]
    
    for start_path in start_menu_paths:
        if os.path.exists(start_path):
            for root, dirs, files in os.walk(start_path):
                for file in files:
                    if ("call of duty" in file.lower() or 
                        "cod" in file.lower() or 
                        "black ops" in file.lower()):
                        shortcut_path = os.path.join(root, file)
                        print(f"🎮 Found shortcut: {file}")
                        print(f"   📍 Path: {shortcut_path}")
    
    print("\n💻 Checking Desktop...")
    desktop_path = os.path.expanduser("~/Desktop")
    if os.path.exists(desktop_path):
        for file in os.listdir(desktop_path):
            if ("call of duty" in file.lower() or 
                "cod" in file.lower() or 
                "black ops" in file.lower()):
                print(f"🎮 Found on desktop: {file}")

def check_database_content():
    """Check what's actually in the database"""
    print("\n" + "=" * 60)
    print("📊 Checking Database Content")
    print("=" * 60)
    
    db_path = "iris_voice_learning.db"
    if not os.path.exists(db_path):
        print("❌ Database not found!")
        return
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Check detected programs
    cursor.execute("SELECT program_name, program_path, category, platform FROM detected_programs WHERE program_name LIKE '%call%' OR program_name LIKE '%cod%' OR program_name LIKE '%black ops%'")
    cod_programs = cursor.fetchall()
    
    if cod_programs:
        print("🎮 Found COD-related programs in database:")
        for name, path, category, platform in cod_programs:
            print(f"   • {name} ({category}, {platform})")
            print(f"     Path: {path}")
    else:
        print("❌ No COD-related programs found in database")
    
    # Check aliases
    cursor.execute("SELECT program_name, alias FROM program_aliases WHERE program_name LIKE '%call%' OR program_name LIKE '%cod%' OR program_name LIKE '%black ops%'")
    cod_aliases = cursor.fetchall()
    
    if cod_aliases:
        print("\n🗣️ Found COD-related aliases:")
        for program, alias in cod_aliases:
            print(f"   • '{alias}' → {program}")
    else:
        print("\n❌ No COD-related aliases found")
    
    conn.close()

def manual_add_cod():
    """Manually add Call of Duty if found"""
    print("\n" + "=" * 60)
    print("➕ Manual COD Addition")
    print("=" * 60)
    
    # Try to find Battle.net launcher (common way to launch COD)
    battlenet_paths = [
        r"C:\Program Files (x86)\Battle.net\Battle.net Launcher.exe",
        r"C:\Program Files\Battle.net\Battle.net Launcher.exe"
    ]
    
    for path in battlenet_paths:
        if os.path.exists(path):
            print(f"✅ Found Battle.net: {path}")
            
            # Add to database
            db_path = "iris_voice_learning.db"
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Add the program
            cursor.execute('''
                INSERT OR REPLACE INTO detected_programs (program_name, program_path, category, platform)
                VALUES (?, ?, ?, ?)
            ''', ("Call of Duty Black Ops 6", path, "Game", "Battle.net"))
            
            # Add aliases
            aliases = ["call of duty", "cod", "black ops", "bo6", "cod black ops"]
            for alias in aliases:
                cursor.execute('''
                    INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
                    VALUES (?, ?, ?)
                ''', ("Call of Duty Black Ops 6", alias, 1.5))
            
            conn.commit()
            conn.close()
            
            print("✅ Added Call of Duty Black Ops 6 to database with Battle.net launcher")
            print("🗣️ Added aliases: call of duty, cod, black ops, bo6, cod black ops")
            return True
    
    print("❌ Battle.net launcher not found")
    
    # Ask user to manually specify path
    print("\n🤔 If you know where Call of Duty is installed, please:")
    print("1. Find the game's .exe file")
    print("2. Run this script again and manually add it")
    
    return False

if __name__ == "__main__":
    debug_call_of_duty_detection()
    check_database_content()
    
    print("\n" + "=" * 60)
    choice = input("Would you like to try manually adding Battle.net launcher for COD? (y/n): ")
    if choice.lower() == 'y':
        if manual_add_cod():
            print("\n✅ Manual addition complete! Try running your main.py again.")
        else:
            print("\n❌ Manual addition failed.")