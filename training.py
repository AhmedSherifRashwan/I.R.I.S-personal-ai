import os
import json
import sqlite3
import winreg
import glob
from pathlib import Path
from fuzzywuzzy import fuzz
import subprocess

DB_PATH = "iris_voice_learning.db"
TRAINING_FILE = "manual_training.json"

class VoiceLearning:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        self.init_voice_db()
        
    def init_voice_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS voice_patterns (
                id INTEGER PRIMARY KEY,
                spoken_command TEXT NOT NULL,
                intended_program TEXT NOT NULL,
                confidence_score REAL,
                success_count INTEGER DEFAULT 0,
                failure_count INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_used TIMESTAMP
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS program_aliases (
                id INTEGER PRIMARY KEY,
                program_name TEXT NOT NULL,
                alias TEXT NOT NULL,
                weight REAL DEFAULT 1.0,
                UNIQUE(program_name, alias)
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS detected_programs (
                id INTEGER PRIMARY KEY,
                program_name TEXT NOT NULL,
                program_path TEXT NOT NULL,
                category TEXT NOT NULL,
                platform TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(program_name, program_path)
            )
        ''')
        conn.commit()
        conn.close()

    def learn_command(self, spoken_text, program_name, success=True):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO voice_patterns 
            (spoken_command, intended_program, confidence_score, success_count, failure_count, last_used)
            VALUES (?, ?, ?, 
                COALESCE((SELECT success_count FROM voice_patterns WHERE spoken_command = ? AND intended_program = ?), 0) + ?,
                COALESCE((SELECT failure_count FROM voice_patterns WHERE spoken_command = ? AND intended_program = ?), 0) + ?,
                datetime('now'))
        ''', (
            spoken_text.lower(), program_name, 1.0 if success else 0.5,
            spoken_text.lower(), program_name, 1 if success else 0,
            spoken_text.lower(), program_name, 0 if success else 1
        ))
        conn.commit()
        conn.close()

    def add_program_alias(self, program_name, alias, weight=1.0):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
            VALUES (?, ?, ?)
        ''', (program_name, alias.lower(), weight))
        conn.commit()
        conn.close()

    def add_detected_program(self, program_name, program_path, category, platform=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO detected_programs (program_name, program_path, category, platform)
            VALUES (?, ?, ?, ?)
        ''', (program_name, program_path, category, platform))
        conn.commit()
        conn.close()

class ProgramDetector:
    def __init__(self, vl: VoiceLearning):
        self.vl = vl
        self.common_locations = [
            os.path.expanduser("~/Desktop"),
            os.path.expanduser("~/Documents"),
            os.path.expanduser("~/Downloads"),
            "C:\\Users\\Public\\Desktop"
        ]
        
    def detect_xbox_gamepass_games(self):
        """Detect Xbox Game Pass games"""
        games = []
        try:
            # Xbox Game Pass games are typically installed in WindowsApps
            xbox_paths = [
                "C:\\Program Files\\WindowsApps",
                "C:\\XboxGames"
            ]
            for xbox_path in xbox_paths:
                if not os.path.exists(xbox_path):
                    continue
                    # Get accessible subdirectories
                    try:
                        for item in os.listdir(xbox_path):
                        item_path = os.path.join(xbox_path, item)
                        if os.path.isdir(item_path):
                            # Look for .exe files in the directory
                            for exe_file in glob.glob(os.path.join(item_path, "*.exe")):
                                game_name = os.path.splitext(os.path.basename(exe_file))[0]
                                # Clean up the name
                                game_name = game_name.replace("_", " ").replace("-", " ")
                                games.append({
                                    'name': game_name,
                                    'path': exe_file,
                                    'category': 'Game',
                                    'platform': 'Xbox Game Pass'
                                })
                except PermissionError:
                    print("⚠️ Cannot access WindowsApps directory (permission required)")
            
            # Alternative: Check Xbox app install locations via registry
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                  r"Software\Microsoft\Windows\CurrentVersion\Uninstall") as key:
                    i = 0
                    while True:
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, subkey_name) as subkey:
                                try:
                                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                    install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                    
                                    # Check if it's an Xbox/Microsoft Store app
                                    if "Microsoft.GamingApp" in subkey_name or "XboxApp" in display_name:
                                        games.append({
                                            'name': display_name,
                                            'path': install_location,
                                            'category': 'Game',
                                            'platform': 'Xbox Game Pass'
                                        })
                                except FileNotFoundError:
                                    pass
                            i += 1
                        except OSError:
                            break
            except Exception as e:
                print(f"⚠️ Error checking Xbox registry: {e}")
                
        except Exception as e:
            print(f"⚠️ Error detecting Xbox Game Pass games: {e}")
        
        return games

    def detect_steam_games(self):
        """Detect Steam games"""
        games = []
        try:
            # Common Steam installation paths
            steam_paths = [
                "C:\\Program Files (x86)\\Steam",
                "C:\\Program Files\\Steam",
                os.path.expanduser("~/Steam")
            ]
            
            for steam_path in steam_paths:
                if os.path.exists(steam_path):
                    # Look for steamapps folder
                    steamapps_path = os.path.join(steam_path, "steamapps", "common")
                    if os.path.exists(steamapps_path):
                        for game_folder in os.listdir(steamapps_path):
                            game_path = os.path.join(steamapps_path, game_folder)
                            if os.path.isdir(game_path):
                                # Look for .exe files
                                exe_files = glob.glob(os.path.join(game_path, "*.exe"))
                                if exe_files:
                                    # Use folder name as game name
                                    game_name = game_folder.replace("_", " ").replace("-", " ")
                                    games.append({
                                        'name': game_name,
                                        'path': exe_files[0],  # Use first exe found
                                        'category': 'Game',
                                        'platform': 'Steam'
                                    })
                    break  # Found Steam, no need to check other paths
                    
        except Exception as e:
            print(f"⚠️ Error detecting Steam games: {e}")
        
        return games

    def detect_installed_applications(self):
        """Detect installed applications from registry"""
        applications = []
        try:
            # Check both 32-bit and 64-bit application registries
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
                                        install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                        
                                        # Filter out system components and updates
                                        if (display_name and install_location and 
                                            "Update" not in display_name and 
                                            "Hotfix" not in display_name and
                                            len(display_name) > 3):
                                            
                                            applications.append({
                                                'name': display_name,
                                                'path': install_location,
                                                'category': 'Application',
                                                'platform': 'Windows'
                                            })
                                    except FileNotFoundError:
                                        pass
                                i += 1
                            except OSError:
                                break
                except Exception as e:
                    print(f"⚠️ Error accessing registry path {reg_path}: {e}")
                    
        except Exception as e:
            print(f"⚠️ Error detecting applications: {e}")
        
        return applications

    def detect_office_files(self):
        """Detect Office files (Excel, Word, PowerPoint, etc.)"""
        files = []
        office_extensions = {
            '.xlsx': 'Excel File',
            '.xls': 'Excel File', 
            '.xlsm': 'Excel File',
            '.docx': 'Word Document',
            '.doc': 'Word Document',
            '.pptx': 'PowerPoint Presentation',
            '.ppt': 'PowerPoint Presentation',
            '.pdf': 'PDF Document',
            '.txt': 'Text File'
        }
        
        # Search in common locations
        search_locations = [
            os.path.expanduser("~/Documents"),
            os.path.expanduser("~/Desktop"),
            os.path.expanduser("~/Downloads")
        ]
        
        for location in search_locations:
            if os.path.exists(location):
                for ext, file_type in office_extensions.items():
                    pattern = os.path.join(location, f"**/*{ext}")
                    for file_path in glob.glob(pattern, recursive=True):
                        if os.path.isfile(file_path):
                            file_name = os.path.splitext(os.path.basename(file_path))[0]
                            files.append({
                                'name': f"{file_name} ({file_type})",
                                'path': file_path,
                                'category': 'Document',
                                'platform': file_type
                            })
        
        return files

    def detect_all_programs(self):
        """Detect all programs and files"""
        print("🔍 Starting program detection...")
        
        all_programs = []
        
        # Detect Xbox Game Pass games
        print("📱 Detecting Xbox Game Pass games...")
        xbox_games = self.detect_xbox_gamepass_games()
        all_programs.extend(xbox_games)
        print(f"✅ Found {len(xbox_games)} Xbox Game Pass games")
        
        # Detect Steam games
        print("🎮 Detecting Steam games...")
        steam_games = self.detect_steam_games()
        all_programs.extend(steam_games)
        print(f"✅ Found {len(steam_games)} Steam games")
        
        # Detect applications
        print("💻 Detecting installed applications...")
        applications = self.detect_installed_applications()
        all_programs.extend(applications)
        print(f"✅ Found {len(applications)} applications")
        
        # Detect Office files
        print("📄 Detecting Office files...")
        office_files = self.detect_office_files()
        all_programs.extend(office_files)
        print(f"✅ Found {len(office_files)} Office files")
        
        # Store in database
        print("💾 Storing detected programs in database...")
        for program in all_programs:
            self.vl.add_detected_program(
                program['name'],
                program['path'],
                program['category'],
                program.get('platform')
            )
            
            # Create voice aliases for the program
            self.create_voice_aliases(program['name'])
        
        print(f"✅ Detection complete! Found {len(all_programs)} total items")
        return all_programs

    def create_voice_aliases(self, program_name):
        """Create voice-friendly aliases for a program"""
        # Clean the name
        clean_name = program_name.lower()
        
        # Remove common words that might cause confusion
        remove_words = ['microsoft', 'the', 'version', '2019', '2020', '2021', '2022', '2023', '2024']
        for word in remove_words:
            clean_name = clean_name.replace(word, '').strip()
        
        # Create aliases
        aliases = [clean_name]
        
        # Add shortened versions
        words = clean_name.split()
        if len(words) > 1:
            # First word only
            aliases.append(words[0])
            # Acronym
            acronym = ''.join([word[0] for word in words if word])
            if len(acronym) > 1:
                aliases.append(acronym)
        
        # Add the aliases to the database
        for alias in set(aliases):  # Remove duplicates
            if alias and len(alias) > 1:
                self.vl.add_program_alias(program_name, alias, weight=1.5)

def train_manual(vl: VoiceLearning, spoken_phrase: str, program_name: str, alias_weight: float = 1.5):
    vl.learn_command(spoken_phrase, program_name, success=True)
    vl.add_program_alias(program_name, spoken_phrase, weight=alias_weight)
    print(f"✅ Learned: '{spoken_phrase}' → {program_name}")

def train_bulk(vl: VoiceLearning, mappings: dict):
    for program_name, phrases in mappings.items():
        for phrase in phrases:
            train_manual(vl, phrase, program_name)

def load_training_from_json(file_path=TRAINING_FILE):
    if not os.path.exists(file_path):
        print(f"No training file found at {file_path}. Skipping.")
        return
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            mappings = json.load(f)
        if isinstance(mappings, dict):
            vl = VoiceLearning()
            train_bulk(vl, mappings)
            print(f"✅ Loaded {len(mappings)} program mappings from {file_path}")
        else:
            print("❌ Invalid JSON format. Expected dictionary.")
    except Exception as e:
        print(f"❌ Error loading training file: {e}")

def run_auto_detection():
    """Run the automatic program detection"""
    vl = VoiceLearning()
    detector = ProgramDetector(vl)
    detector.detect_all_programs()

if __name__ == "__main__":
    print("🚀 Iris Voice Learning - Enhanced Training")
    print("=" * 50)
    
    # Load manual training first
    load_training_from_json()
    
    # Run automatic detection
    print("\n🔍 Running automatic program detection...")
    run_auto_detection()
    
    print("\n✅ Training complete! Your programs are now ready for voice commands.")