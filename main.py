# Enhanced IRIS AI Assistant with RAG System Discovery
import time
import pvporcupine
import pyaudio
import struct
import whisper
import pyttsx3
import os
import glob
import ollama
import json
import wave
import numpy as np
import random
import edge_tts
import asyncio
import pygame
import io
import sqlite3
import winreg
import subprocess
from pathlib import Path
import threading
from fuzzywuzzy import fuzz
import pickle

# === Setup === 
ACCESS_KEY = "5cfCv3qXiQfV5vUl0HkaS1fcp22yIju+lljpqDOn0vtSD4U7PP3VWQ=="

print("Initializing IRIS with RAG capabilities...")
os.environ["PATH"] += os.pathsep + r"C:\ffmpeg\bin"

# Initialize pygame mixer for audio playback
pygame.mixer.init()

# === RAG System Discovery Class ===
class SystemDiscovery:
    def __init__(self, db_path="iris_system.db"):
        self.db_path = db_path
        self.init_database()
        self.program_cache = {}
        self.load_cache()
        
    def init_database(self):
        """Initialize SQLite database for system knowledge"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Programs table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS programs (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                display_name TEXT,
                path TEXT NOT NULL,
                keywords TEXT,
                category TEXT,
                last_accessed TIMESTAMP,
                access_count INTEGER DEFAULT 0,
                UNIQUE(name, path)
            )
        ''')
        
        # Files table for frequently accessed files
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                path TEXT NOT NULL,
                keywords TEXT,
                file_type TEXT,
                last_accessed TIMESTAMP,
                access_count INTEGER DEFAULT 0,
                UNIQUE(path)
            )
        ''')
        
        conn.commit()
        conn.close()
        
    def discover_programs(self):
        """Discover installed programs from multiple sources"""
        print("🔍 Discovering installed programs...")
        programs = []
        
        # Method 1: Windows Registry (Uninstall entries)
        programs.extend(self._get_programs_from_registry())
        
        # Method 2: Common installation directories
        programs.extend(self._scan_program_directories())
        
        # Method 3: Start Menu shortcuts
        programs.extend(self._scan_start_menu())
        
        # Method 4: Desktop shortcuts
        programs.extend(self._scan_desktop())
        
        # Store in database
        self._store_programs(programs)
        print(f"✅ Discovered {len(programs)} programs")
        
    def _get_programs_from_registry(self):
        """Get programs from Windows Registry"""
        programs = []
        registry_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        
        for reg_path in registry_paths:
            try:
                reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                for i in range(winreg.QueryInfoKey(reg_key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(reg_key, i)
                        subkey = winreg.OpenKey(reg_key, subkey_name)
                        
                        try:
                            display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            
                            # Look for executable in install location
                            if install_location and os.path.exists(install_location):
                                exe_files = glob.glob(os.path.join(install_location, "*.exe"))
                                for exe in exe_files:
                                    programs.append({
                                        'name': Path(exe).stem.lower(),
                                        'display_name': display_name,
                                        'path': exe,
                                        'category': 'application'
                                    })
                        except FileNotFoundError:
                            continue
                        finally:
                            winreg.CloseKey(subkey)
                    except (OSError, FileNotFoundError):
                        continue
                winreg.CloseKey(reg_key)
            except Exception as e:
                continue
                
        return programs
    
    def _scan_program_directories(self):
        """Scan common program installation directories"""
        programs = []
        common_dirs = [
            r"C:\Program Files",
            r"C:\Program Files (x86)",
            os.path.expanduser("~\\AppData\\Local"),
            os.path.expanduser("~\\AppData\\Roaming")
        ]
        
        for base_dir in common_dirs:
            if not os.path.exists(base_dir):
                continue
                
            try:
                # Look for .exe files in subdirectories (max 2 levels deep)
                for root, dirs, files in os.walk(base_dir):
                    # Limit depth to avoid scanning everything
                    level = root.replace(base_dir, '').count(os.sep)
                    if level < 3:
                        for file in files:
                            if file.lower().endswith('.exe'):
                                full_path = os.path.join(root, file)
                                programs.append({
                                    'name': Path(file).stem.lower(),
                                    'display_name': Path(file).stem,
                                    'path': full_path,
                                    'category': 'application'
                                })
            except (PermissionError, OSError):
                continue
                
        return programs
    
    def _scan_start_menu(self):
        """Scan Windows Start Menu shortcuts"""
        programs = []
        start_menu_paths = [
            os.path.expanduser(r"~\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"),
            r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
        ]
        
        for start_path in start_menu_paths:
            if not os.path.exists(start_path):
                continue
                
            for root, dirs, files in os.walk(start_path):
                for file in files:
                    if file.lower().endswith('.lnk'):
                        shortcut_path = os.path.join(root, file)
                        target = self._resolve_shortcut(shortcut_path)
                        if target:
                            # For Xbox Game Pass and protocol shortcuts, use the shortcut itself
                            if target.endswith('.lnk') or not target.lower().endswith('.exe'):
                                programs.append({
                                    'name': Path(file).stem.lower(),
                                    'display_name': Path(file).stem,
                                    'path': shortcut_path,  # Use shortcut path for special shortcuts
                                    'category': 'shortcut'
                                })
                            else:
                                programs.append({
                                    'name': Path(file).stem.lower(),
                                    'display_name': Path(file).stem,
                                    'path': target,
                                    'category': 'application'
                                })
        
        return programs
    
    def _scan_desktop(self):
        """Scan desktop shortcuts"""
        programs = []
        desktop_path = os.path.expanduser("~/Desktop")
        
        if os.path.exists(desktop_path):
            for file in os.listdir(desktop_path):
                if file.lower().endswith('.lnk'):
                    shortcut_path = os.path.join(desktop_path, file)
                    target = self._resolve_shortcut(shortcut_path)
                    if target:
                        # For Xbox Game Pass and protocol shortcuts, use the shortcut itself
                        if target.endswith('.lnk') or not target.lower().endswith('.exe'):
                            programs.append({
                                'name': Path(file).stem.lower(),
                                'display_name': Path(file).stem,
                                'path': shortcut_path,  # Use shortcut path for special shortcuts
                                'category': 'shortcut'
                            })
                        else:
                            programs.append({
                                'name': Path(file).stem.lower(),
                                'display_name': Path(file).stem,
                                'path': target,
                                'category': 'application'
                            })
        
        return programs
    
    def _resolve_shortcut(self, shortcut_path):
        """Resolve Windows shortcut to actual target"""
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            
            # For Xbox Game Pass and other special shortcuts
            target = shortcut.Targetpath
            arguments = shortcut.Arguments
            
            # If target is empty but arguments exist, it might be a special protocol
            if not target and arguments:
                # Store the shortcut path itself as the "executable" for special protocols
                return shortcut_path
            elif target:
                return target
            else:
                # Fallback to shortcut path for protocol-based shortcuts
                return shortcut_path
        except Exception as e:
            print(f"Error resolving shortcut {shortcut_path}: {e}")
            return None
    
    def _store_programs(self, programs):
        """Store discovered programs in database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        for program in programs:
            # Generate keywords from name and display name
            keywords = self._generate_keywords(program['name'], program.get('display_name', ''))
            
            cursor.execute('''
                INSERT OR REPLACE INTO programs 
                (name, display_name, path, keywords, category) 
                VALUES (?, ?, ?, ?, ?)
            ''', (
                program['name'],
                program.get('display_name'),
                program['path'],
                keywords,
                program.get('category', 'application')
            ))
        
        conn.commit()
        conn.close()
        self.load_cache()  # Refresh cache
    
    def _generate_keywords(self, name, display_name):
        """Generate search keywords for a program"""
        keywords = set()
        
        # Add name variations
        keywords.add(name.lower())
        if display_name:
            keywords.add(display_name.lower())
            
        # Add partial matches
        name_parts = name.lower().replace('-', ' ').replace('_', ' ').split()
        for part in name_parts:
            if len(part) > 2:  # Avoid very short parts
                keywords.add(part)
                
        if display_name:
            display_parts = display_name.lower().replace('-', ' ').replace('_', ' ').split()
            for part in display_parts:
                if len(part) > 2:
                    keywords.add(part)
        
        # Special handling for common game variations
        name_lower = name.lower()
        if 'call of duty' in name_lower or 'cod' in name_lower:
            keywords.update(['call', 'duty', 'cod', 'callofduty'])
            if 'black ops' in name_lower:
                keywords.update(['black', 'ops', 'blackops', 'bo'])
            if '6' in name_lower or 'six' in name_lower:
                keywords.update(['6', 'six'])
        
        if display_name and display_name.lower():
            display_lower = display_name.lower()
            if 'call of duty' in display_lower or 'cod' in display_lower:
                keywords.update(['call', 'duty', 'cod', 'callofduty'])
                if 'black ops' in display_lower:
                    keywords.update(['black', 'ops', 'blackops', 'bo'])
                if '6' in display_lower or 'six' in display_lower:
                    keywords.update(['6', 'six'])
        
        return ' '.join(keywords)
    
    def load_cache(self):
        """Load programs into memory cache for fast searching"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT name, display_name, path, keywords FROM programs')
        self.program_cache = {}
        
        for row in cursor.fetchall():
            name, display_name, path, keywords = row
            self.program_cache[name] = {
                'display_name': display_name,
                'path': path,
                'keywords': keywords.split() if keywords else []
            }
        
        conn.close()
        print(f"📚 Loaded {len(self.program_cache)} programs into cache")
    
    def find_program_with_aliases(self, query):
        """Enhanced program finding with alias support"""
        query_lower = query.lower()
        best_match = None
        best_score = 0
        
        for name, info in self.program_cache.items():
            score = 0
            
            # Direct name match
            if query_lower in name:
                score = 100
            # Display name match
            elif info['display_name'] and query_lower in info['display_name'].lower():
                score = 95
            else:
                # Check custom aliases
                aliases = voice_learning.get_program_aliases(name)
                for alias, weight in aliases:
                    alias_score = fuzz.ratio(query_lower, alias) * weight
                    score = max(score, alias_score)
                
                # Fuzzy matching on keywords
                for keyword in info['keywords']:
                    keyword_score = fuzz.ratio(query_lower, keyword)
                    score = max(score, keyword_score)
                
                # Partial ratio matching
                name_score = fuzz.partial_ratio(query_lower, name)
                if info['display_name']:
                    display_score = fuzz.partial_ratio(query_lower, info['display_name'].lower())
                    score = max(score, name_score, display_score)
                else:
                    score = max(score, name_score)
            
            if score > best_score and score > 60:  # Minimum threshold
                best_score = score
                best_match = {
                    'name': name,
                    'display_name': info['display_name'],
                    'path': info['path'],
                    'score': score
                }
        
        return best_match
    
    def launch_program(self, program_info):
        """Launch a program and update usage stats"""
        try:
            path = program_info['path']
            
            if path.endswith('.lnk'):
                # For shortcuts (including Xbox Game Pass), use os.startfile
                os.startfile(path)
                print(f"Launched shortcut: {path}")
            elif os.path.exists(path):
                # For regular executables
                os.startfile(path)
                print(f"Launched executable: {path}")
            else:
                print(f"Program path no longer exists: {path}")
                return False
                
            # Update access count
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE programs 
                SET access_count = access_count + 1, last_accessed = datetime('now')
                WHERE path = ?
            ''', (path,))
            conn.commit()
            conn.close()
            
            return True
        except Exception as e:
            print(f"Error launching program: {e}")
            return False

# Initialize system discovery
system_discovery = SystemDiscovery()

# Run initial discovery in background
def run_discovery():
    system_discovery.discover_programs()

discovery_thread = threading.Thread(target=run_discovery, daemon=True)
discovery_thread.start()

# === Original IRIS Code (with enhancements) ===

# Wake word engine
porcupine = pvporcupine.create(access_key=ACCESS_KEY, keyword_paths=["hey-Iris_en_windows_v3_0_0.ppn"])
pa = pyaudio.PyAudio()
stream = pa.open(rate=porcupine.sample_rate, channels=1, format=pyaudio.paInt16, input=True, frames_per_buffer=porcupine.frame_length)

# Speech-to-text model
print("Loading Whisper model...")
stt_model = whisper.load_model("base")

# TTS setup
engine = pyttsx3.init(driverName='sapi5')
voices = engine.getProperty("voices")
uk_voice_found = False
for v in voices:
    if v and hasattr(v, 'name') and v.name and "english" in v.name.lower() and "uk" in v.name.lower():
        engine.setProperty("voice", v.id)
        uk_voice_found = True
        break
if not uk_voice_found:
    print("UK voice not found, using default voice")
engine.setProperty("rate", 170)

USE_EDGE_TTS = True
import uuid

async def _speak_with_edge_tts(text):
    """Use Edge-TTS with proper audio playback"""
    try:
        communicate = edge_tts.Communicate(text, "en-GB-LibbyNeural")
        audio_data = b""
        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                audio_data += chunk["data"]
        
        if audio_data:
            temp_audio_path = f"temp_tts_{uuid.uuid4().hex[:8]}.mp3"
            with open(temp_audio_path, "wb") as f:
                f.write(audio_data)
            
            pygame.mixer.music.stop()
            time.sleep(0.05)
            pygame.mixer.music.load(temp_audio_path)
            pygame.mixer.music.play()
            
            while pygame.mixer.music.get_busy():
                time.sleep(0.05)
            time.sleep(0.1)
            
            try:
                os.remove(temp_audio_path)
            except:
                pass
            return True
    except Exception as e:
        print(f"Edge-TTS error: {e}")
        return False

def _speak_with_pyttsx3(text):
    """Fallback TTS using pyttsx3"""
    engine.say(text)
    engine.runAndWait()

def speak(text):
    """Main speak function with fallback support"""
    print(f"IRIS: {text}")
    if USE_EDGE_TTS:
        success = False
        try:
            success = asyncio.run(_speak_with_edge_tts(text))
        except Exception as e:
            print(f"Edge-TTS failed: {e}")
            success = False
        
        if not success:
            print("Falling back to pyttsx3...")
            _speak_with_pyttsx3(text)
    else:
        _speak_with_pyttsx3(text)

def get_acknowledgment():
    """Returns a random acknowledgment phrase"""
    acknowledgments = [
        "Right away, sir.",
        "Yes, sir.",
        "On it, sir.",
        "Certainly, sir.",
        "Immediately, sir.",
        "Of course, sir."
    ]
    return random.choice(acknowledgments)

def quick_acknowledge_and_execute(action_func, custom_message=None):
    """Immediately acknowledge the command then execute it"""
    if custom_message:
        speak(custom_message)
    else:
        speak(get_acknowledgment())
    time.sleep(0.1)
    action_func()

def listen_and_transcribe(duration=4):
    """Records audio and transcribes it with improved accuracy"""
    print("🎤 Recording command... Speak now!")
    frames = []
    for _ in range(int(porcupine.sample_rate / 1024 * duration)):
        data = stream.read(1024, exception_on_overflow=False)
        frames.append(data)

    temp_file_path = os.path.abspath("temp_command.wav")
    wf = wave.open(temp_file_path, 'wb')
    wf.setnchannels(1)
    wf.setsampwidth(pa.get_sample_size(pyaudio.paInt16))
    wf.setframerate(porcupine.sample_rate)
    wf.writeframes(b''.join(frames))
    wf.close()

    print("🔄 Transcribing...")
    
    # Enhanced transcription with multiple attempts for better accuracy
    transcription_attempts = [
        # Standard transcription
        {
            'temperature': 0.0,
            'initial_prompt': "Commands like open chrome, open browser, open excel, open spotify, open call of duty, open hogwarts legacy"
        },
        # Gaming focused transcription
        {
            'temperature': 0.2,
            'initial_prompt': "Gaming commands: open Hogwarts Legacy, Call of Duty Black Ops 6, Cyberpunk 2077, Grand Theft Auto, Minecraft, Among Us"
        },
        # General application transcription
        {
            'temperature': 0.1,
            'initial_prompt': "Application commands: launch program, start application, open software"
        }
    ]
    
    best_result = None
    best_confidence = 0
    
    for attempt in transcription_attempts:
        try:
            result = stt_model.transcribe(
                temp_file_path,
                language="en",
                temperature=attempt['temperature'],
                initial_prompt=attempt['initial_prompt']
            )
            
            # Simple confidence estimation based on result consistency
            text = result["text"].strip().lower()
            if text and len(text) > 2:
                # Prefer results with gaming/application keywords
                gaming_keywords = ['open', 'launch', 'start', 'hogwarts', 'legacy', 'call', 'duty', 'cod']
                keyword_bonus = sum(1 for word in gaming_keywords if word in text) * 10
                
                # Length bonus (reasonable length commands)
                length_bonus = min(len(text.split()) * 5, 20)
                
                confidence = keyword_bonus + length_bonus
                
                if confidence > best_confidence:
                    best_confidence = confidence
                    best_result = text
                    
        except Exception as e:
            print(f"Transcription attempt failed: {e}")
            continue
    
    # Fallback to first successful transcription if no best found
    if not best_result:
        try:
            result = stt_model.transcribe(temp_file_path, language="en", temperature=0.0)
            best_result = result["text"].strip().lower()
        except:
            best_result = ""
    
    print(f"📝 Heard: {best_result}")

    try:
        os.remove(temp_file_path)
    except:
        pass
    
    return best_result

# === Voice Learning and Custom Training ===
class VoiceLearning:
    def __init__(self, db_path="iris_voice_learning.db"):
        self.db_path = db_path
        self.init_voice_db()
        self.command_patterns = {}
        self.load_patterns()
        
    def init_voice_db(self):
        """Initialize database for voice learning"""
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
        
        conn.commit()
        conn.close()
        
    def load_patterns(self):
        """Load learned voice patterns"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT spoken_command, intended_program, confidence_score FROM voice_patterns')
        for row in cursor.fetchall():
            spoken, intended, confidence = row
            if intended not in self.command_patterns:
                self.command_patterns[intended] = []
            self.command_patterns[intended].append({
                'spoken': spoken.lower(),
                'confidence': confidence or 1.0
            })
        
        conn.close()
        
    def learn_command(self, spoken_text, program_name, success=True):
        """Learn from user interactions"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Update or insert pattern
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
        self.load_patterns()  # Refresh patterns
        
    def add_program_alias(self, program_name, alias, weight=1.0):
        """Add custom aliases for programs"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
            VALUES (?, ?, ?)
        ''', (program_name, alias.lower(), weight))
        
        conn.commit()
        conn.close()
        
    def get_program_aliases(self, program_name):
        """Get all aliases for a program"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT alias, weight FROM program_aliases WHERE program_name = ?', (program_name,))
        aliases = cursor.fetchall()
        conn.close()
        
        return aliases

# Initialize voice learning
voice_learning = VoiceLearning()

# Pre-populate some common game aliases
def setup_game_aliases():
    """Setup common game aliases and variations"""
    game_aliases = {
        'hogwarts legacy': ['hogwarts', 'harry potter', 'hogwarts legacy game', 'legacy'],
        'call of duty black ops 6': ['cod', 'call of duty', 'black ops', 'cod black ops', 'bo6', 'cod6'],
        'grand theft auto': ['gta', 'grand theft auto', 'gta 5', 'gta v'],
        'red dead redemption': ['red dead', 'rdr', 'red dead 2', 'rdr2'],
        'cyberpunk 2077': ['cyberpunk', 'cyber punk', '2077'],
        'the witcher 3': ['witcher', 'witcher 3', 'geralt game'],
        'minecraft': ['mine craft', 'block game'],
        'among us': ['among', 'sus game'],
        'fall guys': ['fall guy', 'falling guys']
    }
    
    for program, aliases in game_aliases.items():
        for alias in aliases:
            voice_learning.add_program_alias(program, alias, 1.5)

# Run setup
setup_game_aliases()

# === Enhanced Skills ===

def open_program_dynamically(query, learn_from_result=True):
    """Use RAG to find and open any program with voice learning"""
    print(f"🔍 Searching for program: {query}")
    
    # Wait for discovery to complete if still running
    if discovery_thread.is_alive():
        speak("Just a moment while I finish learning about your system, sir.")
        discovery_thread.join(timeout=10)  # Wait max 10 seconds
    
    # First check learned patterns
    best_learned_match = None
    best_learned_score = 0
    
    for program_name, patterns in voice_learning.command_patterns.items():
        for pattern in patterns:
            similarity = fuzz.ratio(query.lower(), pattern['spoken'])
            weighted_score = similarity * pattern['confidence']
            if weighted_score > best_learned_score and weighted_score > 70:
                best_learned_score = weighted_score
                best_learned_match = program_name
    
    # If we found a good learned match, try it first
    if best_learned_match:
        program = system_discovery.find_program(best_learned_match)
        if program:
            print(f"✅ Found via learning: {program['display_name']} (learned confidence: {best_learned_score:.1f}%)")
            if system_discovery.launch_program(program):
                if learn_from_result:
                    voice_learning.learn_command(query, best_learned_match, success=True)
                speak(f"Opening {program['display_name']}, sir.")
                return True
            else:
                if learn_from_result:
                    voice_learning.learn_command(query, best_learned_match, success=False)
    
    # Fallback to regular discovery with alias enhancement
    program = system_discovery.find_program_with_aliases(query)
    
    if program:
        print(f"✅ Found: {program['display_name']} (confidence: {program['score']}%)")
        if system_discovery.launch_program(program):
            if learn_from_result:
                voice_learning.learn_command(query, program['name'], success=True)
            speak(f"Opening {program['display_name']}, sir.")
            return True
        else:
            if learn_from_result:
                voice_learning.learn_command(query, program['name'], success=False)
            speak(f"I found {program['display_name']} but couldn't launch it, sir.")
            return False
    else:
        # Ask user for clarification and learn
        available_programs = list(system_discovery.program_cache.keys())[:10]  # Show top 10
        similar_programs = [p for p in available_programs if any(word in p for word in query.split())]
        
        if similar_programs:
            speak(f"I couldn't find '{query}', sir. Did you mean one of these programs: {', '.join(similar_programs[:3])}?")
        else:
            speak(f"I couldn't find a program matching '{query}', sir. Try being more specific or check if it's installed.")
        return False

def add_manual_training_commands():
    """Add manual training commands for voice learning"""
    def train_command():
        """Interactive training mode"""
        speak("Entering training mode, sir. Say a command you want me to learn.")
        time.sleep(0.5)
        
        spoken_command = listen_and_transcribe(duration=5)
        if not spoken_command:
            speak("I didn't hear anything, sir.")
            return
            
        speak("What program should this command open? Say the exact program name.")
        time.sleep(0.5)
        
        program_name = listen_and_transcribe(duration=5)
        if not program_name:
            speak("I didn't hear a program name, sir.")
            return
        
        # Find the program in our system
        program = system_discovery.find_program_with_aliases(program_name)
        if program:
            voice_learning.learn_command(spoken_command, program['name'], success=True)
            speak(f"I've learned that '{spoken_command}' should open {program['display_name']}, sir.")
            
            # Test the new learning
            speak("Let me test this. Please repeat the command.")
            time.sleep(0.5)
            test_command = listen_and_transcribe(duration=5)
            if test_command:
                if open_program_dynamically(test_command, learn_from_result=False):
                    speak("Perfect! The training was successful, sir.")
                else:
                    speak("The training needs more work, sir. Try again later.")
        else:
            speak(f"I couldn't find a program called '{program_name}', sir.")
    
    def show_learned_commands():
        """Show what commands have been learned"""
        conn = sqlite3.connect(voice_learning.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT spoken_command, intended_program, success_count 
            FROM voice_patterns 
            ORDER BY success_count DESC 
            LIMIT 10
        ''')
        
        learned = cursor.fetchall()
        conn.close()
        
        if learned:
            speak("Here are your most successful learned commands, sir:")
            for spoken, program, count in learned[:5]:  # Top 5
                speak(f"'{spoken}' opens {program}, used {count} times successfully.")
        else:
            speak("No learned commands yet, sir.")
    
    return train_command, show_learned_commands

# Create training functions
train_command, show_learned_commands = add_manual_training_commands()
print("🎤 Testing microphone for 5 seconds...")
pa_test = pyaudio.PyAudio()
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 16000
CHUNK = 1024
    
try:
        stream_test = pa_test.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=CHUNK)
        print("Speak now to test your microphone...")
        for i in range(int(RATE / CHUNK * 5)):
            data = stream_test.read(CHUNK, exception_on_overflow=False)
            audio_data = np.frombuffer(data, dtype=np.int16)
            audio_level = np.sqrt(np.mean(audio_data**2))
            level_bars = int(audio_level / 1000)
            level_display = "█" * min(level_bars, 20)
            print(f"\r🔊 Level: [{level_display:<20}] {audio_level:.0f} ", end="", flush=True)
        print("\n✅ Microphone test complete!")
        stream_test.stop_stream()
        stream_test.close()
except Exception as e:
        print(f"❌ Microphone test failed: {e}")
finally:
        pa_test.terminate()

def open_chrome():
    try:
        if os.name == 'nt':
            os.system("start chrome")
        else:
            os.system("google-chrome &")
        print("Chrome opened successfully")
    except Exception as e:
        print(f"Error opening Chrome: {e}")
        speak("I encountered an error opening Chrome, sir.")

def open_last_excel():
    try:
        project_path = r"C:\Users\Ahmed\OneDrive - regardian.com\GRC Projects - Projects\Tanmiah\3. Toolkit"
        if not os.path.exists(project_path):
            speak("I couldn't find your Excel project folder, sir.")
            return
        
        files = glob.glob(os.path.join(project_path, "*.xls*"))
        if not files:
            speak("No Excel files found in your project folder, sir.")
            return
        
        last_file = max(files, key=os.path.getmtime)
        if os.name == 'nt':
            os.system(f'start excel "{last_file}"')
        print(f"Excel project opened: {os.path.basename(last_file)}")
    except Exception as e:
        print(f"Error opening Excel project: {e}")
        speak("I encountered an error opening your Excel project, sir.")

def tell_plan():
    speak("You have no meetings today, sir. Free as a bird.")

# === Enhanced Command Handler ===

def handle_command(command):
    if not command or command.strip() == "":
        speak("I didn't catch that. Could you repeat?")
        return
    
    print(f"Processing command: '{command}'")
    command_clean = command.lower().strip()
    words = command_clean.split()
    
    # Check for training commands
    if "train command" in command_clean or "training mode" in command_clean:
        quick_acknowledge_and_execute(train_command, "Entering training mode, sir.")
        return
        
    if "show learned" in command_clean or "learned commands" in command_clean:
        quick_acknowledge_and_execute(show_learned_commands, "Showing your learned commands, sir.")
        return

    # Check for "open" commands first
    if "open" in command_clean:
        # Extract what to open (everything after "open")
        open_idx = command_clean.find("open")
        if open_idx != -1:
            program_name = command_clean[open_idx + 4:].strip()
            if program_name:
                # Try dynamic program opening first
                if open_program_dynamically(program_name):
                    return
    
    # Enhanced keyword detection
    chrome_keywords = ["chrome", "browser", "google", "web", "internet"]
    excel_keywords = ["excel", "spreadsheet", "worksheet", "xls", "xlsx"]
    plan_keywords = ["plan", "schedule", "agenda", "calendar", "meeting"]
    test_keywords = ["test", "microphone", "mic", "audio"]

    # Check for Excel project specifically
    if "last excel project" in command_clean or "my excel project" in command_clean:
        quick_acknowledge_and_execute(open_last_excel)
        return

    # Check for microphone test
    if any(keyword in command_clean for keyword in test_keywords):
        quick_acknowledge_and_execute(test_microphone, "Testing microphone now, sir.")
        return
    
    # Check for chrome
    if any(keyword in command_clean for keyword in chrome_keywords):
        quick_acknowledge_and_execute(open_chrome)
        return
        
    # Check for excel
    if any(keyword in command_clean for keyword in excel_keywords):
        quick_acknowledge_and_execute(open_last_excel)
        return
        
    # Check for plan
    if any(keyword in command_clean for keyword in plan_keywords):
        quick_acknowledge_and_execute(tell_plan)
        return
    
    # Common responses
    greeting_keywords = ["hello", "hi", "hey", "good morning", "good afternoon"]
    if any(keyword in command_clean for keyword in greeting_keywords):
        speak("Hello sir, how can I assist you?")
        return
        
    thanks_keywords = ["thank you", "thanks", "thank"]
    if any(keyword in command_clean for keyword in thanks_keywords):
        speak("You're welcome, sir.")
        return
        
    goodbye_keywords = ["goodbye", "bye", "see you", "exit", "quit"]
    if any(keyword in command_clean for keyword in goodbye_keywords):
        speak("Goodbye, sir. Have a great day!")
        return

    # If no specific match, try to find any program mentioned
    for word in words:
        if len(word) > 3:  # Only try words longer than 3 characters
            if open_program_dynamically(word):
                return
    
    # Final fallback
    speak("I'm not sure how to help with that, sir. Try saying 'open' followed by the program name.")

# === Main Loop ===
print("IRIS with RAG is ready! Listening for wake word...")
speak("Good morning sir. I'm learning about your system in the background.")

try:
    while True:
        try:
            pcm = stream.read(porcupine.frame_length, exception_on_overflow=False)
            pcm_unpacked = struct.unpack_from("h" * porcupine.frame_length, pcm)
            keyword_index = porcupine.process(pcm_unpacked)

            if keyword_index >= 0:
                print("Wake word detected!")
                speak("Yes, sir?")
                time.sleep(0.05)
                cmd = listen_and_transcribe()
                if cmd:
                    handle_command(cmd)
                else:
                    speak("I didn't hear anything. Please try again.")

        except Exception as e:
            print(f"Error in main loop: {e}")
            continue

except KeyboardInterrupt:
    print("Shutting down IRIS...")
    speak("Goodbye, sir.")
finally:
    try:
        stream.stop_stream()
        stream.close()
        pa.terminate()
        porcupine.delete()
        pygame.mixer.quit()
    except:
        pass
    print("IRIS offline.")