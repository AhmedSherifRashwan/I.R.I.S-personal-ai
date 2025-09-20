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

# === Setup ===
ACCESS_KEY = "5cfCv3qXiQfV5vUl0HkaS1fcp22yIju+lljpqDOn0vtSD4U7PP3VWQ=="  # put your key from Picovoice
print("Initializing JARVIS...")
os.environ["PATH"] += os.pathsep + r"C:\ffmpeg\bin"

# Wake word engine
porcupine = pvporcupine.create(access_key=ACCESS_KEY, keywords=["jarvis"])
pa = pyaudio.PyAudio()
stream = pa.open(rate=porcupine.sample_rate,
                 channels=1,
                 format=pyaudio.paInt16,
                 input=True,
                 frames_per_buffer=porcupine.frame_length)

# Speech-to-text model
print("Loading Whisper model...")
stt_model = whisper.load_model("base")  # try "tiny" for speed

# Text-to-speech engine
engine = pyttsx3.init()
voices = engine.getProperty("voices")
# Use default voice if UK voice not found
uk_voice_found = False
for v in voices:
    if v and hasattr(v, 'name') and v.name and "english" in v.name.lower() and "uk" in v.name.lower():
        engine.setProperty("voice", v.id)
        uk_voice_found = True
        break

if not uk_voice_found:
    print("UK voice not found, using default voice")

engine.setProperty("rate", 170)

def speak(text):
    print(f"JARVIS: {text}")
    engine.say(text)
    engine.runAndWait()
    

# === Listening and transcription ===
def listen_and_transcribe(duration=4):
    """
    Records audio from the same stream used by Porcupine,
    for a fixed duration (default 4s), then transcribes it.
    """
    print("🎤 Recording command... Speak now!")

    frames = []
    for _ in range(int(porcupine.sample_rate / 1024 * duration)):
        data = stream.read(1024, exception_on_overflow=False)
        frames.append(data)

    # Save audio
    temp_file_path = os.path.abspath("temp_command.wav")
    wf = wave.open(temp_file_path, 'wb')
    wf.setnchannels(1)
    wf.setsampwidth(pa.get_sample_size(pyaudio.paInt16))
    wf.setframerate(porcupine.sample_rate)
    wf.writeframes(b''.join(frames))
    wf.close()

    # Transcribe with Whisper
    print("🔄 Transcribing...")
    result = stt_model.transcribe(
        temp_file_path,
        language="en",
        temperature=0.0,
        initial_prompt="Commands like open chrome, open browser, open excel, open spreadsheet"
    )
    text = result["text"].strip().lower()
    print(f"📝 Heard: {text}")

    # Cleanup
    try:
        os.remove(temp_file_path)
    except:
        pass

    return text


# === Skills ===
def test_microphone():
    """Test microphone and show audio levels"""
    print("🎤 Testing microphone for 5 seconds...")
    
    pa_test = pyaudio.PyAudio()
    FORMAT = pyaudio.paInt16
    CHANNELS = 1
    RATE = 16000
    CHUNK = 1024
    
    try:
        stream = pa_test.open(format=FORMAT,
                             channels=CHANNELS,
                             rate=RATE,
                             input=True,
                             frames_per_buffer=CHUNK)
        
        print("Speak now to test your microphone...")
        for i in range(int(RATE / CHUNK * 5)):  # 5 seconds
            data = stream.read(CHUNK, exception_on_overflow=False)
            audio_data = np.frombuffer(data, dtype=np.int16)
            audio_level = np.sqrt(np.mean(audio_data**2))
            
            level_bars = int(audio_level / 1000)
            level_display = "█" * min(level_bars, 20)
            print(f"\r🔊 Level: [{level_display:<20}] {audio_level:.0f}   ", end="", flush=True)
        
        print("\n✅ Microphone test complete!")
        stream.stop_stream()
        stream.close()
        
    except Exception as e:
        print(f"❌ Microphone test failed: {e}")
    finally:
        pa_test.terminate()

def open_chrome():
    try:
        if os.name == 'nt':  # Windows
            os.system("start chrome")
        else:  # Linux/Mac
            os.system("google-chrome &")
        speak("Opening Chrome.")
    except Exception as e:
        print(f"Error opening Chrome: {e}")
        speak("I couldn't open Chrome.")

def open_last_excel():
    try:
        # Path to your project folder
        project_path = r"C:\Users\Ahmed\OneDrive - regardian.com\GRC Projects - Projects\Tanmiah\3. Toolkit"
        
        if not os.path.exists(project_path):
            speak("I couldn't find your Excel project folder.")
            return

        # Search for Excel files in the folder (xls, xlsx, xlsm, etc.)
        files = glob.glob(os.path.join(project_path, "*.xls*"))

        if not files:
            speak("No Excel files found in your project folder.")
            return

        # Get the most recently modified file
        last_file = max(files, key=os.path.getmtime)

        # Open it in Excel
        if os.name == 'nt':  # Windows
            os.system(f'start excel "{last_file}"')

        speak(f"Opening your last Excel project: {os.path.basename(last_file)}")

    except Exception as e:
        print(f"Error opening Excel project: {e}")
        speak("I couldn't open your last Excel project.")

def tell_plan():
    # TODO: integrate Google Calendar
    speak("You have no meetings today, sir. Free as a bird.")

# === LLM Integration ===
def interpret_command(command):
    prompt = f"""
    You are JARVIS, a desktop AI assistant.
    The user said: "{command}".
    Map this to an action if possible.

    Action mapping rules:
    - If they mention "open chrome", "start chrome", "launch chrome", "browser" → action=open_chrome
    - If they mention "excel", "spreadsheet", "worksheet" → action=open_excel
    - If they mention "plan", "schedule", "agenda" → action=tell_plan

    Otherwise, set action=none and just reply conversationally.

    Respond ONLY in JSON:
    {{
      "action": "none | open_chrome | open_excel | tell_plan",
      "response": "your natural language reply"
    }}
    """
    
    try:
        res = ollama.chat(model="mistral",
                          messages=[{"role":"user","content":prompt}])
        response_content = res['message']['content'].strip()
        
        # Clean the response to ensure it's valid JSON
        if response_content.startswith('```json'):
            response_content = response_content.replace('```json', '').replace('```', '').strip()
        elif response_content.startswith('```'):
            response_content = response_content.replace('```', '').strip()
            
        return json.loads(response_content)
        
    except Exception as e:
        print(f"Error with LLM interpretation: {e}")
        return {"action": "none", "response": "I'm not sure how to handle that."}

def handle_command(command):
    if not command or command.strip() == "":
        speak("I didn't catch that. Could you repeat?")
        return
    
    print(f"Processing command: '{command}'")
    print(f"Command length: {len(command)} characters")
    
    # Clean and normalize the command
    command_clean = command.lower().strip()
    print(f"Cleaned command: '{command_clean}'")
    
    # Split command into words for better matching
    words = command_clean.split()
    print(f"Command words: {words}")
    
    # --- Enhanced keyword detection with multiple variations ---
    chrome_keywords = ["chrome", "browser", "google", "web", "internet"]
    excel_keywords = ["excel", "spreadsheet", "worksheet", "xls", "xlsx"]
    plan_keywords = ["plan", "schedule", "agenda", "calendar", "meeting"]
    test_keywords = ["test", "microphone", "mic", "audio"]

    if "last excel project" in command_clean or "my excel project" in command_clean:
        open_last_excel()
    return

    
    # Check for microphone test
    if any(keyword in command_clean for keyword in test_keywords):
        print(f"Test keyword detected in: '{command_clean}'")
        test_microphone()
        return
    
    # Check if any chrome keywords are present
    if any(keyword in command_clean for keyword in chrome_keywords):
        print(f"Chrome keyword detected in: '{command_clean}'")
        open_chrome()
        return
        
    # Check if any excel keywords are present  
    if any(keyword in command_clean for keyword in excel_keywords):
        print(f"Excel keyword detected in: '{command_clean}'")
        open_last_excel()
        return
        
    # Check if any plan keywords are present
    if any(keyword in command_clean for keyword in plan_keywords):
        print(f"Plan keyword detected in: '{command_clean}'")
        tell_plan()
        return
    
    # --- Also check individual words for exact matches ---
    for word in words:
        if word in chrome_keywords:
            print(f"Chrome word match: '{word}'")
            open_chrome()
            return
        elif word in excel_keywords:
            print(f"Excel word match: '{word}'")
            open_last_excel()
            return
        elif word in plan_keywords:
            print(f"Plan word match: '{word}'")
            tell_plan()
            return
    
    # --- Common greetings and simple responses ---
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

    # --- Fallback: Try simple pattern matching ---
    if "open" in command_clean:
        if any(word in command_clean for word in ["chrome", "browser", "google"]):
            print("Fallback: Detected 'open' + browser keyword")
            open_chrome()
            return
        elif any(word in command_clean for word in ["excel", "spreadsheet"]):
            print("Fallback: Detected 'open' + excel keyword")
            open_last_excel()
            return

    # --- Otherwise, let the AI interpret ---
    print("No direct keyword matches found, trying AI interpretation...")
    try:
        decision = interpret_command(command)
        action = decision.get("action", "none")
        response = decision.get("response", "I'm not sure how to help with that.")
        
        print(f"AI decision - Action: {action}, Response: {response}")

        if action == "open_chrome":
            open_chrome()
        elif action == "open_excel":
            open_last_excel()
        elif action == "tell_plan":
            tell_plan()
        else:
            speak(response)
    except Exception as e:
        print(f"Error handling command: {e}")
        speak("I'm having trouble processing that command.")

# === Main loop ===
print("JARVIS is ready! Listening for wake word 'Jarvis'...")
speak("JARVIS is online and ready, sir.")

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
    print("Shutting down JARVIS...")
    speak("Goodbye, sir.")
finally:
    try:
        stream.stop_stream()
        stream.close()
        pa.terminate()
        porcupine.delete()
    except:
        pass
    print("JARVIS offline.")