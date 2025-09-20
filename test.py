import speech_recognition as sr
import pyttsx3
import webbrowser
import requests
import json
import datetime

# --- Setup the speech engine (for speaking) ---
engine = pyttsx3.init()

def speak(text):
    """Convert text to speech."""
    engine.say(text)
    engine.runAndWait()

# --- Setup the listener (for hearing) ---
recogniser = sr.Recognizer()
mic = sr.Microphone()

def listen():
    """Listen from microphone and return recognized text (or None)."""
    with mic as source:
        recogniser.adjust_for_ambient_noise(source)
        print("Listening...")
        audio = recogniser.listen(source)
    try:
        command = recogniser.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        print("Sorry, I did not catch that.")
        return None
    except sr.RequestError:
        print("Speech service is down.")
        return None

# --- Core respond logic ---
def respond(command):
    """Process the command and respond appropriately."""
    if command is None:
        return

    # Greeting
    if "hello" in command or "hi" in command:
        speak("Hello! How can I help you today?")
    
    # Time enquiries
    elif "time" in command:
        now = datetime.datetime.now().strftime("%H:%M")
        speak(f"The time is {now}")

    # Open website
    elif "open google" in command:
        speak("Opening Google")
        webbrowser.open("https://www.google.com")
    
    elif "open youtube" in command:
        speak("Opening YouTube")
        webbrowser.open("https://www.youtube.com")

    # Weather updates
    elif "weather" in command:
        api_key = "YOUR_OPENWEATHERMAP_API_KEY"
        city = "London"  # or parse it from command
        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
        try:
            resp = requests.get(url)
            data = resp.json()
            if data.get("weather"):
                description = data["weather"][0]["description"]
                temp = data["main"]["temp"]
                speak(f"The weather in {city} is currently {description} with a temperature of {temp}°C")
            else:
                speak("I couldn't get the weather right now.")
        except requests.RequestException:
            speak("Failed to connect to the weather service.")

    # Jokes
    elif "joke" in command:
        # a simple joke
        speak("Why did the programmer quit his job? Because he didn't get arrays.")
    
    # Custom / Miscellaneous
    else:
        speak("I am sorry, I don’t know how to do that yet.")

# --- Main loop ---
def main():
    speak("Assistant is now online.")
    while True:
        command = listen()
        if command:
            if "exit" in command or "quit" in command or "stop" in command:
                speak("Goodbye!")
                break
            respond(command)

if __name__ == "__main__":
    main()
