import pyttsx3

engine = pyttsx3.init()
engine.setProperty("rate", 170)

voices = engine.getProperty("voices")
for v in voices:
    print(v.id, v.name)

engine.say("Testing voice system, sir.")
engine.runAndWait()
