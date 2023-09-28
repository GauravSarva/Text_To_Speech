import win32com.client
speakers = win32com.client.Dispatch("SAPI.SpVoice")
names = ["Darshan","Deepak","Harsh", "Katrina", "John"]
for name in names:
    speakers.Speak(name)
