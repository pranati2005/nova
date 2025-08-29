import speech_recognition as sr
import pyttsx3
import subprocess
import os
import shutil

# Initialize recognizer and TTS engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Clipboard for copy-paste buffer
clipboard = []

# Track current working directory (default = user's home)
current_dir = os.path.expanduser("~")

# ----------------- SPEAK -----------------
def speak(text):
    """Convert text to speech output."""
    engine.say(text)
    engine.runAndWait()

# ----------------- LISTEN -----------------
def listen_command():
    """Listen for a voice command and return recognized text."""
    with sr.Microphone() as source:
        print("🎤 Listening...")
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio, language="en-in")
            print(f"👉 Heard: {command}")
            return command.lower()
        except sr.UnknownValueError:
            return ""
        except sr.RequestError:
            speak("Sorry, my speech service is down.")
            return ""

# ----------------- APPLICATION CONTROL -----------------
def open_application(app_name):
    """Open applications with predefined paths + Windows fallback search."""
    apps = {
        "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "google chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "notepad": "notepad.exe",
        "vs code": r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe",
        "visual studio code": r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe",
        "calculator": "calc.exe",
        "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "microsoft edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "spotify": r"C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe",
        "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
        "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
    }

    if "file explorer" in app_name:
        subprocess.Popen("explorer")
        speak("Opening File Explorer")
        return

    for key in apps.keys():
        if key in app_name:
            try:
                subprocess.Popen(apps[key], shell=True)
                speak(f"Opening {key}")
                return
            except Exception as e:
                speak(f"Error opening {key}: {e}")
                return

    # Fallback: use Windows Start search
    try:
        subprocess.Popen(f"start {app_name}", shell=True)
        speak(f"Opening {app_name}")
    except Exception as e:
        speak(f"Sorry, I couldn’t find {app_name}.")
        print(f"[APP OPEN ERROR] {e}")

# ----------------- FILE NAVIGATION -----------------
def navigate(command):
    """Navigate to common directories like Desktop, Downloads."""
    global current_dir
    home = os.path.expanduser("~")

    if "desktop" in command:
        current_dir = os.path.join(home, "Desktop")
        os.chdir(current_dir)
        subprocess.Popen(f"explorer {current_dir}")
        speak("Opening Desktop")

    elif "downloads" in command:
        current_dir = os.path.join(home, "Downloads")
        os.chdir(current_dir)
        subprocess.Popen(f"explorer {current_dir}")
        speak("Opening Downloads folder")

    elif "documents" in command:
        current_dir = os.path.join(home, "Documents")
        os.chdir(current_dir)
        subprocess.Popen(f"explorer {current_dir}")
        speak("Opening Documents")

    else:
        speak("Navigation folder not recognized")

# ----------------- FILE MANAGEMENT -----------------
def file_management(command):
    """Handle file and folder operations with selections."""
    global clipboard, current_dir
    try:
        words = command.split()

        # ----- Select All -----
        if "select all" in command:
            clipboard = [os.path.join(current_dir, f) for f in os.listdir(current_dir)]
            speak(f"Selected all items in {current_dir}")

        # ----- Delete All -----
        elif "delete all" in command:
            for item in os.listdir(current_dir):
                path = os.path.join(current_dir, item)
                try:
                    if os.path.isfile(path):
                        os.remove(path)
                    else:
                        shutil.rmtree(path)
                except Exception as e:
                    print(f"Error deleting {path}: {e}")
            speak("Deleted all items")

        # ----- Copy All -----
        elif "copy all" in command:
            clipboard = [os.path.join(current_dir, f) for f in os.listdir(current_dir)]
            speak("Copied all items")

        # ----- Paste All -----
        elif "paste all" in command:
            for item in clipboard:
                if os.path.exists(item):
                    try:
                        if os.path.isfile(item):
                            shutil.copy(item, current_dir)
                        else:
                            dest = os.path.join(current_dir, os.path.basename(item))
                            if os.path.exists(dest):
                                dest += "_copy"
                            shutil.copytree(item, dest)
                    except Exception as e:
                        print(f"Error pasting {item}: {e}")
            speak("Pasted all items")

        # ----- Create Folder -----
        elif "create" in words and "folder" in words:
            folder_name = words[-1]
            path = os.path.join(current_dir, folder_name)
            os.makedirs(path, exist_ok=True)
            speak(f"Folder {folder_name} created")

        # ----- Delete File -----
        elif "delete file" in command:
            file_name = words[-1]
            path = os.path.join(current_dir, file_name)
            if os.path.exists(path):
                os.remove(path)
                speak(f"File {file_name} deleted")
            else:
                speak(f"File {file_name} not found")

        # ----- Delete Folder -----
        elif "delete folder" in command:
            folder_name = words[-1]
            path = os.path.join(current_dir, folder_name)
            if os.path.exists(path):
                shutil.rmtree(path)
                speak(f"Folder {folder_name} deleted")
            else:
                speak(f"Folder {folder_name} not found")

        # ----- Rename -----
        elif "rename" in words:
            old_name = words[words.index("rename") + 1]
            new_name = words[-1]
            old_path = os.path.join(current_dir, old_name)
            new_path = os.path.join(current_dir, new_name)
            if os.path.exists(old_path):
                os.rename(old_path, new_path)
                speak(f"Renamed {old_name} to {new_name}")
            else:
                speak(f"{old_name} not found")

        # ----- Copy File -----
        elif "copy" in words and "all" not in words:
            src = words[words.index("copy") + 1]
            clipboard = [os.path.join(current_dir, src)]
            speak(f"Copied {src}")

        # ----- Paste File -----
        elif "paste" in words and "all" not in words:
            for item in clipboard:
                if os.path.exists(item):
                    if os.path.isfile(item):
                        shutil.copy(item, current_dir)
                    else:
                        dest = os.path.join(current_dir, os.path.basename(item))
                        shutil.copytree(item, dest)
            speak("Pasted items")

        else:
            speak("File command not recognized")

    except Exception as e:
        speak(f"Error handling files: {e}")
        print(f"[FILE ERROR] {e}")

# ----------------- SYSTEM CONTROL -----------------
def system_control(command):
    """Basic system control like shutdown, restart, volume."""
    if "shutdown" in command:
        speak("Shutting down your system")
        os.system("shutdown /s /t 1")
    elif "restart" in command:
        speak("Restarting your system")
        os.system("shutdown /r /t 1")
    elif "volume up" in command:
        os.system("nircmd.exe changesysvolume 2000")
        speak("Volume increased")
    elif "volume down" in command:
        os.system("nircmd.exe changesysvolume -2000")
        speak("Volume decreased")
    else:
        speak("System command not recognized")

# ----------------- PROCESS COMMAND -----------------
def process_command(command):
    """Process and route commands to appropriate functions."""
    if "open" in command:
        app_name = command.replace("nova", "").replace("open", "").strip()
        open_application(app_name)

    elif any(word in command for word in ["create", "delete", "rename", "copy", "paste", "select"]):
        file_management(command)

    elif any(word in command for word in ["desktop", "downloads", "documents"]):
        navigate(command)

    elif any(word in command for word in ["shutdown", "restart", "volume"]):
        system_control(command)

    else:
        speak("Sorry, I didn’t understand that command.")

# ----------------- MAIN LOOP -----------------
def main():
    speak("Hello, I am Nova. Say 'Nova' followed by a command.")
    while True:
        command = listen_command()
        if "nova" in command:
            process_command(command)

if __name__ == "__main__":
    main()
