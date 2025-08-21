import speech_recognition as sr
import webbrowser
import subprocess
import os
import platform
import re
import json
import queue
import threading
import time
from datetime import datetime
from urllib.parse import quote_plus
from pathlib import Path

# =========================================================
# TTS: Windows SAPI (primary) with a queue + pyttsx3 fallback
# =========================================================
try:
    import win32com.client  # Requires: pip install pywin32
    _HAS_SAPI = True
except Exception:
    _HAS_SAPI = False

try:
    import pyttsx3  # Fallback only
    _HAS_PYTTSX3 = True
except Exception:
    _HAS_PYTTSX3 = False


class TTS:
    """Thread-safe TTS with a single worker and a message queue.

    On Windows we prefer native SAPI (win32com) for stability.
    If unavailable, we fall back to pyttsx3, still through a single worker.
    """

    SVSFlagsAsync = 1
    SVSFPurgeBeforeSpeak = 2

    def __init__(self, rate=None, volume=None, voice=None):
        self.queue = queue.Queue()
        self._stop_event = threading.Event()
        self._engine_kind = None

        self._sapi_voice = None
        self._pytts = None

        if platform.system() == "Windows" and _HAS_SAPI:
            # Primary: SAPI
            try:
                self._sapi_voice = win32com.client.Dispatch("SAPI.SpVoice")
                self._engine_kind = "sapi"
                # Optional tuning
                if rate is not None:
                    try:
                        self._sapi_voice.Rate = int(rate)  # -10..+10
                    except Exception:
                        pass
                if volume is not None:
                    try:
                        self._sapi_voice.Volume = int(volume)  # 0..100
                    except Exception:
                        pass
                if voice is not None:
                    # Attempt to pick a voice by substring match
                    try:
                        for v in self._sapi_voice.GetVoices():
                            if voice.lower() in v.GetDescription().lower():
                                self._sapi_voice.Voice = v
                                break
                    except Exception:
                        pass
            except Exception:
                self._sapi_voice = None
                self._engine_kind = None

        if self._engine_kind is None and _HAS_PYTTSX3:
            # Fallback: pyttsx3 (still single-threaded usage through worker)
            try:
                self._pytts = pyttsx3.init()
                self._engine_kind = "pyttsx3"
                if rate is not None:
                    try:
                        self._pytts.setProperty("rate", int(rate))
                    except Exception:
                        pass
                if volume is not None:
                    try:
                        self._pytts.setProperty("volume", float(volume) / 100.0)
                    except Exception:
                        pass
                if voice is not None:
                    try:
                        vs = self._pytts.getProperty("voices")
                        for v in vs:
                            if voice.lower() in v.name.lower():
                                self._pytts.setProperty("voice", v.id)
                                break
                    except Exception:
                        pass
            except Exception:
                self._pytts = None
                self._engine_kind = None

        if self._engine_kind is None:
            # Last resort: no TTS available
            print("[TTS] No TTS engine available. Text will be printed only.")

        self._worker = threading.Thread(target=self._loop, daemon=True)
        self._worker.start()

    def speak(self, text: str):
        """Queue text to speak (non-blocking)."""
        if not text:
            return
        self.queue.put(str(text))

    def stop(self):
        """Stop current speech and clear queue."""
        # Clear queue
        try:
            while not self.queue.empty():
                _ = self.queue.get_nowait()
                self.queue.task_done()
        except Exception:
            pass

        # Stop engine
        if self._engine_kind == "sapi" and self._sapi_voice is not None:
            try:
                # Purge current buffer by speaking empty string with purge flag
                self._sapi_voice.Speak("", self.SVSFPurgeBeforeSpeak | self.SVSFlagsAsync)
            except Exception:
                pass
        elif self._engine_kind == "pyttsx3" and self._pytts is not None:
            try:
                self._pytts.stop()
            except Exception:
                pass

    def shutdown(self):
        """Cleanly stop the worker."""
        self._stop_event.set()
        self.queue.put(None)
        self._worker.join(timeout=3)
        # No special teardown needed for SAPI or pyttsx3 here

    def _loop(self):
        while not self._stop_event.is_set():
            item = self.queue.get()
            if item is None:
                self.queue.task_done()
                break

            text = str(item)
            print(f"[Assistant]: {text}")

            try:
                if self._engine_kind == "sapi" and self._sapi_voice is not None:
                    # Speak async and wait until done in small chunks to be responsive
                    self._sapi_voice.Speak(text, self.SVSFlagsAsync)
                    # Wait roughly based on length, with a max cap
                    max_wait_ms = min(max(1500, len(text) * 60), 30000)  # heuristic
                    waited = 0
                    step = 150
                    while waited < max_wait_ms:
                        # 0 means still speaking
                        # WaitUntilDone returns True if completed within timeout
                        if self._sapi_voice.WaitUntilDone(step):
                            break
                        waited += step
                elif self._engine_kind == "pyttsx3" and self._pytts is not None:
                    self._pytts.say(text)
                    self._pytts.runAndWait()
                else:
                    # No engine: just print
                    pass
            except Exception as e:
                # If TTS fails mid-run, don't crash the loop
                print(f"[TTS Error] {e}")

            self.queue.task_done()


# Instantiate our TTS (tweak rate/volume if you like)
tts = TTS(rate=0, volume=100)  # Rate -10..+10 for SAPI / ~200 default for pyttsx3

def speak(text: str):
    tts.speak(text)

# =========================================================
# STT
# =========================================================
def listen():
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print('Listening...')
            # Faster settle but still useful; tweak if your room is noisy
            recognizer.adjust_for_ambient_noise(source, duration=0.3)
            # Add timeouts so we never freeze if nothing is said
            audio = recognizer.listen(source, timeout=6, phrase_time_limit=10)
            try:
                command = recognizer.recognize_google(audio)
                print(f'You said: {command}')
                return command.lower().strip()
            except sr.UnknownValueError:
                speak("Sorry, I did not understand that.")
                return ""
            except sr.RequestError:
                speak("Speech service error.")
                return ""
    except sr.WaitTimeoutError:
        # Just loop back around quietly
        return ""
    except OSError as e:
        speak(f"Microphone error: {e}")
        return ""

# =========================================================
# Detect OS
# =========================================================
IS_WINDOWS = platform.system() == 'Windows'
IS_LINUX   = platform.system() == 'Linux'
IS_MAC     = platform.system() == 'Darwin'

# =========================================================
# App Index (Windows)
# =========================================================
APP_INDEX = {
    "shortcuts": {},
    "exes": {},
    "uwp": {}
}

def index_windows_apps():
    user = os.environ.get("USERNAME") or os.environ.get("USER") or ""
    search_paths = [
        Path(f"C:/Users/{user}/OneDrive/Desktop"),
        Path(f"C:/ProgramData/Microsoft/Windows/Start Menu/Programs"),
        Path(f"C:/Users/{user}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs")
    ]

    for path in search_paths:
        if path and path.exists():
            for root, _, files in os.walk(path):
                for file in files:
                    low = file.lower()
                    full = os.path.join(root, file)
                    if low.endswith(".lnk"):
                        APP_INDEX["shortcuts"][low.replace(".lnk", "")] = full
                    elif low.endswith(".exe"):
                        APP_INDEX["exes"][low.replace(".exe", "")] = full

    # UWP Apps
    try:
        # Avoid blocking: short timeout, no shell=True needed here
        cmd = ["powershell", "-NoProfile", "-Command", "Get-StartApps | ConvertTo-Json -Compress"]
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=8)
        out = completed.stdout.strip()
        if out:
            data = json.loads(out.encode("utf-8").decode("utf-8-sig"))
            if isinstance(data, dict):
                data = [data]
            for app in data:
                name = str(app.get("Name", "")).lower()
                appid = app.get("AppID")
                if name and appid:
                    APP_INDEX["uwp"][name] = appid
    except Exception as e:
        print(f"[Index Error] {e}")

# =========================================================
# App Launcher
# =========================================================
def launch_windows(app_name: str):
    name = app_name.lower().strip()

    if name in APP_INDEX["shortcuts"]:
        os.startfile(APP_INDEX["shortcuts"][name])
        return f"Opening {app_name}"

    if name in APP_INDEX["exes"]:
        os.startfile(APP_INDEX["exes"][name])
        return f"Opening {app_name}"

    if name in APP_INDEX["uwp"]:
        subprocess.Popen(["explorer.exe", f"shell:appsFolder\\{APP_INDEX['uwp'][name]}"])
        return f"Opening {app_name}"

    return f"Could not find {app_name}"

# =========================================================
# Websites
# =========================================================
def open_site(alias_or_url: str) -> str:
    known = {
        "youtube": "https://www.youtube.com",
        "gmail": "https://mail.google.com",
        "google": "https://www.google.com",
        "github": "https://github.com",
        "notion": "https://www.notion.so",
        "spotify": "https://open.spotify.com",
    }
    url = known.get(alias_or_url.lower(), alias_or_url)
    if not url.startswith("http"):
        url = "https://" + url
    webbrowser.open(url)
    return f"Opening {url}"

# =========================================================
# Google Search
# =========================================================
def google_search(query: str) -> str:
    url = f"https://www.google.com/search?q={quote_plus(query)}"
    webbrowser.open(url)
    return f"Searching Google for {query}"

# =========================================================
# Time / Date
# =========================================================
def tell_time() -> str:
    now = datetime.now().strftime("%I:%M %p")
    return f"The time is {now}."

def tell_date() -> str:
    today = datetime.now().strftime("%A, %B %d, %Y")
    return f"Today's date is {today}."

# =========================================================
# Notes
# =========================================================
def make_note(text: str) -> str:
    note_path = os.path.join(os.path.expanduser("~"), "voice_ai_notes.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(note_path, "a", encoding='utf-8') as f:
        f.write(f"[{timestamp}] {text}\n")
    return "Saved your note."

# =========================================================
# Folders
# =========================================================
def open_folder(name: str) -> str:
    folder_map = {
        "downloads": os.path.join(os.path.expanduser("~"), "Downloads"),
        "documents": os.path.join(os.path.expanduser("~"), "Documents"),
        "desktop":   os.path.join(os.path.expanduser("~"), "Desktop"),
    }
    target = folder_map.get(name.lower())
    if not target or not os.path.exists(target):
        return f"I couldn't find the {name} folder."
    if IS_WINDOWS:
        subprocess.Popen(["explorer", target])
    elif IS_MAC:
        subprocess.Popen(["open", target])
    else:
        subprocess.Popen(["xdg-open", target])
    return f"Opening {name} folder."

# =========================================================
# Command Patterns
# =========================================================
OPEN_SITE_PAT = re.compile(r"^(open|launch)\s+(?P<what>youtube|gmail|google|github|notion|spotify|[a-z0-9\.\-]+)$")
SEARCH_PAT    = re.compile(r"^(google|search|find)\s+(for\s+)?(?P<q>.+)$")
TIME_PAT      = re.compile(r"(what('s| is)?\s+)?(the\s+)?time(\s+now)?\??$", re.I)
DATE_PAT      = re.compile(r"(what('s| is)?\s+)?(the\s+)?date(\s+today)?\??$", re.I)
NOTE_PAT      = re.compile(r"(make|take|add|note)\s+(that\s*)?(?P<text>.+)", re.I)
FOLDER_PAT    = re.compile(r"^(open)\s+(?P<folder>downloads|documents|desktop)\s+(folder)?$", re.I)

# =========================================================
# Command Handler
# =========================================================
def handle_command(cmd: str) -> str:
    cmd = cmd.strip()

    if cmd in ("exit", "quit", "stop"):
        return "__EXIT__"

    if cmd.startswith("open "):
        app_name = cmd[5:].strip()
        return launch_windows(app_name)

    m = OPEN_SITE_PAT.match(cmd)
    if m:
        return open_site(m.group("what"))

    m = SEARCH_PAT.match(cmd)
    if m:
        return google_search(m.group("q"))

    if TIME_PAT.match(cmd):
        return tell_time()
    if DATE_PAT.match(cmd):
        return tell_date()

    m = NOTE_PAT.match(cmd)
    if m:
        return make_note(m.group("text"))

    m = FOLDER_PAT.match(cmd)
    if m:
        return open_folder(m.group("folder"))

    return "Sorry, I don't have an action for that yet."

# =========================================================
# Main
# =========================================================
if __name__ == "__main__":
    speak("Indexing apps, please wait...")
    if IS_WINDOWS:
        index_windows_apps()
    else:
        print("[Index] Non-Windows OS: skipping Start Menu/UWP indexing.")

    speak("Hello! What do you need?")
    while True:
        command = listen()
        if not command:
            continue
        result = handle_command(command)
        if result == "__EXIT__":
            speak("Goodbye!")
            break
        speak(result)

    # Clean TTS shutdown
    tts.shutdown()
