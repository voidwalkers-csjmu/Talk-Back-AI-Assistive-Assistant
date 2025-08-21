import webbrowser
import subprocess
import os
import json
from datetime import datetime
from urllib.parse import quote_plus
from pathlib import Path
import platform

class Executor():
    def __init__(self):
        self._APP_INDEX = {
            "shortcuts": {},
            "exes": {},
            "uwp": {}
        }
        self._IS_WINDOWS = platform.system() == 'Windows'
        self._IS_MAC = platform.system() == 'Darwin'
        self._IS_LINUX = platform.system() == 'Linux'

    def index_windows_apps(self):
        self.user = os.environ.get('USERNAME') or os.environ.get('USER')
        self.search_paths = [
            Path(f"C:/Users/{self.user}/OneDrive/Desktop"),
            Path(f"C:/ProgramData/Microsoft/Windows/Start Menu/Programs"),
            Path(f"C:/Users/{self.user}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs")
        ] 

        for path in self.search_paths:
            if path and path.exists():
                for root, _, files in os.walk(path):
                    for file in files:
                        low = file.lower()
                        full = os.path.join(root, file)
                        if low.endswith('.lnk'):
                            self._APP_INDEX['shortcuts'][low.replace('.lnk', "")] = full
                        elif low.endswith('.exe'):
                            self._APP_INDEX['exes'][low.replace('.exe', "")] = full
        
        try:
            self.cmd = ["powershell", "-NoProfile", "-Command", "Get-StartApps | ConvertTo-Json -Compress"]
            self.completed_process = subprocess.run(self.cmd, capture_output = True, timeout = 8)
            self.out = self.completed_process.stdout.decode('utf-8').strip()
            if self.out:
                data = json.loads(self.out)
                if isinstance(data, dict):
                    data = [data]
                for app in data:
                    name = str(app.get("Name", "")).lower()
                    appid = app.get("AppID")
                    if name and appid:
                        self._APP_INDEX['uwp'][name] = appid
        except Exception as e:
            print(f"INDEX ERROR: {e}")

    def launch_windows_apps(self, app_name: str)-> str:
        self.name = app_name.lower()

        if self.name in self._APP_INDEX['shortcuts']:
            os.startfile(self._APP_INDEX['shortcuts'][self.name])
            return f"Opening {self.name}"
        
        if self.name in self._APP_INDEX['exes']:
            os.startfile(self._APP_INDEX['exes'][self.name])
            return f"Opening {self.name}"
        
        if self.name in self._APP_INDEX['uwp']:
            subprocess.Popen(["explorer.exe", f"shell:appsFolder\\{self._APP_INDEX['uwp'][self.name]}"])
            return f"Opening {self.name}"
        
        return f"Application {self.name} not found"
    
    def open_site(self, alias_or_url:str)-> str:
        self.known_sites = {
            "youtube": "https://www.youtube.com",
            "gmail": "https://mail.google.com",
            "google": "https://www.google.com",
            "github": "https://github.com",
            "notion": "https://www.notion.so",
            "spotify": "https://open.spotify.com",
        }
        self.url = self.known_sites.get(alias_or_url.lower(), alias_or_url)
        if not self.url.startswith("http"):
            self.url = "https://" + self.url
        webbrowser.open(self.url)

        return f"Opening {self.url}"
        
    def google_search(self, query:str)-> str:
        self.search_url = f"https://www.google.com/search?q={quote_plus(query)}"
        webbrowser.open(self.search_url)
        return f"Searching Google for {query}"
    
    def tell_date(self) -> str:
        self.current_date = datetime.now().strftime('%A, %B %d, %Y')
        return f"Today's date is {self.current_date}"
    
    def tell_time(self) -> str:
        self.current_time = datetime.now().strftime('%I:%M %p')
        return f"The current time is {self.current_time}"
    
    def make_note(self, note:str) -> str:
        self.note_path = os.path.join(os.path.expanduser('~'), "voice_ai_notes.txt")
        self.timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        with open(self.note_path, 'a', encoding = 'utf-8') as note_file:
            note_file.write(f'[{self.timestamp}]\n{note}\n')
        return f"Saved your note."
    
    def open_folder(self, name:str) -> str:
        self.folder_map = {
            "downloads": os.path.join(os.path.expanduser('~'), "Downloads"),
            "documents": os.path.join(os.path.expanduser('~'), "Documents"),
            "desktop"  : os.path.join(os.path.expanduser('~'), "Desktop"),
        }
        self.target_folder = self.folder_map.get(name.lower())
        if not self.target_folder or not os.path.exists(self.target_folder):
            return f"I couldn't find {self.target_folder} folder"
        if self._IS_WINDOWS:
            subprocess.Popen(["explorer", self.target_folder])
        elif self._IS_MAC:
            subprocess.Popen(["open", self.target_folder])
        else:
            subprocess.Popen(["xdg-open", self.target_folder])
        return f"Opening {self.target_folder} folder"    