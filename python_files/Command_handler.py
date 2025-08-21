import re

class Command_Handler():
    def __init__(self, xec = None): 
        
        self._OPEN_SITE_PAT = re.compile(r"^(open|launch)\s+(?P<what>youtube|gmail|google|github|notion|spotify|[a-z0-9\.\-]+)$")
        self._SEARCH_PAT    = re.compile(r"^(google|search|find)\s+(for\s+)?(?P<q>.+)$")
        self._TIME_PAT      = re.compile(r"(what('s| is)?\s+)?(the\s+)?time(\s+now)?\??$", re.I)
        self._DATE_PAT      = re.compile(r"(what('s| is)?\s+)?(the\s+)?date(\s+today)?\??$", re.I)
        self._NOTE_PAT      = re.compile(r"(make|take|add|note)\s+(that\s*)?(?P<text>.+)", re.I)
        self._FOLDER_PAT    = re.compile(r"^(open)\s+(?P<folder>downloads|documents|desktop)\s+(folder)?$", re.I)
        if not xec:
            return print(f'ERROR: NO EXECUTOR WAS PASSED')    
        self.xec = xec

    def handle_command(self, command: str):
        self.cmd = command.strip()
        if self.cmd in ('exit', 'quit', 'stop'):
            return "__EXIT__"
        
        if self.cmd.startswith('open '):
            app_name = self.cmd[5:].strip()
            return self.xec.launch_windows_apps(app_name)
        
        self.m = self._OPEN_SITE_PAT.match(self.cmd)
        if self.m:
            return self.xec.open_site(self.m.group('what'))
        
        self.m = self._SEARCH_PAT.match(self.cmd)
        if self.m:
            return self.xec.google_search(self.m.group('q'))
        
        self.m = self._TIME_PAT.match(self.cmd)
        if self.m:
            return self.xec.tell_time()
        
        self.m = self._DATE_PAT.match(self.cmd)
        if self.m:
            return self.xec.tell_date()
        
        self.m = self._NOTE_PAT.match(self.cmd)
        if self.m:
            return self.xec.make_note(self.m.group('text'))
        
        self.m = self._FOLDER_PAT.match(self.cmd)
        if self.m:
            return self.xec.open_folder(self.m.group('folder'))
        
        return "I haven't been modelled for that action!"
    