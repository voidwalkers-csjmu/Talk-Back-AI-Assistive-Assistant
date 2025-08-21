import platform
import queue
import threading

SVSFlagsAsync = 1
SVSFPurgeBeforeSpeak = 2

try:
    import win32com.client
    _HAS_SAPI = True
except ImportError:
    _HAS_SAPI = False

try:
    import pyttsx3
    _HAS_PYTTSX3 = True
except ImportError:
    _HAS_PYTTSX3 = False



class TTS:
    def __init__(self, voice=None, rate = None, volume = None):
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self.engine_kind = None
        self.sapi_voice = None

        if platform.system() == 'Windows'and _HAS_SAPI:
            try: 
                self.engine_kind = 'sapi'
                self.sapi_voice = win32com.client.Dispatch("SAPI.SpVoice")
                try:
                    self.sapi_voice.Rate = int(rate)
                except Exception:
                    pass
                try:
                    self.sapi_voice.Volume = int(volume)
                except Exception:
                    pass
                if voice:
                    try:
                        for v in self.sapi_voice.GetVoices():
                            if voice.lower() in v.GetDescription().lower():
                                self.sapi_voice.Voice = v
                                break
                    except Exception:
                        pass
            except Exception:
                self.engine_kind = None
                self.sapi_voice = None
        
        if self.engine_kind is None and _HAS_PYTTSX3:
            try:
                self.engine_kind = 'pyttsx3'
                self._pytts = pyttsx3.init()
                try:
                    self._pytts.setProperty('rate', int(rate))
                except Exception:
                    pass
                try:
                    self._pytts.setProperty('volume', float(volume) / 100.0)
                except Exception:
                    pass
                if voice:
                    try:

                        for v in self._pytts.getProperty('voices'):
                            if voice.lower() in v.name.lower():
                                self._pytts.setProperty('voice', v.id)
                                break
                    except Exception:
                        pass
            except Exception:
                self.engine_kind = None
                self._pytts = None
        
            if self.engine_kind is None:
                print('No TTS engine available. Please install pyttsx or win32com.client for TTS supprt.\nOnly text will appear in the console.')

        self._worker_thread = threading.Thread(target=self._loop, daemon =True)
        self._worker_thread.start()
        
    def speak(self, text: str):
        if not text:
            return
        self.queue.put(str(text))
    
    def stop(self):
        try:
            while not self.queue.empty():
               _ = self.queue.get_nowait()
               self.queue.task_done()
        except Exception:
            pass

        if self.engine_kind == 'sapi' and self.sapi_voice:
            try:
                self.sapi_voice.Speak('', SVSFPurgeBeforeSpeak|SVSFlagsAsync)
                # Stop the current speech ourge speech flag triggers the stop and async flag allows the speech to be stopped immediately
            except Exception:
                pass
        elif self.engine_kind == 'pyttsx3'and self._pytts:
            try:
                self._pytts.stop()
            except Exception:
                pass

    def shutdown(self):
        self.stop_event.set()
        self.queue.put(None)
        self._worker_thread.join(timeout = 3)

    def _loop(self):
        while not self.stop_event.is_set():
            item = self.queue.get()
            if item is None:
                break
            text = str(item)
            print(f'[Assistant]: {text}')

           
            try:
                if self.engine_kind == 'sapi' and self.sapi_voice:        
                    self.sapi_voice.Speak(text, SVSFlagsAsync)
                    max_wait_ms = min(max(500,len(text)*50), 30000) 
                    waited = 0
                    step = 150
                    while waited < max_wait_ms:
                        if self.sapi_voice.WaitUntilDone(step):
                            break
                        waited += step
                elif self.engine_kind == 'pyttsx3' and self._pytts:
                    self._pytts.say(text)
                    self._pytts.runAndWait()
                else:
                    pass
            except Exception as e:
                print(f'[TTS Error]: {e}')
            
            self.queue.task_done()


