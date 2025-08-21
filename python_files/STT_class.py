import speech_recognition as sr

class STT():
    def __init__(self, tts = None, ambient_duration = 1, listen_timeout = 6, listen_phrase_time_limit = 5):
        self.recognizer = sr.Recognizer()
        self.ambient_duration = ambient_duration
        self.listen_timeout = listen_timeout
        self.listen_phrase_time_limit = listen_phrase_time_limit
        self.tts = tts

    def listen(self):
        try:
            with sr.Microphone() as source:
                print('Listening . . .')
                self.recognizer.adjust_for_ambient_noise(source, duration = self.ambient_duration)
                self.audio = self.recognizer.listen(source, timeout = self.listen_timeout, phrase_time_limit = self.listen_phrase_time_limit)
                try:
                    self.command = self.recognizer.recognize_google(self.audio)
                    print(f'You said: {self.command}')
                    return self.command.lower().strip()
                except sr.UnknownValueError:
                    print('Sorry, I did not understand that')
                    return ''
                except sr.RequestError:
                    self.tts.speak('ERROR: Speech service error!')
                    return ''
        except sr.WaitTimeoutError:
            return ""
        except OSError as e:
            self.tts.speak(f'ERROR: {e}')
            return ""
        
