from Executor import Executor 
from Command_handler import Command_Handler 
from TTS_class import TTS
from STT_class import STT 
import platform

IS_WINDOWS = platform.system() == 'Windows'

if __name__ == '__main__':
    xec = Executor()
    c_h = Command_Handler(xec=xec) 
    tts = TTS(rate = 0, volume = 100)
    stt = STT(tts = tts)

    if IS_WINDOWS:
        xec.index_windows_apps()
        tts.speak('Indexing the apps.')
    else:
        print("[Index] Non-Windows OS: skipping Start Menu/UWP indexing.")

    
    tts.speak('Jarvis is Online')
    tts.speak('Greetings')
    

    while True:
        command = stt.listen()
        if not command:
            continue
        result = c_h.handle_command(command)
        if result == '__EXIT__':
            tts.speak('Goodbye!')
            break
        tts.speak(result)
    
    tts.shutdown()
