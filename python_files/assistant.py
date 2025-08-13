import speech_recognition as sr
import pyttsx3

engine = pyttsx3.init()

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print('Listening . . .')
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            print(f'You said: {command}')
            return command.lower()
        except sr.UnknownValueError:
            print('Sorry, I did not understand that.')
            return ""
        except sr.RequestError:
            print('Speech Service Error.')
            return ""

# Main

if __name__ == "__main__":
    speak("Hello! What do you need?")
    while True:
        command = listen()
        if 'exit' in command or 'quit' in command:
            speak('Goodbye!')
            break
        elif command:
            speak(f'You said: {command}')