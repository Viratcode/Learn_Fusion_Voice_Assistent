import speech_recognition as sr
import win32com.client
import datetime
import webbrowser

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Adjusting for ambient noise...")
        r.adjust_for_ambient_noise(source)
        print("Listening...")
        r.pause_threshold = 1
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=5)
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"user said: {query}")
        except sr.WaitTimeoutError:
            print("Listening timed out while waiting for phrase to start")
            return "None"
        except sr.UnknownValueError:
            print("Google Speech Recognition could not understand the audio")
            return "None"
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return "None"
        return query.lower()


if __name__ == '__main__':
    print('Pycharm')
    say('Hello, I am Parth AI')
    query = takeCommand()
    print(f"Received command: {query}")

    sites = [
        ["youtube", "https://youtube.com"],
        ["wikipedia", "https://wikipedia.com"],
        ["google", "https://google.com"],
    ]


    if query is not None:
        for site in sites:
            if f"open {site[0]}" in query:
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
                break
        else:
            if "the time" in query:
                strftime = datetime.datetime.now().strftime("%H:%M:%S")
                say(f"The time is {strftime}")
    else:
        say("I did not understand your command.")
