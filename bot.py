import win32com.client
import os
import speech_recognition as sr
import webbrowser
import datetime
from AppOpener import open
import Chat
import time
speaker = win32com.client.Dispatch("SAPI.SpVoice")
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing .....")
            query = r.recognize_google(audio, language="hn-in")
            print(f"user said: {query}")
            return query
        except Exception as e:
            return "sorry, i don't understand, can you please say it again"


def speak(text):
    os.system(f"say{text}")


if __name__ == "__main__":

    speaker.speak("Hello ! I am your personalized digital assistant! How may i help you ?")
    while True:
        print("listening.......")
        query = takecommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipeidia", "https://www.wikipeida.com"],
                 ["google", "https://www.google.com"], ["instagram", "https://www.instagram.com"],
                 ["gfg","https://www.geeksforgeeks.org/myCourses"],["linkedin","https://www.linkedin.com/feed/"]]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                speaker.speak(f"searching for {site[0]} sir....")
                webbrowser.open(site[1])
        musicpath = [["let me love you", "C:/music-3.mp3"], ["summer cruel", "C:/music-5.mp3"],
                     ["mahi", "C:/O Mahi O Mahi_64(PagalWorld.com.sb).mp3"],
                     ["sooraj dooba", "C:/music-4.mp3"],
                     ["ilahi", "C:/music-2.mp3"]]
        for music in musicpath:
            if f"play {music[0]}".lower() in query.lower():
                speaker.speak(f"playing {music[0]} sir...")
                os.system(music[1])
                time.sleep(2)
        if "what's the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.speak(f"the time is : {strfTime} sir..")
            time.sleep(2)
        apps = [["mail"], ["Goole"], ["Microsoft Edge"],[ "File Explorer"],
                      ["VLC Media player"], ["OneNote"],
                     ["WhatsApp"],[ "Camera"]]
        for app in apps:
            if f"open {app[0]}".lower() in query.lower():
                speaker.speak(f"opening {app[0]} sir...")
                open(app[0])
                time.sleep(2)
        if "answer my question" in query.lower():
            Chat.chat_bot()
        if "stop" in query:
            break
