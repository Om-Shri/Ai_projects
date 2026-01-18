import speech_recognition as sr #pip install speechrecognition
import pyttsx3                  #pip install pyttsx3
import webbrowser
import time
import win32com.client as wc
from music import musics       #creat music file with key = name of song and value = link of song.
from pytube import Search      #pip install pytube . in this program it is used for play first video in search.
import pyautogui

r = sr.Recognizer()    
#Speaker system
speaker = wc.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()
speaker.Voice = voices.Item(1)   #Change voice(1=female,0=male).
speaker.Rate = 2                 #Change rate of speaking.
speaker.Volume = 100             #Change volume.


def speak(text):
    speaker.speak(text)
    


def process_command(command : str) -> any:    #Process command function which proccess given command to it.

    command = command.lower()

    if  "open youtube" in command:
        speak("Opening youtube")
        webbrowser.open("https://www.youtube.com/")

    elif "open whatsapp" in command:
        speak("opening whatsapp")
        webbrowser.open("https://web.whatsapp.com/")
    
    elif "open github" in command:
        speak("opening github")
        webbrowser.open("https://github.com/")

    elif "open instagram" in command :
        speak("opening instagram")
        webbrowser.open("https://www.instagram.com/")

    elif "open chat gpt" in command:
        speak("opening chatgpt")
        webbrowser.open("https://chatgpt.com/")

    elif "play" in command :
        song = command.replace("play", "").strip()

        if song in musics:
            speak(f"Playing {song}")
            webbrowser.open(musics[song])
        else:
            s = Search(song)
            video = s.results[0]
            speak(f"playing {song} ")
            webbrowser.open(video.watch_url)

    elif "turn off" in command:
        if "close" in command and "turn off" in command:
            pyautogui.hotkey("win","m")
            time.sleep(1)
            pyautogui.hotkey("alt","f4")
            time.sleep(1)
            pyautogui.click(978,520)
        else:
            pyautogui.hotkey("alt","f4")
            pyautogui.leftClick(978,520)

        
if __name__ == "__main__":
    speak("Initializing Komal")

    while True:
        
        try:
            with sr.Microphone(device_index=2) as source:
                print("Listening...")
                r.adjust_for_ambient_noise(source, duration=1)
                audio = r.listen(source, timeout=10, phrase_time_limit=1)

            print("Recognizing...")
            word = r.recognize_google(audio)

            print(word)
            if  word.lower() == "activate":

                print("Komal activate...")
                speak("hii honey")
                print("Komal is ready for command...")
                speak("Komal is ready for command")
                
                while True :
                    try:
                        with sr.Microphone(device_index=2) as source:

                            print(".........")
                            
                            r.adjust_for_ambient_noise(source, duration=1)
                            audio = r.listen(source, timeout=20, phrase_time_limit=7)

                            print("Command recognizing...")
                            command = r.recognize_google(audio)
                            print(command)

                            if command.lower() == "deactivate" :
                                speak("deactivating")
                                speak("see you later honey")
                                break

                            process_command(command)

                    except Exception as e:
                        print("Error:",e)

        except Exception as e:
            print("Error:", e)

        