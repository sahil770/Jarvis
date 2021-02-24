import subprocess
import wolframalpha
import pyttsx3
# import tkinter
import json
import random
# import operator
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
# import feedparser
# import smtplib
import ctypes
import time
import requests
import shutil
# from twilio.rest import Client
# from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import speech_recognition as sr

 

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[0].id)
engine.setProperty('voices', voices[0].id)



# speak function
def speak(audio):
    engine.say(audio) 
    engine.runAndWait() #Without this command, speech will not be audible to us.

# take commands
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

        try:
            print("Recognizing...")    
            query = r.recognize_google(audio, language='en-in') #Using google for voice recognition.
            print(f"User said: {query}\n")  #User query will be printed.

        except Exception as e:
            print(e)    
            print("Say that again please...")   #Say that again will be printed in case of improper voice 
            return "None" #None string will be returned
        return query

# wish me 
def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour>= 0 and hour<12:
        speak("Good Morning Sir !")
  
    elif hour>= 12 and hour<18:
        speak("Good Afternoon Sir !")   
  
    else:
        speak("Good Evening Sir !")  
        assname =("Jarvis 1 point o")
        speak("I am your Assistant"+assname)
        speak("please Tell me sir how would i help you.")



# main function 
if __name__=="__main__" :

    clear = lambda: os.system('cls')
    
    wishMe()

    while True:
        query  = takeCommand().lower()


        # applying conditions
        # how are you
        if 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")


        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Sahil.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        # iam fine 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
             
        #  bye command
        elif 'bye' in query:
            exit()

        # if wikipedia detected
        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2) 
            speak("According to Wikipedia")
            print(results)
            speak(results)

        # if open youtube is detected 
        elif 'open youtube' in query:
            speak('sure')
            webbrowser.open("https://www.youtube.com")

        # if open google is detected 
        elif 'open google' in query:
            speak('sure')
            webbrowser.open("https://www.google.com")

        elif 'open facebook' in query:
            speak('sure')
            webbrowser.open("https://www.facebook.com")

        elif 'open instagram' in query:
            speak('sure')
            webbrowser.open("https://www.instagram.com")

        elif 'open gmail' in query:
            speak('sure')
            webbrowser.open("https://www.gmail.com")

        elif 'open facebook message' in query or 'open facebook messages' in query:
            speak('sure')
            webbrowser.open("https://www.facebook.com/messages/t/.com")

        elif 'open twitter' in query:
            speak('sure')
            webbrowser.open("twitter.com")

        elif 'open google photos' in query:
            speak('sure')
            webbrowser.open("https://photos.google.com/?tab=mq&pageId=none")

        # stack overflow
        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com") 

        # time and date
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        # reply of nice

        elif 'thank you' in query:
            speak('anything for you sir')

        # reply of nice
        elif 'nice' in query:
            speak('thank you sir')

        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you sir")


        # most asked question from google Assistant
        elif "will you be my girlfriend" in query or "will you be my bf" in query:   
            speak("I'm not sure about, may be you should give me some time")
 
        elif "how are you" in query:
            speak("I'm fine, glad you me that")
 
        elif "i love you" in query:
            speak("It's hard to understand")


        # calculate
        elif "calculate" in query: 
            app_id = "XJ375G-KVJ7QJ8EVX"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer) 


        elif 'search' in query or 'play' in query:
            query = query.replace("search", "") 
            query = query.replace("play", "")          
            webbrowser.open(query)


        elif "who i am" in query:
            speak("If you talk then definately your human.")

        elif "why you came to world" in query:
            speak("Thanks to Sahil. further It's a secret")

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Sahil")

        elif 'reason for you' in query:
            speak("I was created as a Major project by Mister Sahil ")
        
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
            speak("Background changed succesfully")

        elif 'news' in query:
            try:
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                 
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                 
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
        
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
        
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")  
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is" in query or "who is" in query:

            client = wolframalpha.Client("XJ375G-KVJ7QJ8EVX")
            res = client.query(query)
             
            try:
                print (next(res.results).text)
                speak (next(res.results).text)
            except StopIteration:
                print ("No results")


        elif 'open chrome' in query:
            speak('sure')
            src = r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(src)

        elif 'open android stuio' in query:
            speak('sure')
            src = r"C:\\Program Files\\Android\Android Studio\bin\\studio64.exe"
            os.startfile(src)

        elif 'open any desk' in query:
            speak('sure')
            src = r"C:\\Program Files (x86)\\AnyDesk\AnyDesk.exe"
            os.startfile(src)

        elif 'open my project' in query:
            speak('sure')
            src = r"E:\projects\jarvis\jarvis"
            os.startfile(src)
        


        elif 'open code' in query or 'open visual code' in query or 'open code' in query or 'open visual studio code' in query or 'open editor' in query:
            speak('sure')
            appli = r"C:\\Users\\pc\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(appli)





        # close files

        elif 'close google chrome' in query:
            filename = r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            os.close(filename)
