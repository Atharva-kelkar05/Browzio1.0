from sqlite3 import Time
import subprocess
import bs4
import cv2
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from translate import Translator
import language_tool_python

# we are making a engine to recognise the speech and convert to text and then it will be
# stored in a variable named as query.
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

# defining a source and making it to speak using engine.
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
 
# defining a function to make our assistant wish us when we run the program.
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning Sir !")
  
    elif hour>= 12 and hour<18:
        speak("Good Afternoon Sir !")  
  
    else:
        speak("Good Evening Sir !") 
    speak("I am your Assistant, Browzio!")

def getCityName(query_string):
    words = query_string.split(" ")
    for word in range(0,len(words)):
        if words[word] == 'in' or words[word] == 'at':
            return words[word+1]     

def username():
     speak("How can i Help you Sir?")
 
def takeCommand():
     
    r = sr.Recognizer()
     
    with sr.Microphone() as source:
         
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
  
    try:
        print("Recognizing...")   
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")
  
    except Exception as e:
        print(e)   
        print("Unable to Recognize your voice.") 
        return "None"
     
    return query
  
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
     
    # Enable low security in gmail
    server.login('atharva.kelkar@vitbhopal.ac.in', 'atharvak@123')
    server.sendmail('atharva.kelkar@vitbhopal.ac.in', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
     
    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    username()
     
    while True:
         
        query = takeCommand().lower()
         
        # All the commands said by user will be
        # stored here in 'query' and will be
        # converted to lower case for easy
        # recognition of command
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)
 
        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")
 
        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")
 
        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow. Happy coding")
            webbrowser.open("stackoverflow.com")  
 
        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "D:\Personal folder\music"
            songs = os.listdir(music_dir)
            print(songs)   
            random = os.startfile(os.path.join(music_dir, random.choice(songs)))

        elif "stop music" in query:
            subprocess.Popen(["taskkill","/IM","Music.UI.exe","/F"],shell=True)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")   
            speak("Sir, the time is "+ strTime)
 
        elif 'email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "Receiver email address"   
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")
 
        elif 'mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()   
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")
            #general Questions:
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")
 
        elif 'fine' in query or "am good" in query:
            speak("It's good to know that your fine")
 
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
 
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")
 
        elif "what is your name" in query or "What is your name" in query:
            speak("My friends call me ")
            speak("Browzio")
            print("My friends call me Browzio.")
 
        elif 'exit' in query or 'bye' in query or 'goodbye' in query.lower():
            speak("Good Bye! See you later.")
            exit()
 
        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Python. Students from VIT-B created me I don't know the names but I am thankful to them.")
             
        elif 'joke' in query:
            speak(pyjokes.get_joke())
             
        elif 'add' in query:
            t = query.split('add')[-1]
            k = t.split('and')
            sum = 0
            for i in range(len(k)):
                sum += int(k[i])
                print(sum)
                speak(sum)
 
        elif 'search' in query:
             
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open(query)
 
        elif "who i am" in query:
            speak("If you talk then definitely you are human.")
 
        elif "why you came to world" in query:
            speak("Humans made me. I don't know.")
 
        elif 'powerpoint presentation' in query.lower() or 'ppt' in query.lower():
            speak("opening Power Point presentation")
            power = r"C:\\Users\\Voice Assistant.pptx"
            os.startfile(power)
 
        elif 'what is love' in query:
            speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            speak("I am your virtual assistant, Browzio")
 
        elif 'reason for you' in query:
            speak("I was created as a Minor project using")
 
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
            speak("Background changed successfully")

 
        elif 'news' in query:
             
            try:
                jsonObj = urlopen('''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=f856176bfd8d4cea93d45c77cf18ba82''')
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
                 
        elif 'empty recycle bin' in query or 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Emptied")
 
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop Bhaavna from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
 
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")
 
        elif "camera" in query or "take a photo" in query:
            video_capture_object = cv2.VideoCapture(0)
            x = str(datetime.datetime.now())
            print(x)
            result = True
            while result:
                ret, frame = video_capture_object.read()
                cv2.imwrite("x.jpg", frame)
                result = False
            video_capture_object.release()
            cv2.destroyAllWindows()
            os.system("x.jpg")
 
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown", "/f","/h")
 
        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
 
        elif "write a note" in query or "take a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('Browzio.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
        elif "show note" in query:
            speak("Showing Notes")
            file = open("Readme.txt", "r")
            print(file.read())
            speak(file.read(6))
 
        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)
             
            with open("Voice.py", "wb") as Pypdf:
                 
                total_length = int(r.headers.get('content-length'))
                 
                for ch in progress.bar(r.iter_content(chunk_size = 2391975),
                                       expected_size =(total_length / 1024) + 1):
                    if ch:
                      Pypdf.write(ch)
                     
        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "Assistant" in query or "assistant" in query:
             
            wishMe()
            speak("Browzio! in your service Mister")
            speak(username)

        elif 'weather' in query or 'temperature' in query:
            # query = query.replace("how is the weather ", "")
            query = getCityName(query)

            location = query
            speak("You asked for weather " + location)
            url = "https://google.com/search?q=weather+in+" + location
            request_result = requests.get( url )
            soup = bs4.BeautifulSoup( request_result.text , "html.parser" )
            temp = soup.find( "div" , class_='BNeawe' ).text
            x= ('Temperature in '+location +' is  ' + temp)
            print(x)
            speak(x)
            
        elif "send message " in query:
                # You need to create an account on Twilio to use this service
                account_sid = 'Account Sid key'
                auth_token = 'Auth token'
                client = Client(account_sid, auth_token)
 
                message = client.messages \
                    .create(body = takeCommand(),from_='Sender No',to ='Receiver No')
                print(message.sid)
 
        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")
 
        elif "Good Morning" in query or "good morning" in query or "good afternoon" in query or "good evening" in query:
            speak("A warm" +query)
            speak("How are you Mister")
            speak("AK")

##############Translate from some language to some other.############################
        
        elif "Translate" in query or "translate" in query:
            #speak("In which language should I translate?")
            lang=query.split("to") [1]
            print(lang)
            #lang=input("Say your language:")
            if "english" in query:
                speak("what should I translate?")
                Translator=Translator(to_lang="english")
                sentence=takeCommand().lower()
                translation=Translator.translate(sentence)
                print(translation)
                speak(translation)
            elif "hindi" in query:
                speak("what should I translate?")
                Translator=Translator(to_lang="hindi")
                sentence=takeCommand().lower()
                translation=Translator.translate(sentence)
                print(translation)
                speak(translation)
            elif "marathi" in query:
                speak("what should I translate?")
                Translator=Translator(to_lang="marathi")
                sentence=takeCommand().lower()
                translation=Translator.translate(sentence)
                print(translation)
                speak(translation)
            elif "spanish" in query:
                speak("what should I translate?")
                Translator=Translator(to_lang="spanish")
                sentence=takeCommand().lower()
                translation=Translator.translate(sentence)
                print(translation)
                speak(translation)
            elif "german" in query:
                speak("what should I translate?")
                Translator=Translator(to_lang="german")
                sentence=takeCommand().lower()
                translation=Translator.translate(sentence)
                print(translation)
                speak(translation)
            elif "french" in query:
                speak("what should I translate?")
                Translator=Translator(to_lang="french")
                sentence=takeCommand().lower()
                translation=Translator.translate(sentence)
                print(translation)
                speak(translation)
            else:
                pass
        
#################################################################################### 
        # most asked question from Google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:  
            speak("I'm not sure about, may be you should give me some time")
 
        elif "how are you" in query:
            speak("I'm fine, glad you me that")
 
        elif "i love you" in query:
            speak("It's hard to understand")
 
        elif "what is" in query or "who is" in query:
             
            # Use the same API key
            # that we have generated earlier
            client = wolframalpha.Client("API_ID")
            res = client.query(query)
             
            try:
                print (next(res.results).text)
                speak (next(res.results).text)
            except StopIteration:
                print ("No results")
 
        # elif "" in query:
            # Command go here
            # For adding more commands


