"""
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import time
import pyttsx3
import speech_recognition as sr
import pytz
import subprocess
"""
import speech_recognition as sr
import os
#import pyttsx3
#from gtts import gTTS
import win32com.client as wincl
import datetime
import warnings
import calendar
import random
import wikipedia
import subprocess
import webbrowser
import smtplib
import requests
import re
import pyautogui
#from weather import Weather

#Ignore any warnings
warnings.filterwarnings('ignore')

class person:
    name = ''
    def setName(self, name):
        self.name = name

class asis:
    name = ''
    def setName(self, name):
        self.name = name


def recordAudio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Say something: ')
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
        
    data = ""

    try:
        data = r.recognize_google(audio)
        print(data)
    except Exception as e:
        print("Exception: " + str(e))
    return data

# Function to get the virtual assistant response
def assistantResponse(text):
    print(text)
    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak(text)

def wakeWord(text):
    WAKE_WORDS = ['hey computer','ok computer','hello']
    text = text.lower()
    #Check to see if users command have a wake word
    for phrase in WAKE_WORDS:
        if phrase in text:
            return phrase

def getDate():
    now = datetime.datetime.now()
    my_date = datetime.datetime.today()
    weekday = calendar.day_name[my_date.weekday()]
    monthNum = now.month
    dayNum = now.day
    month_names = ["january", "february", "march", "april", "may", "june","july", "august", "september","october","november", "december"]
    day_names = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    odinalNumbers = ['1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th','11th','12th','13th','14th','15th','16th','17th','18th','19th','20th','21st','22nd','13rd','24th','25th','26th','27th','28th','29th','30th','31st']

    return 'Today is ' + weekday +' '+ month_names[monthNum-1]+' '+ odinalNumbers[dayNum-1]+'.' 


def greeting(text):
    GREETING_INPUT = ['hi','hey','hola','greeting','wassup','hello']
    GREETING_RESPONCE = ['howdy','whats good','hello','hey there','hey ayush','hello ayush']
    for word in text.split():
        if word.lower() in GREETING_INPUT:
            return random.choice(GREETING_RESPONCE) + "."
    return ''

def getPerson(text):
    wordList = text.split()
    for i in range(0,len(wordList)):
        if i + 3 <= len(wordList) -1 and wordList[i].lower() == 'who' and wordList[i+1].lower() == 'is':
            return wordList[i+2] + ' '+ wordList[i+3]  

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])

person_obj = person()
asis_obj = asis()
asis_obj.name = 'kiki'
person_obj.name = 'ayush'


while True:
    # Record the audio
    text = recordAudio()
    response = '' #Empty response string
     
    # Checking for the wake word/phrase
    if ((wakeWord(text) == True)+asis_obj.name):
         # Check for greetings by the user
        response = response + greeting(text)
        # Check to see if the user said date
        # Check to see if the user said date
        if ('date' in text):
            get_date = getDate()
            response = response + ' ' + get_date
         # Check to see if the user said time
        if('time' in text):
            now = datetime.datetime.now()
            meridiem = ''
            if now.hour >= 12:
                meridiem = 'pm' #Post Meridiem (PM)
                hour = now.hour - 12
            else:
                meridiem = 'am'#Ante Meridiem (AM)
                hour = now.hour
           # Convert minute into a proper string
            if now.minute < 10:
                minute = '0'+str(now.minute)
            else:
                minute = str(now.minute)
            response = response + ' '+ 'It is '+ str(hour)+ ':'+minute+' '+meridiem+' .'
                
        # Check to see if the user said 'who is'
        if ('who is' in text):
            person = getPerson(text)
            wiki = wikipedia.summary(person, sentences=2)            
            response = response + ' ' + wiki

        NOTE_STRS = ["make a note", "write this down", "remember this"]
        for phrase in NOTE_STRS:
            if phrase in text:
                assistantResponse("What would you like me to write down?")
                note_text = recordAudio()
                note(note_text)
                assistantResponse("I've made a note of that.")

        BREAK_WORD = ['by computer','goodbye computer','good bye']
        BREAK_RESPONCE = ['have a nice day','goodbye sir','bye','have a good day sir','goodbye and take care']
        for phrase in BREAK_WORD:
            if phrase in text:
                assistantResponse(random.choice(BREAK_RESPONCE))
                exit()
        
        OPEN_BROW = ['open web browser','open chrome','open google in chrome','open google search in chrome','open google chrome']
        OPEN_RESPONCE = ['done','completed','task completed','Yes, i have open it']
        for phrase in OPEN_BROW:
            if phrase in text.lower():
                webbrowser.open('https://www.google.co.in/',new=2)
                assistantResponse(random.choice(OPEN_RESPONCE))

        SEND_EMAIL = ['write an email','please send an email','send an email']
        for phrase in SEND_EMAIL:
            if phrase in text.lower():
                assistantResponse('write down the recipient name')
                recipient =  input()
                server = smtplib.SMTP('64.233.184.108')
                server.ehlo()
                server.starttls()
                server.ehlo()
                #Next, log in to the server
                server.login("07ayushkasera@gmail.com","satyam8546091008")

                assistantResponse('What do want to text?')
                #Send the mail
                msg = "\n " + recordAudio() # The /n separates the message from the headers
                #fullmsg = msg.as_string()
                server.sendmail("07ayushkasera@gmail.com", recipient, msg)
                server.quit()
                assistantResponse(random.choice(OPEN_RESPONCE))
                break
        
        if 'play music' in text:
            music_dir = 'F:\\music'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[0]))
        
        SEARCH_INPUT = ['search for','google for']
        for phrase in SEARCH_INPUT:
            if phrase in text.lower():
                search_term = text.split('for',1)[-1]
                final_search_term = search_term.split('on',1)[0]
                url = 'https://www.google.com/search?q='+final_search_term
                webbrowser.get().open(url)
                assistantResponse('Here what i found for'+final_search_term+' on google')

        YOUTUBE_SEARCH = ['do a youtube search','search on youtube']
        YOUTUBE_SEARCH_QUESTION = ['what do you want to search?','what do you want to make me search?','what should i search for?']
        for phrase in YOUTUBE_SEARCH:
            if phrase in text.lower():
                assistantResponse('OK,'+random.choice(YOUTUBE_SEARCH_QUESTION))
                search_input = recordAudio()
                youtube_search_term = search_input.split('for',1)[-1]
                url = 'https://www.youtube.com/results?search_query='+youtube_search_term
                webbrowser.get().open(url)
                assistantResponse('Here what i found for'+youtube_search_term+' on youtube')

        CAL_INPUT = ['do a calculation','please, perform some calculation']
        for phrase in CAL_INPUT:
            if phrase in text.lower():
                assistantResponse('what do you want to calculate?')
                calculator = recordAudio()
                opr = calculator.split()[1]

                if opr == '+':
                    assistantResponse(int(calculator.split()[0]) + int(calculator.split()[2]))
                elif opr == '-':
                    assistantResponse(int(calculator.split()[0]) - int(calculator.split()[2]))
                elif opr == 'multiply':
                    assistantResponse(int(calculator.split()[0]) * int(calculator.split()[2]))
                elif opr == 'divide':
                    assistantResponse(int(calculator.split()[0]) / int(calculator.split()[2]))
                elif opr == 'power':
                    assistantResponse(int(calculator.split()[0]) ** int(calculator.split()[2]))
                else:
                    assistantResponse("Wrong Operator")

            
        SCREENSHOT_INPUT = ['capture','my screen','screenshot']
        for phrase in SCREENSHOT_INPUT:
            if phrase in text.lower():
                myScreenshot = pyautogui.screenshot()
                date = datetime.datetime.now()
                file_name = str(date).replace(":", "-") + "-screenshot.png"
                myScreenshot.save('C:\\Users\\07ayu\\Desktop\\Voice Assistant\\Voice-Assistent\\screenshots\\'+file_name)
                assistantResponse('Yeah! it\'s done. i took a screenshot')

        NAME_INPUT = ["what is your name","what's your name","tell me your name"]
        for phrase in NAME_INPUT:
            if phrase in text.lower():
                if asis_obj.name:
                    assistantResponse("My name is"+asis_obj.name)
                else:
                    assistantResponse("i dont know my name . what's your name?")

        MY_NAME_INPUT = ['my name is']
        for phrase in MY_NAME_INPUT:
            if phrase in text.lower():
                person_name = text.split("is")[-1].strip()
                assistantResponse("okay, i will remember that " + person_name)
                person_obj.setName(person_name) # remember name in person object
        
        ASSIS_NAME_INPUT = ['your name should be']
        for phrase in ASSIS_NAME_INPUT:
            if phrase in text.lower():
                asis_name = text.split("be")[-1].strip()
                assistantResponse("okay, i will remember that my name is " + asis_name)
                asis_obj.setName(asis_name)
        
       # Assistant Audio Response
        assistantResponse(response)








"""
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january", "february", "march", "april", "may", "june","july", "august", "september","october","november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENTIONS = ["rd", "th", "st", "nd"]

def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        said = ""

        try:
            said = r.recognize_google(audio)
            print(said)
        except Exception as e:
            print("Exception: " + str(e))

    return said.lower()


def authenticate_google():

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day, service):
    # Call the Calendar API
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("-")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12) + start_time.split(":")[1]
                start_time = start_time + "pm"

            speak(event["summary"] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENTIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    # THE NEW PART STARTS HERE
    if month < today.month and month != -1:  # if the month mentioned is before the current month set the year to the next
        year = year+1

    # This is slighlty different from the video but the correct version
    if month == -1 and day != -1:  # if we didn't find a month, but we have a day
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month

    # if we only found a dta of the week
    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)

    if day != -1:  # FIXED FROM VIDEO
        return datetime.date(month=month, day=day, year=year)

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])



##I will fix it latter there are some issues
WAKE = "hello"
SERVICE = authenticate_google()
print("Start")

while True:
    print("Listening")
    text = get_audio()

    if text.count(WAKE) > 0:
        speak("I am ready")
        text = get_audio()

        CALENDAR_STRS = ["what do i have", "do i have plans", "am i busy"]
        for phrase in CALENDAR_STRS:
            if phrase in text:
                date = get_date(text)
                if date:
                    get_events(date, SERVICE)
                else:
                    speak("I don't understand")

        NOTE_STRS = ["make a note", "write this down", "remember this"]
        for phrase in NOTE_STRS:
            if phrase in text:
                speak("What would you like me to write down?")
                note_text = get_audio()
                note(note_text)
                speak("I've made a note of that.")
"""
