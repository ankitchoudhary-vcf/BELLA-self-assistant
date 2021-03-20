import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import smtplib
import pywhatkit
import pyjokes
import googletrans
import smtplib
import random
import wolframalpha
import twilio
from twilio.rest import Client
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import json
import winshell
import time
import requests
import PyQt5
import pyautogui
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from projectsui import Ui_MainWindow
import sys
import subprocess


gt = googletrans.Translator()

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 170)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!") 


    strdate = datetime.datetime.now().strftime("%Y-%m-%d")
    speak(f"today date is {strdate}")    

    strTime = datetime.datetime.now().strftime("%H:%M:%S")    
    speak(f"and current time is {strTime}") 

        

    speak("i am your personal assistant Bella. Please tell me how may I help you") 

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('sg247938@gmail.com', 'Shivam1319@')
    server.sendmail('sg247938@gmail.com', to, content)
    server.close()






class MainThread(QtCore.QThread):
    def __init__(self):
        super(MainThread,self).__init__()

    def run(self):
        self.run_Bella()   

    def takeCommand(self):
        #It takes microphone input from the user and returns string output

        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            audio = r.listen(source)
            

        try:
            print("Recognizing...")    
            self.query = r.recognize_google(audio, language='en-in')
            self.query = self.query.lower()
            if 'Bella' in self.query:
                self.query = self.query.replace('Bella','')
            print(self.query)

        except Exception as e:
            # print(e)    
            print("Say that again please...")
            speak("Say that again please...")  
            return "None"
        return self.query
        
    def run_Bella(self):
        wishMe()
        while True:
        # if 1:
      
            self.query = self.takeCommand().lower()

            # Logic for executing tasks based on self.query
            if 'wikipedia' in self.query:
                speak('Searching Wikipedia...')
                self.query = self.query.replace("wikipedia", "")
                results = wikipedia.summary(self.query, sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)

            elif 'open youtube' in self.query:
                webbrowser.open("youtube.com")

            elif 'open google' in self.query:
                webbrowser.open("google.com")

            elif 'open stack overflow' in self.query:
                webbrowser.open("stackoverflow.com")   

            elif 'play' in self.query:
                song = self.query.replace('play', '')
                speak('playing ' + song)
                pywhatkit.playonyt(song)
            
            elif 'joke' in self.query:
                speak(pyjokes.get_joke())    

            elif 'the time' in self.query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")    
                speak(f"Sir, the time is {strTime}")

            elif 'open code' in self.query:
                codePath = " C:/Users/dell/AppData/Local/Programs/Python/Python39/python.exe"
                os.startfile(codePath)

            elif 'open facebook' in self.query:
                webbrowser.open("https://www.facebook.com")
                speak("opening facebook") 

            elif 'open instagram' in self.query:
                webbrowser.open("https://www.instagram.com")
                speak("opening instagram")   

            elif 'open yahoo' in self.query:
                webbrowser.open("https://www.yahoo.com")
                speak("opening yahoo")
                
            elif 'open gmail' in self.query:
                webbrowser.open("https://mail.google.com")
                speak("opening google mail") 
                
            elif 'open snapdeal' in self.query:
                webbrowser.open("https://www.snapdeal.com") 
                speak("opening snapdeal")  
                
            elif 'open amazon' in self.query or 'shop online' in self.query:
                webbrowser.open("https://www.amazon.com")
                speak("opening amazon")

            elif 'open flipkart' in self.query:
                webbrowser.open("https://www.flipkart.com")
                speak("opening flipkart")  

            elif 'open ebay' in self.query:
                webbrowser.open("https://www.ebay.com")
                speak("opening ebay")     

            elif 'mail to sister' in self.query:
                try:
                    speak("What should I say?")
                    content = self.takeCommand()
                    to = "gsheetal774@gmail.com"    
                    sendEmail(to, content)
                    speak("Email has been sent!")
                except Exception as e:
                    print(e)
                    speak("Sorry sir. I am not able to send this email")  

            elif 'send a mail' in self.query:
                try:
                    speak("What should I say?")
                    content = self.takeCommand()
                    speak("whome should i send")
                    to = input()    
                    sendEmail(to, content)
                    speak("Email has been sent !")
                except Exception as e:
                    print(e)
                    speak("I am not able to send this email")

            elif "change name" in self.query:
                speak("What would you like to call me, Sir ")
                assname = self.takeCommand()
                speak("Thanks for naming me")
    
            elif "what's your name" in self.query or "What is your name" in self.query:
                speak("My friends call me")
                speak(assname)
                print("My friends call me", assname)                


            elif 'good bye' in self.query:
                speak("good bye")
                exit()

            elif "shutdown" in self.query:
                speak("shutting down")
                os.system('shutdown -s') 

            elif "what\'s up" in self.query or 'how are you' in self.query:
                stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy','i am okey ! How are you']
                ans_q = random.choice(stMsgs)
                speak(ans_q)  
                ans_take_from_user_how_are_you = self.takeCommand()
                if 'fine' in ans_take_from_user_how_are_you or 'happy' in ans_take_from_user_how_are_you or 'okey' in ans_take_from_user_how_are_you:
                    speak('okey..')  
                elif 'not' in ans_take_from_user_how_are_you or 'sad' in ans_take_from_user_how_are_you or 'upset' in ans_take_from_user_how_are_you:
                    speak('oh sorry..') 

            elif 'make you' in self.query or 'created you' in self.query or 'develop you' in self.query:
                ans_m = " For your information shivam gupta Created me ! I give Lot of Thannks to Him "
                print(ans_m)
                speak(ans_m)

            elif "who are you" in self.query or "about you" in self.query or "your details" in self.query:
                about = "I am Bella an A I based computer program but i can help you lot like a your close friend ! i promise you ! Simple try me to give simple command ! like playing music or video from your directory i also play video and song from web or online ! i can also entain you i so think you Understand me ! ok Lets Start "
                print(about)
                speak(about)

            elif "hello" in self.query or "hello Bella" in self.query:
                hel = "Hello Sir ! How May i Help you.."
                print(hel)
                speak(hel)

            elif "your name" in self.query or "sweat name" in self.query:
                na_me = "Thanks for Asking my name my self ! Bella"  
                print(na_me)
                speak(na_me)

            elif "you feeling" in self.query:
                print("feeling Very sweet after meeting with you")
                speak("feeling Very sweet after meeting with you") 
            

            elif 'exit' in self.query or 'abort' in self.query or 'stop' in self.query or 'bye' in self.query or 'quit' in self.query :
                ex_exit = 'I feeling very sweet after meeting with you but you are going! i am very sad'
                speak(ex_exit)
                exit()   

            elif "calculate" in self.query: 
                
                app_id = "UA6V6W-X2PXG24J98"
                client = wolframalpha.Client("UA6V6W-X2PXG24J98")
                indx = self.query.lower().split().index('calculate') 
                self.query = self.query.split()[indx + 1:] 
                res = client.query(' '.join(self.query)) 
                answer = next(res.results).text
                print("The answer is " + answer) 
                speak("The answer is " + answer)    
            
            

            elif 'news' in self.query:
                
                try: 
                    jsonObj = urlopen('''https://newsapi.org/v2/top-headlines?country=in&apiKey=9655f3abc604484586897775536c8622''')
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

            elif 'lock window' in self.query:
                    speak("locBella the device")
                    subprocess.call("shutdown / h")
    
            elif 'shutdown system' in self.query:
                    speak("Hold On a Sec ! Your system is on its way to shut down")
                    os.system('shutdown -s')
                    
            elif 'empty recycle bin' in self.query:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin Recycled")
    
            elif "don't listen" in self.query or "stop listening" in self.query:
                speak("for how much time you want to stop jarvis from listening commands")
                a = int(self.takeCommand())
                time.sleep(a)
                print(a)
    
            elif "where is" in self.query:
                self.query = self.query.replace("where is", "")
                location = self.query
                speak("User asked to Locate")
                speak(location)
                webbrowser.open("https://www.google.nl / maps / place/" + location + "")

            elif "restart" in self.query:
                subprocess.call(["shutdown", "/r"])
                
            elif "hibernate" in self.query or "sleep" in self.query:
                speak("Hibernating")
                subprocess.call("shutdown / h")
    
            elif "log off" in self.query or "sign out" in self.query:
                speak("Make sure all the application are closed before sign-out")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])
    
            elif "write a note" in self.query:
                speak("What should i write, sir")
                note = self.takeCommand()
                file = open('jarvis.txt', 'w')
                speak("Sir, Should i include date and time")
                snfm = self.takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("% H:% M:% S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note) 

            elif "show note" in self.query:
                speak("Showing Notes")
                file = open("jarvis.txt", "r") 
                print(file.read())
                speak(file.read(6))
    
                       

            elif "send message" in self.query:
                # You need to create an account on Twilio to use this service
                account_sid = 'AC82d4360947707beb34f155ae1d19d31b'
                auth_token = '43b4aa339a59e50e115fd8954580dae3'
                client = Client(account_sid, auth_token)

                message = client.messages \
                .create(
                    body="Join Earth's mightiest heroes. Like Kevin Bacon.",
                    from_='+18707298456',
                    to='7014217098'
                )

                print(message.sid)
                speak("message sent successfully")

            elif "what is" in self.query or "who is" in self.query:
                
                # Use the same API key 
                # that we have generated earlier
                client = wolframalpha.Client("UA6V6W-X2PXG24J98")
                res = client.query(self.query)
                
                try:
                    print (next(res.results).text)
                    speak (next(res.results).text)
                except StopIteration:
                    print ("No results") 

            elif 'open command' in self.query:
                codePath = "C:/WINDOWS/system32/cmd.exe"
                os.startfile(codePath) 

            elif 'close command' in self.query:
                os.system("taskkill /IM cmd.exe")    

            elif 'open notepad' in self.query:
                codePath = "C:\WINDOWS\system32/notepad.exe"
                os.startfile(codePath) 

            elif 'close notepad' in self.query:
                os.system("taskkill /IM notepad.exe")   

            elif 'open chrome' in self.query:
                codePath = "C:\Program Files (x86)\Google\Chrome\Application/chrome.exe"
                os.startfile(codePath)

            elif 'close chrome' in self.query:
                os.system("taskkill /IM chrome.exe")    

            elif 'open word' in self.query:
                codePath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
                os.startfile(codePath)

            elif 'open excel' in self.query:
                codePath = "C:\Program Files\Microsoft Office\root\Office16\excel.exe"
                os.startfile(codePath)

            elif 'open skype' in self.query:
                codePath = "skype.exe"
                os.startfile(codePath)

            elif 'open power point' in self.query:
                codePath = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.exe"
                os.startfile(codePath)

            elif 'open onenote' in self.query:
                codePath = "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.exe"
                os.startfile(codePath)

            elif 'take screenshot' in self.query or 'take a screenshot' in self.query:
                speak("sir,please tell me the name of this screenshot")
                name = self.takeCommand().lower()
                speak("sir,please hold the screen ,i am taBella a screenshot")
                time.sleep(3)
                img = pyautogui.screenshot()
                img.save(f"{name}.png")
                speak("i am done sir, screenshot is saved ")
            
            elif 'thank you ' in self.query or 'thanks' in self.query:
                speak("welcome sir")   

            elif "weather" in self.query:
             
                # Google Open weather website
                # to get API of Open weather 
                api_key = "6eee755a887428d041cdef4103794fac"
                base_url = "http://api.openweapip thermap.org / data / 2.5 / weather?"
                # speak(" City name ")
                # print("City name : ")
                # city_name = self.takeCommand()
                complete_url = 'api.openweathermap.org/data/2.5/weather?q={jaipur}&appid={6eee755a887428d041cdef4103794fac}'

                response = requests.get(complete_url) 
                x = response.json() 
                
                if x["cod"] != "404": 
                    y = x["main"] 
                    current_temperature = y["temp"] 
                    current_pressure = y["pressure"] 
                    current_humidiy = y["humidity"] 
                    z = x["weather"] 
                    weather_description = z[0]["description"] 
                    print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
                
                else: 
                    speak(" City Not Found ")

            elif 'location' in self.query:
                speak('What is the location?')
                location = self.takeCommand()
                url = 'https://google.nl/maps/place/' + location + '/&amp;'
                webbrowser.get('chrome').open_new_tab(url)
                speak('Here is the location ' + location)        




                                    
startexecution = MainThread()       



class Main(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.starttask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def showtime(self):
        # current_time = QTime.currentTime()
        # current_date = QDate.currentDate()
        label_time = QtCore.QTime.currentTime.tostring("%H:%M:%S")
        label_date = QtCore.QDate.currentDate.tostring(QtCore.Qt.ISODate)
        label_show = "Created By Shivam"
        self.ui.textBrowser.setText(label_time)
        print(label_date)
        self.ui.textBrowser_2.setText(label_date)
        print(label_time)
        self.ui.textBrowser_3.setText(label_show)    


    def starttask(self):
        
        self.ui.movie = QtGui.QMovie("C:/Users/dell/Documents/blueborder.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("C:/Users/dell/Documents/tumblr_ouo8yb59lg1s32c21o1_r2_500.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("C:/Users/dell/Documents/1_e_Loq49BI4WmN7o9ItTADg.gif")
        self.ui.label_7.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("C:/Users/dell/Documents/113cbe7d34450c483a70d182d70b3669.gif")
        self.ui.label_8.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("C:/Users/dell/Documents/0a8671a21422eecab8189a2941bfb132.gif")
        self.ui.label_9.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("C:/Users/dell/Documents/T8bahf.gif")
        self.ui.label_10.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.showtime)
        self.timer = QtCore.QTime()
        self.timer.start()
        startexecution.start()
        
    




app = QtWidgets.QApplication(sys.argv)
project = Main()
project.show()
exit(app.exec_())







        

     
