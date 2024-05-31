import win32com.client
import os
import speech_recognition as sr
import webbrowser
import datetime
import smtplib
from ctypes import cast,POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities,IAudioEndpointVolume
import playsound
import time
import requests
import pyautogui
import psutil
import subprocess

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
volume = cast(interface,POINTER(IAudioEndpointVolume))
def volume_ten():
    volume.SetMasterVolumeLevel(-10.0,None)
    # playsound.playsound("sunflower-street-drumloop-85bpm-163900.mp3")
def volume_twenty():
    volume.SetMasterVolumeLevel(-20.0,None)
    # playsound.playsound("sunflower-street-drumloop-85bpm-163900.mp3")
def volume_thirty():
    volume.SetMasterVolumeLevel(-30.0,None)
    # playsound.playsound("sunflower-street-drumloop-85bpm-163900.mp3")

def get_weather_data(city):
    API_key = "0c06b84562efff52297a715069dba2ff"
    url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_key}&units=metric"

    response = requests.get(url).json()
    if 'main' in response:
        weather_data = {
            'city': city,
            'temperature': response['main'].get('temp'),
            'wind': response['wind'].get('speed'),
            'humidity': response['main'].get('humidity'),
            # Converting the current pc time to current indian standard time
            # 'date': datetime.utcnow() + timedelta(hours=5, minutes=30),
            'visibility': response['visibility'],
            'feels_like': response['weather'][0]['main'],
            'description': response['weather'][0]['description'],
        }
        weather_details = ''
        return weather_details + ("The weather in {} is currently {} with a temperature of {} degrees and wind speeds reaching {} kilometres per hour".format(weather_data['city'], weather_data['feels_like'], weather_data['temperature'], weather_data['wind']))
    else:
        return None


def say(s):
    # while 1:
    speaker.Speak(s)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 1
        audio = r.listen(source)
        try :
            print("Recognizing...")
            query = r.recognize_google(audio,language="en-in")
            # query = r.recognize_google(audio,language="hi-in")
            print(f"User said : {query}")
            return query
        except Exception as e:
            return "Some Error Occured . Sorry from Alexa"

if __name__ =='__main__':
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak("Alexa Playing...,How may I help you")
    while True:
        # print("Enter text" )
        # s=input()
        # say(s)
        print("Listening...")
        text = takeCommand()
        print(text)
        say(text)

        sites = [["youtube","https://www.youtube.com"],["wikipedia","https://www.wikipedia.com"],["google" ,"https://www.google.com"],["spotify" ,"https://www.spotify.com"],["spotify song","https://open.spotify.com/track/06uNwpGEx0XElATDWFwsxZ"]]
        for site in sites :
            if f"Open {site[0]}".lower() in text.lower():
                say(f"Opening {site[0]} DDT...")
                webbrowser.open(site[1])

        apps = [["chrome","C:\Program Files\Google\Chrome\Application\chrome.exe"],["Packet Tracer","C:/Program Files/Cisco Packet Tracer 8.2.1/bin/PacketTracer.exe"],["telegram" ,"C:/Users/BAPS/AppData/Roaming/Telegram Desktop/Telegram.exe"],["virtual box","C:/Program Files/Oracle/VirtualBox"],["whatsapp","C:\\Program Files\\WindowsApps\\5319275A.WhatsAppDesktop_2.2409.8.0_x64__cv1g1gvanyjgm\\WhatsApp.exe"]]

        # C:/Program Files/WindowsApps/5319275A.WhatsAppDesktop_2.2401.4.0_x64__cv1g1gvanyjgm/WhatsApp.exe
        for app in apps :
            if f"open {app[0]}".lower() in text.lower():
                say(f"Opening {app[0]} DDT...")
                webbrowser.open(app[1])
        if "the time" in text :
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Tanvi the time is {strfTime}")
        if "open music" in text :
            musicpath = "https://open.spotify.com/track/06uNwpGEx0XElATDWFwsxZ"
            # os.system(f"open {musicpath}")
            webbrowser.open(musicpath)
        if "send email" in text :
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login('priyanshisojitra1712@gmail.com', 'bfsx otlx bpdb hluh')
            send = server.sendmail('priyanshisojitra1712@gmail.com', 'kashishsavaliya10@gmail.com', 'Hii Tanvi , Greetings from Dipti...')
            say(send)
        if "decrease volume by 10" in text :
            volume_ten()
        if "decrease volume by 20" in text :
            volume_twenty()
        if "decrease volume by 30" in text :
            volume_thirty()
        if "weather" in text :
            # city = input("City name : ")
            lis = list(text.split(" "))
            length = len(lis)
            city = lis[length-1]
            weather_details = get_weather_data(city)
            print(weather_details)
            say(weather_details)
        if "where is " in text:
            order = text.replace("where is","")
            location = order
            say("Locating......")
            say(location)
            webbrowser.open("https://www.google.co.in/maps/place/"+location+"")

        if "switch window" in text:
            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.keyUp('alt')

        if "take a screenshot" in text or "screenshot this" in text:
            say("Please tell me the name for this file.")
            name = takeCommand().lower()
            say("Please hold the screen.")
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f"{name}.png")
            say("Screenshot captured!")

        if "cpu status" in text.lower():
            usage = str(psutil.cpu_percent())
            say("CPU is at " + usage)
            battery = str(psutil.sensors_battery())
            say("CPU is at " + battery)

        if "shutdown" in text or "turn off" in text:
            say("Hold on a second.Your system is on its way to shutdown")
            say("Make sure all of your applications are closed")
            time.sleep(5)
            subprocess.call(['shutdown', '/s'])

        elif "restart" in text:
            say("Restarting....")
            subprocess.call(['shutdown', '/r'])

        elif "hibernate" in text:
            say("Hibernating....")
            subprocess.call(['shutdown', '/h'])

