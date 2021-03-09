from multiprocessing import Process
from win32com.client import Dispatch
import speech_recognition as sr 
import pyttsx3  
import webbrowser
import os 
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium import webdriver 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import urllib.request
import re
from datetime import datetime
import datetime
from datetime import datetime
import pytz 
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pynput.keyboard import Key, Controller
import keyboard
import sys
import winsound
import time
import psutil
import pyautogui
import shutil
import json

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')


engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    import datetime
    hour = int(datetime.datetime.now().hour)
    
    time_2 = time.strftime("%I : %M")
    if "0" in time_2[0]:
        new_time_2 = time_2.replace(time_2[0], "")
        am = "sir time is" + new_time_2 + "a m"
        pm = "sir time is" + new_time_2 + "p m"
    else:
        am =  "sir time is" + time_2 + "a m"
        pm = "sir time is" + time_2 + "p m"
    if hour>=0 and hour<12:
        print("Good Morning Sir  ")
        speak("Good Morning Sir ")
        speak(am)
        weather()

    elif hour>=12 and hour<18:
        print("Good Afternoon Sir")
        speak("Good Afternoon Sir") 
        speak(pm)
        weather()
       
    else:
        print("Good Evening Sir")
        speak("Good Evening Sir")
        speak(pm)  
        weather()

    speak("sir your last updated list says")
    file = open('geek.txt', 'r') 
                # This will print every line one by one in the file 
    for each in file: 
        print (each) 
    speak(each)
    speak("how can i help you sir")
    winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
    winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)

def weather():
    # import required modules 
    import requests, json 

    # Enter your API key here 
    api_key = "Enter your API key here "

    # base_url variable to store url 
    base_url = "url"

    
    complete_url = base_url + "appid=" + api_key + "&q=" + "Rewari,Haryana" 


    response = requests.get(complete_url) 


    x = response.json() 


    if x["cod"] != "404": 

        
        y = x["main"] 

        # store the value corresponding 
        # to the "temp" key of y 
        current_temperature = y["temp"] 

        z = x["weather"] 

        
        weather_description = z[0]["description"] 

        # print following values 
        print(" Temperature (in kelvin unit) = " +str(current_temperature-273.15)) 
        c = current_temperature-273.15
        c_new = round(c,2)
        var = str(c_new)
        var_2 = str(weather_description)
        temp_weather = "sir temprature is " + var + "degree celsius" +  "and weather is " + var_2
        
        speak(temp_weather)
    else: 
        print(" City Not Found ") 

def google():
    next_command = takeCommand().lower()
        # new = next_command.strip("open")
        # def remove(string): 
        # 	return string.replace(" ", "") 
        # string = new
        # x= remove(string) 
        # print(x)
            
    try:
        from googlesearch import search
    except ImportError:
        print("No module named 'google' found")
    speak("opening"+next_command)
    
    for j in search(next_command, tld="com", num=3, start = 1, stop=3, pause=2): 
        
        webbrowser.open(j)

def for_files():
    try:
        next_command = takeCommand().lower()
        drive_name = next_command[-1] + ":\\ "
        file_name = next_command.split()[0]
        file_name_2 = "\\" +  file_name

        p = os.path.join(drive_name,file_name_2)
        speak("opening file")
        os.startfile(p)
    
    except Exception as e:
        print(e)
        speak("sir please say it again")
 
def calculator():
    next_command = takeCommand().lower()
    e = next_command.replace(" ","")
    equation = e
    print(e)
    if("divide" in equation):
        replace_it_for_div = equation.replace("divide","/")
        f = eval(replace_it_for_div)
        print(f)
        c = "its" , f
        print(c)
        speak(c)
    elif("multiply" in equation or "X" in equation):
        replace_it_for_mul = equation.replace("multiply","*")
        replace_it_for_mul_2 = replace_it_for_mul.replace("X","*")
        f = eval(replace_it_for_mul_2 )
        print(f)
        c = "its" , f
        print(c)
        speak(c)
    else:
        f = eval(equation)
        print(f)
        c = "its" , f
        speak(c)

def take_screenshot():
    next_command = takeCommand().lower()
    myScreenshot = pyautogui.screenshot()
    file_name = next_command
    path = "C:\\Users\\Samyak\\Pictures\\Screenshots\\" + file_name + ".png"
    myScreenshot.save(path)
    speak("done sit")

def typing():
    next_command = takeCommand().lower()
    import keyboard
    keyboard.press_and_release('space')
    keyboard.write(next_command)
    speak("data typed sir")
    winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
    winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)

def takeCommand():
    
    r = sr.Recognizer()
    with sr.Microphone() as source2: 
	
        r.adjust_for_ambient_noise(source2, duration=1) 

        audio = r.listen(source2) 

        mytext = r.recognize_google(audio)

        mytext = mytext.lower() 

    return mytext

def working():
    if __name__ == '__main__':
            wishMe()
            while(1):
                try:
                    print("speak up")
                    print("Listening.......")
                    
                    mytext = takeCommand().lower()
                   
                    
                    if( "take a break" in mytext):
                            speak("ok sir, bye bye")
                            print("ok sir bye! bye")
                            break
                    elif("search" and "file" in mytext):
                        speak("which file")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        #this code clear the string in mytext

                        otherText=mytext
                        new = otherText.replace(mytext,"")
                        mytext = new
                        for_files()
                        
                    elif("do a google" in mytext or "search on google" in mytext or "open google" in mytext):
                        speak("what you want to search")
                        try:
                            winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                            winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                            
                            otherText=mytext
                            new = otherText.replace(mytext,"")
                            mytext = new
                            google()
                        except Exception as e:
                            speak("sir please say it again")
                    elif("turn off" in mytext):
                        speak("sir you wanna set timer or want to do it right now")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        next_command = takeCommand().lower()    
                        if("set timer" in next_command):
                            x = 0
                            for i in next_command:
                                if("hour" in next_command):
                                    if (i>= '0' and i<='9'):
                                        x=x*10+int(i)
                                        x_2 = x*3600
                            new_time_for_sleep = x_2
                            print(new_time_for_sleep)

                            time.sleep(new_time_for_sleep)
                            speak("ok turning  down")
                            os.system("shutdown /s /t 1")
                        elif("right now" in next_command):
                            speak("ok turning  down")
                            os.system("shutdown /s /t 1")
                        elif("cancel" in next_command):
                            speak("ok sir request turned down")
                        else:
                            speak("cant understand that command please try again")
                                
                    elif("sleep" in mytext):
                        speak("sir you wanna set timer or want to do it right now")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        next_command = takeCommand().lower()
                            
                        if("set timer" in next_command):
                            x = 0
                            for i in next_command:
                                if("hour" in next_command):
                                    if (i>= '0' and i<='9'):
                                        x=x*10+int(i)
                                        x_2 = x*3600
                            new_time_for_sleep = x_2
                            print(new_time_for_sleep)

                            time.sleep(new_time_for_sleep)
                            speak("ok putting pc on sleep")
                            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
                        elif("right now" in next_command):
                            speak("ok sir putting pc on sleep")
                            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
                        elif("cancel" in next_command):
                            speak("ok sir request turned down")
                        else:
                            speak("cant understand that command please try again")	

                    elif("restart" in mytext):
                        speak("sir you wanna set timer or want to do it right now")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        next_command = takeCommand().lower()
                        if("set timer" in next_command):
                            x = 0
                            for i in next_command:
                                if("hour" in next_command):
                                    if (i>= '0' and i<='9'):
                                        x=x*10+int(i)
                                        x_2 = x*3600
                            new_time_for_sleep = x_2
                            print(new_time_for_sleep)
                            time.sleep(new_time_for_sleep)
                            speak("ok restarting pc")
                            os.system("shutdown /r /t 1")
                            
                        elif("right now" in next_command):
                            speak("ok sir restarting pc")
                            os.system("shutdown /r /t 1")	
                        
                        elif("cancel" in next_command):
                            speak("ok sir request turned down")
                        else:
                            speak("cant understand that command please try again")
                    elif("log off" in mytext):
                        speak("sir you wanna set timer or want to do it right now")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        next_command = takeCommand().lower()
                        if("set timer" in next_command):
                            x = 0
                            for i in next_command:
                                if("hour" in next_command):
                                    if (i>= '0' and i<='9'):
                                        x=x*10+int(i)
                                        x_2 = x*3600
                            new_time_for_sleep = x_2
                            print(new_time_for_sleep)
                            time.sleep(new_time_for_sleep)
                            speak("ok logging off")
                            os.system("shutdown -l ")
                            
                            
                        elif("right now" in next_command):
                            speak("ok sir logging it off")
                            os.system("shutdown -l ")	
                        
                        elif("cancel" in next_command):
                            speak("ok sir request turned down")
                        else:
                            speak("cant understand that command please try again")

                    elif("search on youtube" in mytext or "do a youtube" in mytext or "open youtube" in mytext):
                        speak("what you want to search")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        next_command_1 = takeCommand().lower()
                        speak("should i open first video or show the whole list")
                        next_command_2 = takeCommand().lower()
                        if "first video" in next_command_2:
                            new = (next_command_1.replace(" ",""))
                            next_command_1 = new
                            print(new)
                            speak("opening"+new+"on youtube")
                            print("opening"+ " " + new)
                            search_keyword=new
                            html = urllib.request.urlopen("https://www.youtube.com/results?search_query=" + search_keyword)
                            video_ids = re.findall(r"watch\?v=(\S{11})", html.read().decode())
                            new2 = ("https://www.youtube.com/watch?v=" + video_ids[0])
                            webbrowser.open(new2)
                            speak("done sir")
                        elif "whole list" in next_command_2:
                                website = 'https://www.youtube.com/results?search_query='
                                com = "&page&utm_source=opensearch"
                                new2 =  website + next_command_1 + com
                                webbrowser.open(new2)
                                speak("done sir")
                  
                    elif("what time is it" in mytext):
                        x = datetime.now() 
                        times = x.strftime("%I  %M ")
                        speak(times)

                    elif ("jor se bolo" in mytext):
                        winsound.PlaySound('final_jai_mata_di.wav', winsound.SND_FILENAME)
                    

                    elif("change volume to " in mytext):
                        speak("changing volume sir")
                        x = 0
                        for i in mytext:
                            if (i>= '0' and i<='9'):
                                x=x*10+int(i)
                        devices = AudioUtilities.GetSpeakers()
                        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        volume = cast(interface, POINTER(IAudioEndpointVolume))
                        new_character = int(x)/100
                        volume.SetMasterVolumeLevelScalar(new_character, None) 
                        speak("volume changed sir")
                    elif("pause it" in mytext or "play it" in mytext):
                        keyboard = Controller()
                        keyboard.press(Key.space)
                        speak("ok sir")
                    elif("scroll down" in mytext or  "scroll it down" in mytext):
                        keyboard = Controller()
                        keyboard.press(Key.page_down )
                    elif("scroll up" in mytext):
                        keyboard = Controller()
                        keyboard.press(Key.page_up)
                    elif("bro" in mytext):
                        speak("hello sir how can i help you")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                    
                    elif("turn wifi" in mytext):
                        if("on" in mytext):
                            speak("ok turning wifi on")
                            os.system("netsh interface show interface")
                            os.system("netsh interface set interface name = "'Wi-Fi'" admin=ENABLED ")
                        elif("off" in mytext):
                            speak("ok turning wifi off")
                            os.system("netsh interface show interface")
                            os.system("netsh interface set interface name = "'Wi-Fi'" admin=DISABLED ")
                    #calculation
                    elif("do a sum" in mytext):
                        speak("please tell me the equation")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        otherText=mytext
                        new = otherText.replace(mytext,"")
                        mytext = new
                        calculator()
                    elif("open word" in mytext):
                        speak("opening word")
                        os.startfile(r"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE")

                    elif("start typing" in mytext):
                        speak("what should i type sir")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        typing()
                    
                    elif("save it" in mytext):
                        import keyboard
                        keyboard.press_and_release('ctrl + s')

                    elif("close it" in mytext):
                        import keyboard 
                        keyboard.press_and_release('alt+F4')
                    # list update
                    elif("update my list" in mytext):
                        speak("what i have to add sir")
                        next_command = takeCommand().lower()
                        file = open('geek.txt','a')
                        file.write( "\t" ) 
                        file.write(next_command) 
                        file.write( "!" )
                        file.close()
                        speak("list updated sir")
                    elif("read my list" in mytext):
                        speak("reading file sir")
                        file = open('geek.txt' , 'r')
                        speak(file.read())
                        print(file.read())
                    # screen shot vala code
                    elif("take" and "screenshot" in mytext):
                        speak("what should i name the file sir")
                        otherText=mytext
                        new = otherText.replace(mytext,"")
                        mytext = new
                        take_screenshot()
                    # key bindings vala code 
                    elif("press" in mytext):
                        new_2 = mytext[6:]
                        import keyboard
                        keyboard.press_and_release(new_2)
                    
                    elif("play" and "music" in mytext or "play" and "song" in mytext):
                        
                        if("song" in mytext):
                            speak("which song sir")
                        else:
                            speak("what music would you like sir")
                        
                        next_command_1 = takeCommand().lower()
                        speak("ok sir wait a second")
                        new = (next_command_1.replace(" ",""))
                        next_command_1 = new
                        path = r"path of chrome driver"
                        options = webdriver.ChromeOptions()
                        options.headless = True

                        search_keyword=new
                        html = urllib.request.urlopen("https://www.youtube.com/results?search_query=" + search_keyword)
                        video_ids = re.findall(r"watch\?v=(\S{11})", html.read().decode())
                        new2 = ("https://www.youtube.com/watch?v=" + video_ids[0])
                        driver = webdriver.Chrome(path,options=options)
                        driver.get(new2)
                        driver.maximize_window() 
                        wait = WebDriverWait(driver, 3)
                        #presence = EC.presence_of_element_located
                        visible = EC.visibility_of_element_located
                        # Play the video.
                        wait.until(visible((By.ID, "video-title")))
                        driver.find_element_by_id("movie_player").click()
                        try:
                            driver.find_element_by_id("toggleButton").click()
                        except:
                            pass
                        print(driver.title)
                        
                        
                    
                    #only for stopping music which is playing
                    elif("stop" and "this" in mytext):
                        speak("ok sir")
                        driver.close()
                    elif("launch the moving program" in mytext):
                        old_path = r"D:\\testing"
                        dirs = os.listdir(old_path)
                        #slen = len(dirs)
                        
                        for stuff in dirs:
                            if ".png" in stuff:
                                old_path_1 = r"D:\\testing" + "\\" + stuff
                                print(old_path_1)
                                shutil.move(old_path_1, r"D:\\testing_2")
                                print("done!")
                                
                            elif(".docx" in stuff):
                                old_path_1 = r"D:\\testing" + "\\" + stuff
                                print(old_path_1)
                                shutil.move(old_path_1, r"D:\\testing_2")
                                print("done!")
                                
                        speak("sir moving program is successfully completed")
                    
                    
                    elif ("open notepad" in mytext):
                        speak("opening notepad")
                        os.startfile(r"C:\\Windows\\system32\\notepad.exe")
                    
                    elif("terminate" in mytext):
                        speak("ok sir i am out!")
                        sys.exit()
                    
                        
                      
                    else:
                        speak("i cant understand that yet please say the right command")
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                        winsound.PlaySound('beep_sound.wav', winsound.SND_FILENAME)
                    #elif("open recent document" in mytext):code m dikkat h path sahi nahi mil raha
                        #speak("opening word")
                        #os.startfile(r"C:\Users\Samyak\Desktop\last opened document.EXE")
          
                except sr.RequestError as e: 
                    print("Could not request results; {0}".format(e)) 
                except sr.UnknownValueError: 
                    print("Can't hear anything") 
 
if __name__ == '__main__':
    while True:
        try:
            print("still listening")
            Permission = takeCommand()
            if "wake up" in Permission:
                working()
            elif "bye" in Permission or "shut up" in Permission :
                speak("ok sir i am out")
                sys.exit()
        except sr.RequestError as e: 
            print("Could not request results; {0}".format(e)) 
        except sr.UnknownValueError: 
            print("right now in sleep")

