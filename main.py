import webbrowser
import speech_recognition as sr
import os
import pyttsx3
import datetime
import wikipedia
import sys
from sys import platform
import psutil
from bs4 import BeautifulSoup
import smtplib
import win32com.client
import openai
import pyautogui
import requests
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
import time
import pyautogui
import webbrowser



engine = pyttsx3.init()

openai.api_key = " "

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    pass


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        speak("Listening ...")
        r.pause_threshold = 1
        try:
            audio = r.listen(source, timeout=5)  # Set a timeout to avoid indefinite waiting
        except sr.WaitTimeoutError:
            print("Timeout. No audio input.")
            speak("Timeout. No audio input.")
            return " "
        
    try:
        print("Recognizing")
        speak("Recognizing")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}")
        return query

    except sr.UnknownValueError:
        print("Sorry, I couldn't understand the audio.")
        speak("Sorry, I couldn't understand the audio.")
        return " "

    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")
        speak("Could not request results from Google Speech Recognition service.")
        return " "

    except Exception as e:
        print(f"An error occurred: {e}")
        speak("An error occurred.")
        return " "




def send_email():
    # Get the email details from the user
    speak("Enter the recipient's email address: ")
    recipient_email = input("Enter the recipient's email address: ")
    speak("Enter the email subject: ")
    subject = input("Enter the email subject: ")
    speak("Enter the email message: ")
    message = input("Enter the email message: ")

    # Get the sender's email credentials
    speak("Enter your email address: ")
    sender_email = input("Enter your email address: ")
    speak("Enter your email password: ")
    password = input("Enter your email password: ")

    try:
        # Establish an SMTP connection
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, password)

        # Compose the email
        email_body = f"Subject: {subject}\n\n{message}"

        # Send the email
        server.sendmail(sender_email, recipient_email, email_body)
        speak("Email sent successfully!")
        print("Email sent successfully!")

    except Exception as e:
        speak("An error occurred while sending the email:")
        print("An error occurred while sending the email:", str(e))


def Temperature():
    city = query.split("in",1)
    url = f"https://www.google.com/search?q=weather+in+{city}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")
    # Find the temperature element by searching for a span with the temperature class
    temp = soup.find("span", {"class": "wob_t"})
    # Find the region or location element by searching for a span with the region class
    region = soup.find("span", {"class": "wob_loc"})
    # Find the weather description element by searching for a span with the weather class
    weather = soup.find("span", {"class": "wob_dcp"})
    if temp and region and weather:
        temperature = temp.text
        location = region.text
        weather_description = weather.text
        speak(f"It's currently {weather_description} and {temperature} in {location}")
    else:
        speak("Weather information not found.")

# def Temperature(query):
#     city = query.split("in", 1)
#     if len(city) < 2:
#         speak("Please specify a city.")
#         return

#     city_name = city[1].strip()
#     formatted_city = quote(city_name)
#     url = f"https://www.google.com/search?q=weather+in+{formatted_city}"
    
#     response = requests.get(url)
#     soup = BeautifulSoup(response.text, "html.parser")

#     temp = soup.find("span", {"class": "wob_t"})
#     region = soup.find("span", {"class": "wob_loc"})
#     weather = soup.find("span", {"class": "wob_dcp"})

#     if temp and region and weather:
#         temperature = temp.text
#         location = region.text
#         weather_description = weather.text
#         speak(f"It's currently {weather_description} and {temperature} in {location}")
#     else:
#         speak("Weather information not found.")

        
def openai_chat():
    print("Chat mode activated. How can I assist you?")
    speak("Chat mode activated. How can I assist you?")
    active = True  # Flag to track if chat mode is active

    while active:
        user_input = takecommand()
        if user_input:
            if "chat mode deactivate" in user_input:
                print("Deactivating chat mode.")
                speak("Deactivating chat mode.")
                active = False  # Set the flag to False to exit the loop
            else:
                response = generate_response(user_input)
                print("Assistant:", response)
                speak(response)
        else:
            True


def generate_response(prompt):
    response = openai.Completion.create(
        engine="text-babbage-001",
        prompt=prompt,
        max_tokens=160,
        n=1,
        stop=None,
        temperature=0.7
    )
    return response.choices[0].text.strip()


def take_screenshot():
    screenshot = pyautogui.screenshot()
    screenshot_path = "C:/Users/heman/OneDrive/Pictures/Screenshots/scrren.png"
    screenshot.save(screenshot_path)
    speak("Screenshot saved successfully.")
    print("Screenshot saved successfully.")

def close_program():
    pyautogui.hotkey('alt', 'f4')

def open_task_manager():
    pyautogui.hotkey('ctrl','shift','esc')

def check_driver_errors():
    wmi = win32com.client.GetObject("winmgmts:")
    devices = wmi.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0")

    if len(devices) > 0:
        print("Driver errors found:")
        for device in devices:
            print(f"Device Name: {device.Name}")
            print(f"Error Code: {device.ConfigManagerErrorCode}")
            print(f"Error Description: {device.ConfigManagerUserConfig}")
            print()

    else:
        print("No driver errors found.")

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour <= 12:
        speak("Good morning !")
    elif hour >= 12 and hour <= 18:
        speak("Good afternoon !")

    else:
        speak("Good evening!")

    print("I am Teja sir... Please tell me how may I help you ")
    speak("I am Teja sir... Please tell me how may I help you ")

def end_running_apps():
    for proc in psutil.process_iter():
        try:
            proc.terminate()
        except psutil.AccessDenied:
            # Skip processes that can't be terminated due to access restrictions
            continue


def open_website(url):
    speak(f"Opening {site[0]}...")
    webbrowser.open(url)


def search_wikipedia(query):
    print("Searching Wikipedia...")
    speak("Searching Wikipedia...")
    query = query.replace("wikipedia", "")
    results = wikipedia.summary(query, sentences=2)
    speak("According to Wikipedia")
    print(results)
    speak(results)


def play_music():
    music_path = r"https://wynk.in/u/uttWGzeSR"
    os.startfile(music_path)

def stop_music():
    # Find the process ID of the media player (e.g., Windows Media Player)
    for proc in psutil.process_iter():
        if "wmplayer.exe" in proc.name().lower():
            proc.kill()
            break
    


def search_google(query):
    query = query.replace("search", "").strip()
    query = query.replace("in google", "").strip()
    formatted_query = query.replace(' ', '+')
    search_url = f"https://www.google.com/search?q={formatted_query}"
    webbrowser.open(search_url)

def get_current_time():
    current_time = datetime.datetime.now().strftime("%H:%M")
    speak(f"The current time is {current_time}")


def open_downloads_folder():
    downloads_path = r"C:/Users/heman/Downloads"
    os.startfile(downloads_path)


def stop_personal_assistant():
    print("Your personal assistant is stopping...")
    speak("Your personal assistant is stopping...")
    sys.exit()


def shutdown_system():
    if platform == "win32":
        os.system('shutdown /p /f')
    elif platform == "linux" or platform == "linux2" or platform == "darwin":
        os.system('poweroff')


def clear_all_apps():
    end_running_apps()

def remember_message():
    speak("What should I remember, sir?")
    remember_message = takecommand()
    speak("You told me to remember " + remember_message)
    remember = open('data.txt', 'w')
    remember.write(remember_message)
    remember.close()


def recall_message():
    remember = open('data.txt', 'r')
    speak("You told me to remember this: " + remember.read())

def reset_chat():
    chatStr = " "


def open_location_in_web():
    speak("What is the location?")
    location = takecommand()
    url = 'https://google.nl/maps/place/' + location + '/&amp;'
    try:
        webbrowser.open_new_tab(url)
        speak('Here is the location ' + location)
    except webbrowser.Error:
        webbrowser.open_new_tab(url)
        speak('Here is the location ' + location)

def navigate_me():
    speak("what is your location to start ! ")
    location = takecommand()
    speak("What is the destination ?")
    destination = takecommand()
    url = f"https://www.google.com/maps/dir/{location}/{destination}"
    try:
        webbrowser.open_new_tab(url)
        speak(f"Navigating to {destination}")
    except webbrowser.Error:
        webbrowser.open_new_tab(url)
        speak(f"Navigating to {destination}")

def open_app(app_name, app_path):
    speak(f"Opening {app_name}...")
    os.startfile(app_path)

def conversations(input):
    if "hi" in input:
        speak("Hi, how are you?")
        print("Hi, how are you?")
    elif "how are you" in input:
        speak("I'm fine, thank you. How about you, sir?")
        print("I'm fine, thank you. How about you, sir?")
    elif "good" in input or "fine" in input:
        speak("That's great to hear!")
        print("That's great to hear!")
    elif "what's your name" in input:
        speak("I'm an AI assistant. You can call me Mark.")
        print("I'm an AI assistant. You can call me Mark.")
    elif "tell me a joke" in input:
        speak("Sure, here's one: Why don't scientists trust atoms? Because they make up everything!")
        print("Sure, here's one: Why don't scientists trust atoms? Because they make up everything!")
    elif "what can you do" in input:
        speak("I can answer questions, provide information, and assist you with various tasks.")
        print("I can answer questions, provide information, and assist you with various tasks.")
    else:
        speak("I'm sorry, I didn't understand. Can you please rephrase?")
        print("I'm sorry, I didn't understand. Can you please rephrase?")


if __name__ == "__main__":
    wishMe()
    while True:
        query = takecommand().lower()

        sites = [
            ["youtube", "https://www.youtube.com/"],
            ["wikipedia", "https://www.wikipedia.org/"],
            ["google", "https://www.google.com/"],
            ["facebook", "https://www.facebook.com/"],
            ["twitter", "https://www.twitter.com/"],
            ["instagram", "https://www.instagram.com/"],
            ["reddit", "https://www.reddit.com/"],
            ["amazon", "https://www.amazon.com/"],
            ["ebay", "https://www.ebay.com/"],
            ["netflix", "https://www.netflix.com/"],
            ["linkedin", "https://www.linkedin.com/"],
            ["pinterest", "https://www.pinterest.com/"],
            ["twitch", "https://www.twitch.tv/"],
            ["yahoo", "https://www.yahoo.com/"],
            ["github", "https://github.com/"],
            ["wordpress", "https://wordpress.com/"],
            ["imdb", "https://www.imdb.com/"],
            ["craigslist", "https://www.craigslist.org/"],
            ["spotify", "https://www.spotify.com/"],
            ["walmart", "https://www.walmart.com/"],
            ["apple", "https://www.apple.com/"],
            ["bing", "https://www.bing.com/"],
            ["stackoverflow", "https://stackoverflow.com/"],
            ["quora", "https://www.quora.com/"],
            ["espn", "https://www.espn.com/"],
            ["weather", "https://www.weather.com/"],
            ["cnn", "https://www.cnn.com/"],
            ["nytimes", "https://www.nytimes.com/"],
            ["bbc", "https://www.bbc.co.uk/"],
            ["foxnews", "https://www.foxnews.com/"],
            ["hulu", "https://www.hulu.com/"],
            ["cnet", "https://www.cnet.com/"],
            ["forbes", "https://www.forbes.com/"],
            ["buzzfeed", "https://www.buzzfeed.com/"],
            ["flickr", "https://www.flickr.com/"],
            ["tumblr", "https://www.tumblr.com/"],
            ["dropbox", "https://www.dropbox.com/"],
            ["vimeo", "https://www.vimeo.com/"],
            ["paypal", "https://www.paypal.com/"],
            ["microsoft", "https://www.microsoft.com/"],
            ["target", "https://www.target.com/"],
            ["adobe", "https://www.adobe.com/"],
            ["wellsfargo", "https://www.wellsfargo.com/"],
            ["chase", "https://www.chase.com/"],
            ["bankofamerica", "https://www.bankofamerica.com/"],
            ["citibank", "https://online.citi.com/"],
            ["hsbc", "https://www.hsbc.com/"],
            ["americanexpress", "https://www.americanexpress.com/"],
            ["capitalone", "https://www.capitalone.com/"],
            ["usbank", "https://www.usbank.com/"],
            ["nike", "https://www.nike.com/"],
            ["adidas", "https://www.adidas.com/"],
            ["newegg", "https://www.newegg.com/"],
            ["sathyabama placement","https://placement.sathyabama.ac.in/"],
            ["sathyabama erp","https://erp.sathyabama.ac.in/account/login?returnUrl=%2F"]]

        for site in sites:
            if f"open {site[0]}".lower() in query:
                open_website(site[1])
                break

        if 'wikipedia' in query:
            search_wikipedia(query)

        elif 'task manager' in query :
            open_task_manager()

        elif 'chat mode' in query or 'chat with' in query:
            openai_chat()

        elif 'play music' in query:
            play_music()

        elif 'the time' in query:
            get_current_time()

        elif 'open my downloads' in query:
            open_downloads_folder()

        elif 'terminate' in query or 'bye' in query or 'sleep' in query:
            stop_personal_assistant()

        elif 'shutdown' in query:
            shutdown_system()

        elif 'clear all apps' in query:
            clear_all_apps()

        elif 'remember that' in query:
            remember_message()

        elif 'do you remember anything' in query:
            recall_message()

        elif 'location in web' in query:
            open_location_in_web()

        elif 'clear chat' in query:
            reset_chat()

        elif 'driver' in query:
            check_driver_errors()

        elif 'send a mail' in query:
            send_email()

        elif 'stop music' in query:
            stop_music()

        elif 'screenshot' in query:
            take_screenshot()

        elif 'close current program' in query:
            close_program()

        elif 'in google' in query:
            search_google(query)

        elif 'navigate me' in query :
            navigate_me()

        elif 'temperature in ' in query:
            Temperature()
        

        apps = [
            ["notepad", "C:\\Windows\\system32\\notepad.exe"],
            ["calculator", "C:\\Windows\\system32\\calc.exe"],
            ["chrome", "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"],
            ["vlc", "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe"],
            ["brave", "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"],
            ["file explorer", "C:\\Windows\\explorer.exe"],
            ["task manager", "C:\\Windows\\system32\\taskmgr.exe"],
            ["command prompt", "C:\\Windows\\system32\\cmd.exe"],
            ["settings", "C:\\Windows\\system32\\control.exe"],
            ["snipping tool", "C:\\Windows\\system32\\SnippingTool.exe"],
            ["task scheduler", "C:\\Windows\\system32\\taskschd.msc"],
            ["anydesk", "C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe"]
        ]

        for app in apps:
            if f"open {app[0]}".lower() in query:
                open_app(app[0], app[1])
                break

