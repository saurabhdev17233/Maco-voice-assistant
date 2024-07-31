import speech_recognition as sr
import os
import nltk
from nltk.tokenize import word_tokenize
from tkinter import Tk, Label, Button, Text, Scrollbar, VERTICAL, END
import threading
import logging
import pyttsx3
import spacy
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import time
import openai
import webbrowser

# Download NLTK data
nltk.download('punkt')

# Setup logging
logging.basicConfig(filename='maco.log', level=logging.INFO, format='%(asctime)s - %(message)s')

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Load spaCy model
nlp = spacy.load('en_core_web_sm')

# Initialize OpenAI API
openai.api_key = 'your_openai_api_key'

# Context management
conversation_context = []

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
            logging.info(f"Command recognized: {command}")
            return command.lower()
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            logging.error("UnknownValueError: Could not understand audio")
            speak("Sorry, I did not understand that.")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            logging.error(f"RequestError: {e}")
            speak("Sorry, my speech service is down.")
            return ""

def converse_with_gpt3(prompt):
    global conversation_context
    conversation_context.append({"role": "user", "content": prompt})
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=conversation_context
    )
    reply = response.choices[0].message['content'].strip()
    conversation_context.append({"role": "assistant", "content": reply})
    return reply

def execute_command(command):
    doc = nlp(command)
    tokens = [token.text for token in doc]
    
    if 'open' in tokens and 'notepad' in tokens:
        os.system('notepad')
        speak("Opening Notepad")
    elif 'open' in tokens and 'calculator' in tokens:
        os.system('calc')
        speak("Opening Calculator")
    elif 'shutdown' in tokens:
        os.system('shutdown /s /t 1')
        speak("Shutting down the system")
    elif 'restart' in tokens:
        os.system('shutdown /r /t 1')
        speak("Restarting the system")
    elif 'play' in tokens and 'music' in tokens:
        os.system('start wmplayer')
        speak("Playing music")
    elif 'open' in tokens and 'browser' in tokens:
        os.system('start chrome')
        speak("Opening browser")
    elif 'search' in tokens:
        search_query = " ".join([token.text for token in doc if token.text not in ['search', 'for']])
        search_web(search_query)
    elif 'open' in tokens and 'word' in tokens:
        os.system('start winword')
        speak("Opening Microsoft Word")
    elif 'open' in tokens and 'excel' in tokens:
        os.system('start excel')
        speak("Opening Microsoft Excel")
    elif 'lock' in tokens:
        os.system('rundll32.exe user32.dll,LockWorkStation')
        speak("Locking the workstation")
    elif 'log' in tokens and 'off' in tokens:
        os.system('shutdown /l')
        speak("Logging off")
    elif 'sleep' in tokens:
        os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')
        speak("Putting the system to sleep")
    elif 'create' in tokens and 'file' in tokens:
        with open('newfile.txt', 'w') as f:
            f.write('New file created')
        speak("Creating a new file")
    elif 'delete' in tokens and 'file' in tokens:
        os.remove('newfile.txt')
        speak("Deleting the file")
    elif 'connect' in tokens and 'wifi' in tokens:
        os.system('netsh wlan connect name=YourWiFiName')
        speak("Connecting to Wi-Fi")
    elif 'disconnect' in tokens and 'wifi' in tokens:
        os.system('netsh wlan disconnect')
        speak("Disconnecting from Wi-Fi")
    elif 'weather' in tokens:
        get_weather()
    elif 'news' in tokens:
        get_news()
    elif 'send' in tokens and 'email' in tokens:
        send_email()
    elif 'set' in tokens and 'reminder' in tokens:
        set_reminder(command)
    elif 'set' in tokens and 'alarm' in tokens:
        set_alarm(command)
    else:
        # Use GPT-3 for more complex or conversational commands
        response = converse_with_gpt3(command)
        speak(response)
        logging.info(f"GPT-3 response: {response}")

def search_web(query):
    url = f"https://www.google.com/search?q={query}"
    webbrowser.open(url)
    speak(f"Searching for {query} on Google")

def get_weather():
    api_key = "your_openweather_api_key"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    city_name = "your_city_name"
    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
    response = requests.get(complete_url)
    data = response.json()
    if data["cod"] != "404":
        main = data["main"]
        weather_description = data["weather"][0]["description"]
        temperature = main["temp"]
        speak(f"The temperature is {temperature - 273.15:.2f} degrees Celsius with {weather_description}.")
    else:
        speak("City not found.")

def get_news():
    api_key = "your_newsapi_key"
    base_url = "https://newsapi.org/v2/top-headlines?"
    country = "us"
    complete_url = base_url + "country=" + country + "&apiKey=" + api_key
    response = requests.get(complete_url)
    data = response.json()
    if data["status"] == "ok":
        articles = data["articles"]
        speak("Here are the top news headlines.")
        for article in articles[:5]:
            speak(article["title"])
    else:
        speak("Unable to fetch news at the moment.")

def send_email():
    sender_email = "your_email@example.com"
    receiver_email = "receiver_email@example.com"
    password = "your_email_password"

    subject = "Test Email"
    body = "This is a test email from your voice assistant."

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        speak("Email has been sent successfully.")
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        speak("Failed to send email.")

def set_reminder(command):
    doc = nlp(command)
    reminder_text = " ".join([token.text for token in doc if token.text not in ['set', 'reminder', 'for']])
    speak(f"Reminder set for: {reminder_text}")
    # Here you can add code to save the reminder to a file or a database

def set_alarm(command):
    doc = nlp(command)
    time_text = " ".join([token.text for token in doc if token.text not in ['set', 'alarm', 'for']])
    alarm_time = datetime.strptime(time_text, '%H:%M')
    speak(f"Alarm set for {alarm_time.strftime('%I:%M %p')}")
    while True:
        if datetime.now().time() >= alarm_time.time():
            speak("Wake up! It's time!")
            break
        time.sleep(1)

def start_listening(app):
    while app.listening:
        command = listen_command()
        if command:
            app.log_command(command)
            execute_command(command)

# GUI Setup
class MacoApp:
    def __init__(self, master):
        self.master = master
        master.title("Maco Voice Assistant")

        self.label = Label(master, text="Press the button and speak a command")
        self.label.pack()

        self.listen_button = Button(master, text="Listen", command=self.start_listening_thread)
        self.listen_button.pack()

        self.stop_button = Button(master, text="Stop", command=self.stop_listening)
        self.stop_button.pack()

        self.text_area = Text(master, wrap='word', height=10, width=50)
        self.text_area.pack()

        self.scrollbar = Scrollbar(master, orient=VERTICAL, command=self.text_area.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.text_area['yscrollcommand'] = self.scrollbar.set

        self.listening = False

    def log_command(self, command):
        self.text_area.insert(END, f"You said: {command}\n")
        self.text_area.see(END)

    def start_listening_thread(self):
        self.listening = True
        self.listen_thread = threading.Thread(target=start_listening, args=(self,))
        self.listen_thread.start()

    def stop_listening(self):
        self.listening = False
        if hasattr(self, 'listen_thread'):
            self.listen_thread.join()

if __name__ == '__main__':
    root = Tk()
    maco_app = MacoApp(root)
    root.mainloop()