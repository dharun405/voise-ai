import speech_recognition as sr
import pyttsx3
import subprocess
import webbrowser
import platform
import os
import datetime
import time
import sys
import threading

# Initialize the speech recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Adjust speech rate for better understanding
engine.setProperty('rate', 150)

# Get the operating system for platform-specific commands
OS_NAME = platform.system()

# Wake word variations
WAKE_WORDS = ["hi joel", "hey joel", "hello joel", "", "hi jo", "hey jo","hello jo"]

# Global flag to track if assistant is listening
listening_active = False

# Function to speak text
def speak(text):
    print(f"joel: {text}")
    engine.say(text)
    engine.runAndWait()

# Improved listening function with better error handling
def listen(timeout=5):
    global listening_active
    try:
        with sr.Microphone() as source:
            print("🎤 Listening...")
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            # More lenient timeout settings
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=6)
            
            # Try to recognize speech
            command = recognizer.recognize_google(audio).lower()
            print(f"✅ You said: {command}")
            return command
            
    except sr.UnknownValueError:
        print("❌ Could not understand audio")
        return None
    except sr.RequestError as e:
        print(f"❌ Google Speech Recognition error: {e}")
        return None
    except sr.WaitTimeoutError:
        print("⏰ Listening timeout")
        return None
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return None

# Function to wait for wake word with better detection
def wait_for_wake_word():
    print("🔄 Waiting for wake word...")
    print("💡 Say: 'Hi joel', 'Hey joel', 'Hello joel', or just 'joel'")
    
    while True:
        command = listen(timeout=10)
        if command:
            # Check if any wake word is in the command
            if any(wake_word in command for wake_word in WAKE_WORDS):
                speak("Yes, I'm here! How can I help you?")
                return True
            else:
                print(f"🔍 Heard: '{command}' - Not a wake word, continuing to listen...")
        time.sleep(0.1)

# Improved application opening function
def open_application(app_name):
    app_name = app_name.lower().strip()
    speak(f"Opening {app_name}")
    
    try:
        if OS_NAME == "Windows":
            # Dictionary of applications with their executable commands
            app_commands = {
                "notepad": "notepad.exe",
                "calculator": "calc.exe",
                "paint": "mspaint.exe",
                "word": "start winword",
                "excel": "start excel",
                "powerpoint": "start powerpnt",
                "chrome": "start chrome",
                "firefox": "start firefox",
                "cmd": "start cmd",
                "command prompt": "start cmd",
                "file explorer": "start explorer",
                "task manager": "taskmgr",
                "control panel": "control",
                "settings": "start ms-settings:"
            }
            
            # Find the best match for the app name
            matched_app = None
            for app_key in app_commands.keys():
                if app_key in app_name or app_name in app_key:
                    matched_app = app_key
                    break
            
            if matched_app:
                command = app_commands[matched_app]
                print(f"🔧 Executing: {command}")
                
                if command.startswith("start"):
                    os.system(command)
                else:
                    subprocess.Popen(command, shell=True)
                    
                speak(f"{matched_app} opened successfully!")
                return True
            else:
                # Try to open with start command for unknown apps
                os.system(f"start {app_name}")
                speak(f"Trying to open {app_name}")
                return True
                
    except Exception as e:
        speak(f"Sorry, I couldn't open {app_name}. Error: {str(e)}")
        print(f"❌ Error opening {app_name}: {e}")
        return False

# Website opening function
def open_website(site_name):
    site_name = site_name.lower().strip()
    speak(f"Opening {site_name}")
    
    try:
        # Common websites
        websites = {
            "youtube": "https://www.youtube.com",
            "google": "https://www.google.com",
            "facebook": "https://www.facebook.com",
            "twitter": "https://www.twitter.com",
            "instagram": "https://www.instagram.com",
            "linkedin": "https://www.linkedin.com",
            "github": "https://www.github.com",
            "stackoverflow": "https://stackoverflow.com",
            "reddit": "https://www.reddit.com",
            "netflix": "https://www.netflix.com",
            "amazon": "https://www.amazon.com",
            "gmail": "https://mail.google.com",
            "yahoo": "https://www.yahoo.com",
            "bing": "https://www.bing.com"
        }
        
        # Find the best match
        matched_site = None
        for site_key in websites.keys():
            if site_key in site_name or site_name in site_key:
                matched_site = site_key
                break
        
        if matched_site:
            webbrowser.open(websites[matched_site])
            speak(f"{matched_site} opened successfully!")
        else:
            # Try to open as a general website
            if not site_name.startswith("http"):
                site_name = f"https://www.{site_name}.com"
            webbrowser.open(site_name)
            speak(f"Website opened successfully!")
            
        return True
    except Exception as e:
        speak(f"Sorry, I couldn't open {site_name}")
        print(f"❌ Error opening website: {e}")
        return False

# Search function
def search_web(query):
    speak(f"Searching for {query}")
    try:
        search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        webbrowser.open(search_url)
        speak("Here are the search results")
        return True
    except Exception as e:
        speak(f"Sorry, I couldn't search for {query}")
        print(f"❌ Search error: {e}")
        return False

# Time and date functions
def get_time():
    current_time = datetime.datetime.now().strftime("%I:%M %p")
    speak(f"The current time is {current_time}")

def get_date():
    current_date = datetime.datetime.now().strftime("%A, %B %d, %Y")
    speak(f"Today is {current_date}")

def get_day():
    current_day = datetime.datetime.now().strftime("%A")
    speak(f"Today is {current_day}")

# System control functions
def system_shutdown():
    speak("Are you sure you want to shut down the computer? Say yes to confirm.")
    confirmation = listen(timeout=10)
    if confirmation and "yes" in confirmation:
        speak("Shutting down the computer in 5 seconds. Goodbye!")
        os.system("shutdown /s /t 5")
    else:
        speak("Shutdown cancelled.")

def system_restart():
    speak("Are you sure you want to restart the computer? Say yes to confirm.")
    confirmation = listen(timeout=10)
    if confirmation and "yes" in confirmation:
        speak("Restarting the computer in 5 seconds. Goodbye!")
        os.system("shutdown /r /t 5")
    else:
        speak("Restart cancelled.")

def take_screenshot():
    speak("Taking a screenshot")
    try:
        # Use Windows Snipping Tool
        os.system("snippingtool")
        speak("Screenshot tool opened")
    except Exception as e:
        speak("Sorry, I couldn't take a screenshot")
        print(f"❌ Screenshot error: {e}")

# Enhanced command processing
def process_command(command):
    if not command:
        return True
    
    command = command.lower().strip()
    print(f"🔍 Processing command: '{command}'")
    
    # Exit commands
    if any(word in command for word in ["goodbye", "bye", "exit", "quit", "stop", "sleep"]):
        speak("Going to sleep. Say 'Hi joel' to wake me up again!")
        return False
    
    # Application commands - improved parsing
    elif "open" in command:
        if "application" in command:
            app_name = command.replace("open", "").replace("application", "").strip()
            open_application(app_name)
        else:
            app_name = command.replace("open", "").strip()
            # Check if it's a common website
            if any(site in app_name for site in ["youtube", "google", "facebook", "twitter", "instagram", "linkedin", "github", "stackoverflow", "reddit", "netflix", "amazon", "gmail"]):
                open_website(app_name)
            else:
                open_application(app_name)
    
    # Website commands
    elif "go to" in command:
        site_name = command.replace("go to", "").strip()
        open_website(site_name)
    
    # Search commands
    elif "search for" in command:
        query = command.replace("search for", "").strip()
        search_web(query)
    elif "search" in command and "for" not in command:
        query = command.replace("search", "").strip()
        search_web(query)
    
    # Time and date commands
    elif any(phrase in command for phrase in ["what time", "current time", "time is"]):
        get_time()
    elif any(phrase in command for phrase in ["what date", "current date", "date is", "today's date"]):
        get_date()
    elif any(phrase in command for phrase in ["what day", "current day", "day is", "today is"]):
        get_day()
    
    # System commands
    elif "shutdown" in command or "shut down" in command:
        system_shutdown()
    elif "restart" in command:
        system_restart()
    elif "screenshot" in command:
        take_screenshot()
    
    # Information commands
    elif any(phrase in command for phrase in ["what can you do", "help", "commands"]):
        speak("I can open applications like notepad, calculator, chrome. Open websites like youtube, google, facebook. Search the web, tell you the time and date, take screenshots, and control your computer. Just say what you need!")
    
    elif any(phrase in command for phrase in ["what is your name", "who are you", "your name"]):
        speak("I am joel, your personal voice assistant. I'm here to help you with various tasks on your computer.")
    
    elif any(phrase in command for phrase in ["how are you", "how do you do"]):
        speak("I'm doing great! Ready to help you with anything you need.")
    
    # Default response
    else:
        speak("I'm not sure how to help with that. Try saying 'help' to see what I can do, or try rephrasing your request.")
    
    return True

# Main assistant loop
def run_assistant():
    speak("Hello! I'm joel, your personal voice assistant.")
    speak("Say 'Hi joel' anytime to wake me up!")
    
    while True:
        try:
            # Wait for wake word
            wait_for_wake_word()
            
            # Process commands until user says goodbye
            while True:
                command = listen(timeout=15)
                if not process_command(command):
                    break
                
                # Small delay before next command
                time.sleep(0.3)
                
        except KeyboardInterrupt:
            speak("Assistant stopped. Goodbye!")
            break
        except Exception as e:
            print(f"❌ Unexpected error in main loop: {e}")
            speak("I encountered an error. Restarting...")
            time.sleep(2)

# Start the voice assistant
if __name__ == "__main__":
    try:
        run_assistant()
    except KeyboardInterrupt:
        speak("Assistant stopped. Goodbye!")
        sys.exit(0)
