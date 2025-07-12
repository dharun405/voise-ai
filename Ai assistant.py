import speech_recognition as sr
import pyttsx3
import nltk
from nltk.chat.util import Chat, reflections
import subprocess
import webbrowser
import platform # To detect the operating system
import os # For general OS interactions
import pyautogui # For keyboard/mouse automation (install with: pip install pyautogui)

# Initialize the speech recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Get the operating system for platform-specific commands
OS_NAME = platform.system()

# Function to speak text
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen for audio input
def listen():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source) # Adjust for ambient noise
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=5) # Add timeout
        try:
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            speak("Sorry, I did not understand that.")
            return None
        except sr.RequestError:
            print("Could not request results from Google Speech Recognition service.")
            speak("I'm having trouble connecting to the speech recognition service.")
            return None
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            speak("An unexpected error occurred while listening.")
            return None

# --- Device Control Functions ---

def open_application(app_name):
    """Opens a specified application based on the OS."""
    speak(f"Opening {app_name}")
    try:
        if OS_NAME == "Windows":
            # Common applications
            if "notepad" in app_name.lower():
                subprocess.Popen(["notepad.exe"])
            elif "calculator" in app_name.lower():
                subprocess.Popen(["calc.exe"])
            elif "paint" in app_name.lower():
                subprocess.Popen(["mspaint.exe"])
            elif "word" in app_name.lower():
                # For Microsoft Office apps, you might need the full path or check common executables
                subprocess.Popen(["start", "winword"], shell=True) # 'start' command can help find it
            elif "excel" in app_name.lower():
                subprocess.Popen(["start", "excel"], shell=True)
            elif "powerpoint" in app_name.lower():
                subprocess.Popen(["start", "powerpnt"], shell=True)
            else:
                # Attempt to open as a general executable or via shell
                subprocess.Popen(["start", app_name], shell=True)
        elif OS_NAME == "Darwin": # macOS
            subprocess.Popen(["open", "-a", app_name])
        elif OS_NAME == "Linux":
            # Common Linux commands, may vary by distro
            subprocess.Popen([app_name.lower().replace(" ", "")]) # Remove spaces for command names
        else:
            speak(f"I don't know how to open applications on your operating system: {OS_NAME}")
            return False
        speak(f"{app_name} opened.")
        return True
    except FileNotFoundError:
        speak(f"Sorry, I couldn't find {app_name}. Please make sure it's installed or specify the full path.")
        return False
    except Exception as e:
        speak(f"An error occurred while trying to open {app_name}.")
        print(f"Error opening application: {e}")
        return False

def open_website(url):
    """Opens a specified URL in the default web browser."""
    speak(f"Opening {url} in your browser.")
    try:
        if not url.startswith("http://") and not url.startswith("https://"):
            url = "https://" + url # Assume HTTPS if not specified
        webbrowser.open(url)
        speak("Done.")
        return True
    except Exception as e:
        speak(f"Sorry, I couldn't open {url}.")
        print(f"Error opening website: {e}")
        return False

def adjust_volume(level_change):
    """Adjusts system volume (simplified example, depends heavily on OS and setup)."""
    if OS_NAME == "Windows":
        try:
            # Requires a third-party tool like 'nircmdc.exe' or 'sndvol32.exe' with parameters,
            # or more complex COM object interaction. For simplicity, we'll use a placeholder.
            # A more robust solution for Windows might involve 'pycaw' or similar.
            speak("Volume control on Windows is complex without external tools. I can't directly adjust it yet.")
            # Example with nircmdc (if installed and in PATH):
            # if "increase" in level_change:
            #     subprocess.run(["nircmdc.exe", "changesysvolume", "5000"]) # Increase by 5000 (out of 65535)
            # elif "decrease" in level_change:
            #     subprocess.run(["nircmdc.exe", "changesysvolume", "-5000"]) # Decrease
            # else:
            #     speak("Please specify 'increase' or 'decrease' volume.")
            # return True
        except Exception as e:
            speak("Could not adjust volume on Windows.")
            print(f"Volume error: {e}")
            return False
    elif OS_NAME == "Darwin": # macOS
        try:
            if "increase" in level_change:
                subprocess.run(["osascript", "-e", "set volume output volume (get volume settings)'s output volume + 10"])
                speak("Volume increased.")
            elif "decrease" in level_change:
                subprocess.run(["osascript", "-e", "set volume output volume (get volume settings)'s output volume - 10"])
                speak("Volume decreased.")
            elif "mute" in level_change or "off" in level_change:
                subprocess.run(["osascript", "-e", "set volume output muted true"])
                speak("Volume muted.")
            elif "unmute" in level_change or "on" in level_change:
                subprocess.run(["osascript", "-e", "set volume output muted false"])
                speak("Volume unmuted.")
            else:
                speak("Please specify 'increase', 'decrease', 'mute', or 'unmute' volume.")
                return False
            return True
        except Exception as e:
            speak("Could not adjust volume on macOS.")
            print(f"Volume error: {e}")
            return False
    elif OS_NAME == "Linux":
        try:
            # Uses 'amixer' for ALSA, adjust for PulseAudio ('pactl') if needed
            if "increase" in level_change:
                subprocess.run(["amixer", "-D", "pulse", "sset", "Master", "5%+"])
                speak("Volume increased.")
            elif "decrease" in level_change:
                subprocess.run(["amixer", "-D", "pulse", "sset", "Master", "5%-"])
                speak("Volume decreased.")
            elif "mute" in level_change or "off" in level_change:
                subprocess.run(["amixer", "-D", "pulse", "sset", "Master", "mute"])
                speak("Volume muted.")
            elif "unmute" in level_change or "on" in level_change:
                subprocess.run(["amixer", "-D", "pulse", "sset", "Master", "unmute"])
                speak("Volume unmuted.")
            else:
                speak("Please specify 'increase', 'decrease', 'mute', or 'unmute' volume.")
                return False
            return True
        except FileNotFoundError:
            speak("Amixer or pactl command not found. Cannot adjust volume.")
            return False
        except Exception as e:
            speak("Could not adjust volume on Linux.")
            print(f"Volume error: {e}")
            return False
    else:
        speak("I cannot adjust volume on this operating system.")
        return False

def search_web(query):
    """Performs a web search using the default browser."""
    speak(f"Searching the web for {query}.")
    try:
        webbrowser.open(f"https://www.google.com/search?q={query}")
        speak("Here are the search results.")
        return True
    except Exception as e:
        speak(f"Sorry, I couldn't perform the web search for {query}.")
        print(f"Search error: {e}")
        return False

def shutdown_computer():
    """Initiates computer shutdown."""
    speak("Are you sure you want to shut down the computer? Say 'yes' to confirm or 'no' to cancel.")
    confirmation = listen()
    if confirmation and "yes" in confirmation.lower():
        speak("Shutting down the computer now. Goodbye!")
        if OS_NAME == "Windows":
            os.system("shutdown /s /t 1") # /s for shutdown, /t 1 for 1 second delay
        elif OS_NAME == "Darwin" or OS_NAME == "Linux":
            os.system("sudo shutdown -h now") # Requires sudo password, might not work without configuration
            # Alternatively for macOS: subprocess.Popen(["osascript", "-e", 'tell app "System Events" to shut down'])
        else:
            speak("I cannot shut down on this operating system.")
    else:
        speak("Shutdown cancelled.")

def restart_computer():
    """Initiates computer restart."""
    speak("Are you sure you want to restart the computer? Say 'yes' to confirm or 'no' to cancel.")
    confirmation = listen()
    if confirmation and "yes" in confirmation.lower():
        speak("Restarting the computer now. Goodbye!")
        if OS_NAME == "Windows":
            os.system("shutdown /r /t 1") # /r for restart, /t 1 for 1 second delay
        elif OS_NAME == "Darwin" or OS_NAME == "Linux":
            os.system("sudo reboot") # Requires sudo password, might not work without configuration
            # Alternatively for macOS: subprocess.Popen(["osascript", "-e", 'tell app "System Events" to restart'])
        else:
            speak("I cannot restart on this operating system.")
    else:
        speak("Restart cancelled.")

# --- Enhanced Command Processing ---

# Add more sophisticated patterns for device control
# The order of patterns matters; more specific patterns should come before general ones.
pairs = [
    # Basic Chatbot
    ['hi|hello|hey', ['Hello!', 'Hi there!', 'Greetings!']],
    ['how are you?', ['I am fine, thank you for asking!', 'Doing well, how about you?']],
    ['what is your name?', ['I am your voice assistant, here to help.']],
    ['bye|exit|quit|goodbye', ['Goodbye!', 'See you later!', 'Farewell!']],

    # Device Control
    ['open (notepad|calculator|paint|word|excel|powerpoint)', ['_open_application']],
['open youtube', ['_open_youtube']], # Open YouTube
    ['open (.+) application', ['_open_application']], # General application opening
    [r'go to (.+)\.com', ['_open_website']],
    ['open website (.+)', ['_open_website']],
    ['search for (.+)', ['_search_web']],
    ['find (.+) on the web', ['_search_web']],
    ['increase volume', ['_adjust_volume_increase']],
    ['decrease volume', ['_adjust_volume_decrease']],
    ['mute volume', ['_adjust_volume_mute']],
    ['unmute volume', ['_adjust_volume_unmute']],
    ['shut down computer', ['_shutdown_computer']],
    ['restart computer', ['_restart_computer']],
    ['what can you do', ['I can open applications and websites, search the web, adjust volume, and perform basic system commands. How can I help you?']],
    ['tell me the time', ['_get_time']], # Placeholder for new function
    ['tell me the date', ['_get_date']], # Placeholder for new function
    ['where am i', ['_get_location']], # Placeholder for new function
]

# Custom reflections (optional, but good for more natural responses)
custom_reflections = {
    "i am": "you are",
    "i was": "you were",
    "i feel": "you feel",
    "i have": "you have",
    "i would": "you would",
    "my": "your",
    "you are": "I am",
    "you were": "I was",
    "you have": "I have",
    "you would": "I would",
    "your": "my",
}

# Function to respond to user input - now handles custom actions
def open_youtube():
    """Opens YouTube in the default browser."""
    speak("Opening YouTube.")
    webbrowser.open("https://www.youtube.com")
    speak("YouTube is now open.")


def respond(command):
    if not command:
        return

    command_lower = command.lower()
    chat = Chat(pairs, custom_reflections) # Use custom reflections

    # --- Direct Command Handling (more explicit for device control) ---
    # This allows more precise control over how commands are interpreted
    # before falling back to the NLTK Chat patterns.

    if "open" in command_lower and "application" in command_lower:
        app_name = command_lower.replace("open", "").replace("application", "").strip()
        if app_name:
            open_application(app_name)
            return
    elif "open" in command_lower and ".com" in command_lower:
        url = command_lower.split("open", 1)[1].strip()
        open_website(url)
        return
    elif "go to" in command_lower:
        url = command_lower.split("go to", 1)[1].strip()
        open_website(url)
        return
    elif "search for" in command_lower:
        query = command_lower.split("search for", 1)[1].strip()
        search_web(query)
        return
    elif "find" in command_lower and "on the web" in command_lower:
        query = command_lower.replace("find", "").replace("on the web", "").strip()
        search_web(query)
        return
    elif "increase volume" in command_lower:
        adjust_volume("increase")
        return
    elif "decrease volume" in command_lower:
        adjust_volume("decrease")
        return
    elif "mute volume" in command_lower:
        adjust_volume("mute")
        return
    elif "unmute volume" in command_lower:
        adjust_volume("unmute")
        return
    elif "shut down computer" in command_lower:
        shutdown_computer()
        return
    elif "restart computer" in command_lower:
        restart_computer()
        return
    elif "what is the time" in command_lower or "tell me the time" in command_lower:
        get_time()
        return
    elif "what is the date" in command_lower or "tell me the date" in command_lower:
        get_date()
        return
    elif "where am i" in command_lower or "my location" in command_lower:
        get_location()
        return
    elif "open youtube" in command_lower:
        open_youtube()
        return

    # Fallback to NLTK Chat for general conversations
    response = chat.respond(command_lower)
    if response:
        # Handle custom action flags within NLTK responses (less common with direct handling)
        if response == '_open_application':
            # This branch would require more complex NLTK pattern parsing or user input
            # For simplicity, direct handling above is preferred for device control
            speak("Which application would you like to open?")
            app_name_confirm = listen()
            if app_name_confirm:
                open_application(app_name_confirm)
        elif response == '_open_website':
            speak("Which website would you like to open?")
            url_confirm = listen()
            if url_confirm:
                open_website(url_confirm)
        elif response == '_search_web':
            speak("What would you like to search for?")
            query_confirm = listen()
            if query_confirm:
                search_web(query_confirm)
        elif response == '_adjust_volume_increase':
            adjust_volume("increase")
        elif response == '_adjust_volume_decrease':
            adjust_volume("decrease")
        elif response == '_adjust_volume_mute':
            adjust_volume("mute")
        elif response == '_adjust_volume_unmute':
            adjust_volume("unmute")
        elif response == '_shutdown_computer':
            shutdown_computer()
        elif response == '_restart_computer':
            restart_computer()
        elif response == '_get_time':
            get_time()
        elif response == '_get_date':
            get_date()
        elif response == '_get_location':
            get_location()
        else:
            speak(response) # Speak the regular chatbot response
    else:
        speak("I am not sure how to respond to that, or I don't have a command for it yet.")

# --- New functions for time, date, location (using current context) ---
import datetime

def get_time():
    current_time = datetime.datetime.now().strftime("%I:%M %p")
    speak(f"The current time is {current_time}.")

def get_date():
    current_date = datetime.datetime.now().strftime("%A, %B %d, %Y")
    speak(f"Today is {current_date}.")

def get_location():
    # As per your current context, the location is fixed.
    # For dynamic location, you'd need geolocation APIs (e.g., Google Maps API)
    # which require API keys and internet access.
    speak("Based on my current information, you are in T.Kavundampalayam, Tamil Nadu, India.")


# Main loop to run the voice assistant
def run_assistant():
    speak("Hello! How can I help you today?")
    while True:
        command = listen()
        if command:
            respond(command)
            if "bye" in command.lower() or "exit" in command.lower() or "quit" in command.lower():
                speak("Goodbye!")
                break

# Start the voice assistant
if __name__ == "__main__":
    run_assistant()
