import pyttsx3
import speech_recognition as sr
from datetime import date
import time
import webbrowser
import datetime
from pynput.keyboard import Key, Controller
import pyautogui
import sys
import os
from os import listdir
from os.path import isfile, join
import smtplib
# import wikipedia
import Gesture_Controller
#import Gesture_Controller_Gloved as Gesture_Controller
import app
from threading import Thread
import re
import screen_brightness_control as sbc
import requests
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import win32com.client as win32


# -------------Object Initialization---------------
today = date.today()
r = sr.Recognizer()
keyboard = Controller()
engine = pyttsx3.init('sapi5')
engine = pyttsx3.init()
voices = engine.getProperty('voices')
# print("Available voices:")
# for i, voice in enumerate(voices):
#     print(f"Index {i}: {voice.name} ({voice.languages})")
engine.setProperty('voice', voices[0].id)  # Select a voice by index (0 for first, 1 for second, etc.)

# ----------------Variables------------------------
file_exp_status = False
files =[]
path = ''
is_awake = True  #Bot status

# ------------------Functions----------------------
def reply(audio):
    app.ChatBot.addAppMsg(audio)

    print(audio)
    engine.say(audio)
    engine.runAndWait()


def wish():
    time.sleep(3)  # Add a 3-second delay
    hour = int(datetime.datetime.now().hour)

    if hour>=0 and hour<12:
        reply("Good Morning!")
    elif hour>=12 and hour<18:
        reply("Good Afternoon!") 
    else:
        reply("Good Evening!")  
        
    reply("Hi, this is SAM, your personal desktop assistant. How may i help you")

# Set Microphone parameters
with sr.Microphone() as source:
        r.energy_threshold = 500 
        r.dynamic_energy_threshold = False
        r.adjust_for_ambient_noise(source, duration=1)
# Audio to String
def record_audio():

    with sr.Microphone() as source:
        r.pause_threshold = 0.8
        voice_data = ''
        audio = r.listen(source, phrase_time_limit=5)

        try:
            voice_data = r.recognize_google(audio)
        except sr.RequestError:
            reply('Sorry my Service is down. Plz check your Internet connection')
            pass
        except sr.UnknownValueError:
            print('cant recognize')
            pass
        return voice_data.lower()

def set_volume(volume_level):
    """Set the system-wide volume level (0 to 100)."""
    if 0 <= volume_level <= 100:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        # Set volume level as a scalar (0.0 to 1.0)
        volume.SetMasterVolumeLevelScalar(volume_level / 100.0, None)
        reply(f"System volume set to {volume_level}%")
    else:
        reply("Please say a volume level between 0 and 100.")

def process_voicecommand(command):
    """Process the command to adjust volume."""
    match = re.search(r"set volume to (\d+)|increase volume to (\d+)", command)
    if match:
        # Extract volume level from the command
        volume = int(match.group(1) or match.group(2))
        set_volume(volume)
    else:
        reply("Sorry, I couldn't recognize a valid volume command.")

def set_brightness(brightness_level):
    """Set the brightness level (0 to 100)"""
    if 0 <= brightness_level <= 100:
        sbc.set_brightness(brightness_level)
        reply(f"Brightness set to {brightness_level}%")
    else:
        reply("Please say a brightness level between 0 and 100.")

def process_lightcommand(command):
    """Process the command to adjust brightness"""
    match = re.search(r"set brightness to (\d+)|increase brightness to (\d+)", command)
    if match:
        # Extract brightness level from the command
        brightness = int(match.group(1) or match.group(2))
        set_brightness(brightness)
    else:
        reply("Sorry, I couldn't recognize a valid brightness command.")

def fetch_joke():
    url = "https://v2.jokeapi.dev/joke/Any?type=single"  # API endpoint for single-line jokes
    try:
        response = requests.get(url)
        data = response.json()
        
        if 'joke' in data:
            joke = data['joke']
        else:
            joke = "I couldn't fetch a joke. Please try again later."
            return joke
    except requests.exceptions.RequestException as e:
        return "Sorry, I couldn't fetch a joke. Please try again later."

def search_file(query):
    """Function to search for a file in Windows Search bar."""
    if query:
        reply(f"Finding {query}")
        
        # Simulate pressing the Win key to open the search bar
        pyautogui.press("win")
        time.sleep(1)  # Wait for the search bar to appear

        # Type the query into the search bar
        pyautogui.typewrite(query)
        pyautogui.press("enter")  # Press Enter to open the search result
    else:
        reply("No search term provided.")

def show_capabilities():
    reply("I am sam, your desktop assistant. I can help you control your desktop dynamically with hand gestures and voice commands!")
    time.sleep(1)
    reply("Here are a few things I can do for you:")
    time.sleep(1)
    reply("1. Manage your files and folders, like opening or organizing them.")
    reply("2. Provide real-time weather updates.")
    reply("3. Control your desktop with hand gestures for tasks like mouse movement and clicking.")
    reply("4. Play your favorite music.")
    reply("5. Write and save documents in Microsoft Word using your voice.")
    reply("6. Search the internet or fetch information for you.")
    time.sleep(1)
    reply("Just tell me what you need, and I’ll be ready to assist!")


def get_weather():
        time.sleep(2)  # Simulate a delay for fetching weather data
        
        # Define location (can be changed to user input or a default location)
        city = "Lahore"  # Change to desired city
        api_key = "257011c089a859910e4d4e594a44ead3" 
        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
        
        try:
            # Send request to OpenWeatherMap API
            response = requests.get(url)
            data = response.json()

            # Check if the request was successful
            if response.status_code == 200:
                condition = data['weather'][0]['description']  # Weather condition (e.g., sunny, cloudy)
                temperature = data['main']['temp']  # Temperature in Celsius
                
                # Send the real weather report
                reply(f"The current weather in {city} is {condition} with a temperature of {temperature}°C. Have a great day!")
            else:
                reply("Sorry, I couldn't fetch the weather data. Please try again later.")
        except Exception as e:
            reply(f"Error fetching weather data: {str(e)}")

def time_now():
    current_time = datetime.datetime.now().strftime("%H:%M")  # Hour:Minute format
    reply(f"The current time is: {current_time}")


def handle_right_sam():
    engine.say("Starting dictation now.")
    engine.runAndWait()
    start_dictation()

def start_dictation():
    global doc
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    # Access Word application
    word = win32.Dispatch("Word.Application")
    word.Visible = True  # Make sure Word is visible

    # Open existing document or create a new one
    if word.Documents.Count > 0:
        doc = word.Documents.Item(1)  # Use the first open document
    else:
        doc = word.Documents.Add()  # Create a new document if none are open

    engine.say("Ready to write. Start speaking.")
    engine.runAndWait()

    with mic as source:
        recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise

        while True:
            try:
                print("Listening for speech...")
                audio = recognizer.listen(source, timeout=5)  # Listen for audio
                speech = recognizer.recognize_google(audio).lower()  # Convert speech to text
                print(f"You said: {speech}")

                write_to_word(speech)

                # Check for specific commands
                if "sam save" in speech or "sams save" in speech or "save" in speech or "sam-save" in speech:
                    engine.say("Please provide a name for the document.")
                    engine.runAndWait()
                    audio = recognizer.listen(source, timeout=10)  # Listen for the file name
                    try:
                        file_name = recognizer.recognize_google(audio).strip().replace(" ", "_")
                        file_path = f"C:\\Users\\LENOVO\\Documents\\{file_name}.docx"
                        save_document(file_path)
                    except sr.UnknownValueError:
                        engine.say("Sorry, I didn't catch the file name. Please try again.")
                        engine.runAndWait()
                    except sr.RequestError:
                        engine.say("Could not request results; check your internet connection.")
                        engine.runAndWait()
                
                
                elif "sam close" in speech:
                    if close_and_save(source):
                        break  # Stop dictation

            except sr.UnknownValueError:
                reply("Sorry, I didn't catch that.")
            except sr.RequestError:
                reply("Could not request results; check your internet connection.")

is_bold = False
def write_to_word(text):
    global doc, is_bold

    # Skip processing if the document isn't initialized
    if not doc:
        return
    
    # If the word "sam" is in the text, handle commands
    if "sam" in text.lower():
        # Handle specific commands
        if "sam period" in text.lower():
            doc.Range().InsertAfter(". ")  # Add a period and a space
            return
        
        elif "bold" in text.lower():
            # Activate bold mode
            is_bold = True
            engine.say("Bold mode activated.")
            engine.runAndWait()
            return
        
        elif "sam unbold" in text.lower():
            # Deactivate bold mode
            is_bold = False
            engine.say("Bold mode deactivated.")
            engine.runAndWait()
            return
    
    # For regular dictation: capitalize and add the text
    text = text.capitalize()

    # Set the text formatting (bold or normal)
    if is_bold:
        # Create a range for the text and apply bold
        range = doc.Range(doc.Content.End - 1, doc.Content.End)
        range.Text = text + " "  # Add space after the text
        range.Font.Bold = True
    else:
        doc.Range().InsertAfter(text + " ")

    is_bold = False


def save_document(file_path):
    global doc
    if doc:
        doc.SaveAs(file_path)  # Save the document at the specified path
        engine.say("Document saved")
        engine.runAndWait()
    else:
        engine.say("No document to save.")
        engine.runAndWait()

def close_and_save(source):
    global doc
    if doc:
        # Check if the document has been saved by checking if it has a name
        if doc.Name == "Document1":  # Default name indicates it hasn't been saved
            engine.say("You haven't saved the document yet. Please provide a name.")
            engine.runAndWait()
            audio = r.listen(source, timeout=10)  # Listen for the file name
            try:
                file_name = r.recognize_google(audio).strip().replace(" ", "_")
                file_path = f"C:\\Users\\LENOVO\\Documents\\{file_name}.docx"
                save_document(file_path)
            except sr.UnknownValueError:
                engine.say("Sorry, I didn't catch the file name. Please try again.")
                engine.runAndWait()
                return False
            except sr.RequestError:
                engine.say("Could not request results; check your internet connection.")
                engine.runAndWait()
                return False

        # Save and close the document
        doc.Save()
        doc.Close()
        engine.say("Document saved and closed.")
        engine.runAndWait()
        return True
    else:
        engine.say("No document to save.")
        engine.runAndWait()
        return False



# This is a code snippet that adds a new feature to the existing code. 
# It allows the user to ask sam to repeat the last message.

def repeat_last_message():
    if app.ChatBot.last_message:
        reply(app.ChatBot.last_message)
    else:
        reply("No last message to repeat.")

# Executes Commands (input: string)
def respond(voice_data):
    global file_exp_status, files, is_awake, path
    print(voice_data)
    voice_data.replace('sam','')
    app.eel.addUserMsg(voice_data)

    if is_awake==False:
        if 'hey sam' in voice_data:
            is_awake = True
            reply('Hello! How can I assist you today?')
        return

    # STATIC CONTROLS
    elif 'hello' in voice_data:
        wish()
        return

    elif 'what is your name' in voice_data:
        reply('My name is sam!')
        return

    elif 'date' in voice_data:
        reply(today.strftime("%B %d, %Y"))
        return
    
    elif 'what time is it' in voice_data:
        time_now()
        return
        
    elif 'google' in voice_data:
        reply('Searching for ' + voice_data.split('google')[1])
        url = 'https://google.com/search?q=' + voice_data.split('google')[1]
        try:
            webbrowser.get().open(url)
            reply('This is what I found Sir')
        except:
            reply('Please check your Internet')
        return
    
    elif "what can you do" in voice_data:
        show_capabilities()
        return
    
    elif 'tell me a joke' in voice_data:
        joke = fetch_joke()
        reply(joke)
        return
    
    elif ('bye' in voice_data) or ('by' in voice_data):
        reply("Good bye Sir! Have a nice day.")
        is_awake = False
        return

    elif ('exit' in voice_data) or ('terminate' in voice_data):
        if Gesture_Controller.GestureController.gc_mode:
            Gesture_Controller.GestureController.gc_mode = 0
        app.ChatBot.close()
        #sys.exit() always raises SystemExit, Handle it in main loop
        sys.exit()
        return
    
    elif ("what's the weather like" in voice_data.lower()) or ('what is the weather today' in voice_data.lower()):
        reply("Let me check the weather for you. Just a moment...")
        get_weather()
        return

    
    elif ("thank you" in voice_data):
        reply("You're so welcome! I'm always here to help!")
        return


    # DYNAMIC CONTROLS
    elif any(keyword in voice_data for keyword in ['launch gesture recognition', 'launch hand control', 'launch hand mouse']):
        if Gesture_Controller.GestureController.gc_mode:
            reply('Gesture control has been activated.')
        else:
            gc = Gesture_Controller.GestureController()
            t = Thread(target = gc.start)
            t.start()
            reply('Launched Successfully')
        return

    elif any(keyword in voice_data for keyword in ['stop gesture', 'stop hand control', 'stop hand mouse', 'top hand control', 'top hand', 'top geture', 'top hand control']):
        if Gesture_Controller.GestureController.gc_mode:
            Gesture_Controller.GestureController.gc_mode = 0
            reply('Gesture control has been deactivated.')
        else:
            reply('Gesture recognition is already inactive')
        return
        
    elif 'copy' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('c')
            keyboard.release('c')
        reply('Copied')
        return
          
    elif 'page' in voice_data or 'pest'  in voice_data or 'paste' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('v')
            keyboard.release('v')
        reply('Pasted')
        return
        
    # File Navigation (Default Folder set to C://)
    elif 'list' in voice_data:
        counter = 0
        path = 'C:\\Users\\LENOVO\\Desktop'
        files = listdir(path)
        filestr = ""
        for f in files:
            counter += 1
            print(str(counter) + ':  ' + f)
            filestr += str(counter) + ':  ' + f + '<br>'
        file_exp_status = True
        reply('Here are your files. Let me know if you need anything specific.')
        app.ChatBot.addAppMsg(filestr)
        return

    elif file_exp_status == True:
        counter = 0   
        if 'open' in voice_data:
            if 'open' in voice_data:
                if isfile(join(path,files[int(voice_data.split(' ')[-1])-1])):
                    os.startfile(path + files[int(voice_data.split(' ')[-1])-1])
                    file_exp_status = False
                else:
                    try:
                        index = int(voice_data.split(' ')[-1]) - 1  # Get the index of the file/folder to open
                        if 0 <= index < len(files):  # Check if the index is within valid range
                            selected_file = files[index]
                            selected_path = join(path, selected_file)

                            if isfile(selected_path):  # If it's a file, open it
                                os.startfile(selected_path)
                                reply(f'Opened {selected_file} successfully')
                            else:  # If it's a directory, list its contents
                                path = selected_path
                                files = listdir(path)
                                filestr = ""
                                counter = 0  # Ensure counter is reset here
                                for f in files:
                                    counter += 1
                                    filestr += str(counter) + ':  ' + f + '<br>'
                                    print(str(counter) + ':  ' + f)
                                reply(f'Opened folder {selected_file} successfully')
                                app.ChatBot.addAppMsg(filestr)
                        else:
                            reply('Invalid file or folder index.')
                    except ValueError:
                        reply('Please specify a valid number to open the file or folder.')
                    except Exception as e:
                        reply(f'An error occurred: {str(e)}')
                        print(f"Error: {str(e)}")
                                                    
        if 'back' in voice_data:
            if path == 'C:\\Users\\LENOVO\\Desktop':
                reply('This is the root directory. Cannot go back further.')
            else:
                # Go up one directory
                path = os.path.dirname(path)
                files = listdir(path)
                filestr = ""
                counter = 0  # Ensure counter is reset here
                for f in files:
                    counter += 1
                    filestr += str(counter) + ':  ' + f + '<br>'
                    print(str(counter) + ':  ' + f)
                reply(f'Back to {path}')
                app.ChatBot.addAppMsg(filestr)
        return
    
    elif ("set volume to" in voice_data) or ("increase volume to" in voice_data):
        process_voicecommand(voice_data)
        return
    
    elif ('set brightness to' in voice_data):
        process_lightcommand(voice_data)
        return

    elif ("set volume to" in voice_data) or ("increase volume to" in voice_data):
        process_voicecommand(voice_data)
        return
    
    elif ('set brightness to' in voice_data):
        process_lightcommand(voice_data)
        return
    
    elif 'sam right' in voice_data:
        handle_right_sam()
        return
    

    elif 'song' in voice_data or 'play' in voice_data:
        keyword = 'play' if 'play' in voice_data else 'song'
        query = voice_data.split(keyword, 1)[1].strip()  # Get the query after 'play' or 'song'

        reply('Searching song ' + query)
        url = 'https://youtube.com/search?q=' + query

        try:
            webbrowser.get().open(url)
            reply('This is what I found, Sir.')
        except Exception as e:
            reply('Please check your Internet connection.')
        return

    elif 'location' in voice_data:
        reply('Which place are you looking for ?')
        temp_audio = record_audio()
        app.eel.addUserMsg(temp_audio)
        reply('Locating...')
        url = 'https://google.nl/maps/place/' + temp_audio + '/&'
        try:
            webbrowser.get().open(url)
            reply('This is what I found Sir')
        except:
            reply('Please check your Internet')
        return
    
    elif "search" in voice_data:
        # Extract the part of the voice data after "find file"
        # Split and clean up the command
        if "search" in voice_data:
            search_term = voice_data.split("search", 1)[1].strip()  # Extract everything after 'find file'
            if search_term:
                search_file(search_term)  # Pass the extracted term to the search_file function
            else:
                reply("Please specify the file or application you want to find.")
        return
    
    elif "repeat" in voice_data:
        if app.ChatBot.last_message:
            reply(app.ChatBot.last_message)
        else:
            reply("No last message to repeat.")
        return
    
    else:
        reply("Sorry, I didn't understand that. Please try again.")
    

# ------------------Driver Code--------------------

t1 = Thread(target = app.ChatBot.start)
t1.start()

# Lock main thread until Chatbot has started
while not app.ChatBot.started:
    time.sleep(0.5)

wish()
voice_data = None
while True:
    if app.ChatBot.isUserInput():
        #take input from GUI
        voice_data = app.ChatBot.popUserInput()
    else:
        #take input from Voice
        voice_data = record_audio()

    #process voice_data
    if any(keyword in voice_data for keyword in ['sam', 'sams', 'protn']):
        try:
            #Handle sys.exit()
            respond(voice_data)
        except SystemExit:
            reply("Exit Successfull")
            break
        except Exception as e:
            print(f"Exception raised while closing: {e}")
