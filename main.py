import json
import os
import subprocess
import webbrowser
import threading
import queue
import speech_recognition as sr
from datetime import datetime
import tkinter as tk
from tkinter import scrolledtext, filedialog, ttk
import pyttsx3
import requests
import win32com.client as wincl
from google import genai
from google.genai import types
from PIL import Image, ImageTk

class DennisAssistantUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dennis Assistant")
        self.root.geometry("800x600")
        self.root.configure(bg="#343541")
        
        # Set app icon if available
        try:
            self.root.iconbitmap("dennis_icon.ico")
        except:
            pass
        
        # Create the message queue for thread-safe communication
        self.message_queue = queue.Queue()
        self.processing = False
        
        # Initialize text-to-speech engine
        self.tts_engine = pyttsx3.init()
        self.speech_enabled = True
        
        # Initialize the speech recognizer
        self.recognizer = sr.Recognizer()
        self.is_listening = False
        
        # Create Dennis Assistant backend
        self.assistant = DennisAssistant(self.message_queue)
        
        self.setup_ui()
        self.check_message_queue()
    
    def setup_ui(self):
        # Create main frame
        main_frame = tk.Frame(self.root, bg="#343541")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create chat display area
        self.chat_frame = tk.Frame(main_frame, bg="#343541")
        self.chat_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a canvas with a frame inside it for the messages
        self.canvas = tk.Canvas(self.chat_frame, bg="#343541", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.chat_frame, orient="vertical", command=self.canvas.yview)
        self.messages_frame = tk.Frame(self.canvas, bg="#343541")
        
        # Configure the canvas
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create a window in the canvas for the messages frame
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.messages_frame, anchor="nw")
        
        # Configure the messages frame to expand to the width of the canvas
        self.messages_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # Add a welcome message
        self.add_message("assistant", "Hi, I'm Dennis! How can I help you today?")
        
        # Create input area
        input_frame = tk.Frame(main_frame, bg="#40414f", height=100)
        input_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        # Text input
        self.user_input = scrolledtext.ScrolledText(input_frame, height=3, bg="#40414f", fg="white", 
                                              wrap=tk.WORD, font=("Arial", 11), relief=tk.FLAT)
        self.user_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 5), pady=10)
        self.user_input.bind("<Return>", self.on_enter_pressed)
        self.user_input.bind("<Shift-Return>", lambda e: None)  # Allow Shift+Enter for new line
        self.user_input.focus_set()
        
        # Buttons frame
        buttons_frame = tk.Frame(input_frame, bg="#40414f")
        buttons_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10))
        
        # Speech input button
        self.mic_icon = tk.PhotoImage(file="mic_icon.png") if os.path.exists("mic_icon.png") else None
        mic_button_text = "" if self.mic_icon else "🎤"
        self.mic_button = tk.Button(buttons_frame, text=mic_button_text, image=self.mic_icon, 
                               compound=tk.LEFT, bg="#40414f", fg="white", relief=tk.FLAT,
                               activebackground="#565869", command=self.toggle_speech_input)
        self.mic_button.pack(side=tk.TOP, pady=5)
        
        # Send button
        send_button_text = "Send"
        self.send_button = tk.Button(buttons_frame, text=send_button_text, bg="#5c64f4", fg="white", 
                                relief=tk.FLAT, activebackground="#4954f2", command=self.send_message)
        self.send_button.pack(side=tk.TOP, pady=5)
        
        # Speech toggle button
        speech_toggle_text = "🔊" if self.speech_enabled else "🔇"
        self.speech_toggle = tk.Button(buttons_frame, text=speech_toggle_text, bg="#40414f", fg="white", 
                                  relief=tk.FLAT, activebackground="#565869", command=self.toggle_speech)
        self.speech_toggle.pack(side=tk.TOP, pady=5)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Configure the style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TScrollbar", background="#40414f", troughcolor="#343541", 
                       arrowcolor="white", bordercolor="#40414f")
    
    def on_canvas_configure(self, event):
        # Update the width of the frame inside the canvas when the canvas is resized
        self.canvas.itemconfig(self.canvas_frame, width=event.width)
    
    def add_message(self, sender, message):
        # Create frame for the message
        message_frame = tk.Frame(self.messages_frame, bg="#343541")
        message_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        
        # Set the background and text colors based on sender
        if sender == "user":
            bg_color, fg_color = "#40414f", "white"
            avatar_text = "You"
        else:
            bg_color, fg_color = "#444654", "white"
            avatar_text = "Dennis"
        
        # Create avatar label
        avatar_frame = tk.Frame(message_frame, bg=bg_color, width=50)
        avatar_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        avatar_label = tk.Label(avatar_frame, text=avatar_text[0], bg="#5c64f4" if sender == "assistant" else "#10a37f",
                                fg="white", width=2, height=1, font=("Arial", 12, "bold"))
        avatar_label.pack(pady=10)
        
        # Create message text
        msg_frame = tk.Frame(message_frame, bg=bg_color, padx=10, pady=10)
        msg_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        msg_label = tk.Label(msg_frame, text=message, justify=tk.LEFT, bg=bg_color, fg=fg_color,
                            font=("Arial", 11), wraplength=600)
        msg_label.pack(anchor=tk.W)
        
        # Auto-scroll to the bottom of the canvas
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)
    
    def send_message(self):
        """Send the message from the input field"""
        user_message = self.user_input.get("1.0", tk.END).strip()
        if user_message:
            self.user_input.delete("1.0", tk.END)
            self.add_message("user", user_message)
            
            # Show thinking status
            self.status_var.set("Dennis is thinking...")
            self.processing = True
            
            # Process the message in a separate thread
            threading.Thread(target=self.process_message, args=(user_message,), daemon=True).start()
    
    def process_message(self, message):
        """Process the user message in a background thread"""
        self.assistant.process_user_input(message)
    
    def check_message_queue(self):
        """Check for messages from the assistant thread"""
        try:
            while True:
                message_type, message = self.message_queue.get_nowait()
                
                if message_type == "assistant_response":
                    self.add_message("assistant", message)
                    self.processing = False
                    self.status_var.set("Ready")
                    
                    # Speak the response if speech is enabled
                    if self.speech_enabled:
                        threading.Thread(target=self.speak_response, args=(message,), daemon=True).start()
                
                elif message_type == "status":
                    self.status_var.set(message)
                
                self.message_queue.task_done()
        except queue.Empty:
            pass
        finally:
            # Schedule the next check
            self.root.after(100, self.check_message_queue)
    
    def toggle_speech(self):
        """Toggle speech output"""
        self.speech_enabled = not self.speech_enabled
        self.speech_toggle.config(text="🔊" if self.speech_enabled else "🔇")
    
    def speak_response(self, text):
        """Speak the response using TTS"""
        self.tts_engine.say(text)
        self.tts_engine.runAndWait()
    
    def toggle_speech_input(self):
        """Toggle speech recognition for input"""
        if self.is_listening:
            self.is_listening = False
            self.mic_button.config(bg="#40414f", activebackground="#565869")
            self.status_var.set("Speech input stopped")
        else:
            self.is_listening = True
            self.mic_button.config(bg="#ff4c4c", activebackground="#ff6b6b")
            self.status_var.set("Listening...")
            threading.Thread(target=self.listen_for_speech, daemon=True).start()
    
    def listen_for_speech(self):
        """Listen for speech input"""
        with sr.Microphone() as source:
            try:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=5)
                self.status_var.set("Processing speech...")
                
                text = self.recognizer.recognize_google(audio)
                self.user_input.delete("1.0", tk.END)
                self.user_input.insert(tk.END, text)
                
                # Automatically send the recognized speech
                self.root.after(100, self.send_message)
                
            except sr.WaitTimeoutError:
                self.status_var.set("No speech detected")
            except sr.UnknownValueError:
                self.status_var.set("Could not understand audio")
            except sr.RequestError as e:
                self.status_var.set(f"Error with speech recognition service; {e}")
            finally:
                self.is_listening = False
                self.mic_button.config(bg="#40414f", activebackground="#565869")
    
    def on_enter_pressed(self, event):
        """Handle Enter key press"""
        if not self.processing and not (event.state & 0x0001):  # 0x0001 is the shift state
            self.send_message()
            return "break"  # Prevents the default behavior (newline)


class DennisAssistant:
    def __init__(self, message_queue):
        self.message_queue = message_queue

    def process_user_input(self, user_input):
        # Get response from Gemini API
        response = self.generate_response(user_input)
        
        try:
            response_data = json.loads(response)
            
            if "response" in response_data:
                # Model-handled task - just display the response
                self.message_queue.put(("assistant_response", response_data["response"]))
            else:
                # Developer-handled task - process it
                result = self.handle_developer_task(response_data)
                self.message_queue.put(("assistant_response", result))
        except json.JSONDecodeError:
            self.message_queue.put(("assistant_response", "I couldn't process that request. Please try again."))

    def generate_response(self, user_input):
        client = genai.Client(
            api_key=("AIzaSyAY8iaAkR4LSfNPA8RntMPE4ZkoOWX9mRM"),
        )

        model = "gemini-2.5-flash-preview-04-17"
        
        system_instruction = """
      You are a Desktop Virtual Assistant named "Dennis". Your responsibilities include understanding user commands and responding in a natural, helpful manner. You operate in two modes:

Model-Handled Tasks (within your control):  
You can directly respond to the following tasks:  
- Thinking, reasoning, and simple conversation  
- Web searches  
- General queries that don’t require real-time system access or dynamic data  

For these, respond naturally and conversationally without referencing internal logic, code structures, or developer concepts.  

Developer-Handled Tasks (outside your control):  
You cannot perform system-level tasks or provide real-time data (like current date/time/weather). Instead, your role is to:  
- Extract the user’s intent  
- Identify the operation type (e.g., creation, deletion, opening)  
- Parse the name of the file, folder, application, or media  
- Determine the location/path if mentioned  
- Return a structured JSON-style object to guide backend implementation  
For model handled tasks, respond naturally. For developer-handled tasks, provide a JSON object with the following structure:
```json
{
  "response": "<natural_response>"
}
```
For developer-handled tasks, return a JSON object with the following structure:
Tasks in this category include:  
- open_app / 
- create_file / delete_file / open_file / close_file / read_file  
- create_folder / delete_folder / open_folder / close_folder  
- current_date / current_day / current_year / current_time  
- read_pdf / read_docx  
- play_music  
- get_weather / get_location / check_network / open_url  
For url decide the domain type (e.g., .com, .org, etc.) on your own.
For files, decide the file type (e.g., .txt, .pdf, etc.) on your own.
For developer-handled tasks, return a JSON object with the following structure:
JSON Output Format (only for developer-handled tasks):  
```json
{
  "task": "<task_name>",
  "operation": "<operation_type>",
  "name": "<entity_name>",
  "location": "<location_path>",
  "filetype": "<file_extension>"
}

        """
        
        contents = [
            types.Content(
                role="user",
                parts=[types.Part.from_text(text=user_input)],
            )
        ]
        
        generate_content_config = types.GenerateContentConfig(
            temperature=0.7,
            response_mime_type="application/json",
            system_instruction=[types.Part.from_text(text=system_instruction)],
        )

        response = client.models.generate_content(
            model=model,
            contents=contents,
            config=generate_content_config,
        )
        
        return response.text

    def handle_developer_task(self, task_data):
        task = task_data.get("task")
        operation = task_data.get("operation")
        name = task_data.get("name")
        location = task_data.get("location", None)
        filetype = task_data.get("filetype", None)
        
        try:
            if task == "open_app":
                return self.open_application(name)
            elif task == "create_file":
                return self.create_file(name, location, filetype)
            elif task == "create_folder":
                return self.create_folder(name, location)
            elif task == "open_folder":
                return self.open_folder(name)
            elif task == "open_url":
                return self.open_url(name)
            elif task == "open_file":
                return self.open_file(name, filetype)
            elif task == "play_music":
                return self.play_music(name)
            elif task == "read_pdf":
                return self.read_pdf(name, location)
            elif task == "read_docx":
                return self.read_docx(name, location)
            elif task == "current_date":
                return self.get_current_date()
            elif task == "current_time":
                return datetime.now().strftime("%H:%M:%S")
            elif task == "get_weather":
                return self.get_weather(location)
            elif task == "get_location":
                return self.get_location()
            elif task == "close_folder":
                return self.close_folder(name)
            elif task == "play_music":
                return self.play_music(name)
            else:
                return f"I received your request to {operation} {task} '{name}', but this feature isn't implemented yet."
        except Exception as e:
            return f"Sorry, I encountered an error while processing your request: {str(e)}"

    # Implementations of all task methods with parameters

    def open_file(self, name, filetype=None):
        try:
            # Show searching message
            self.message_queue.put(("status", "Searching for file, please wait..."))
            
            found_paths = []
            lock = threading.Lock()
            threads = []
    
            def search_drive(drive, target_name, target_type):
                drive_path = f"{drive}:\\"
                skip_dirs = [
                    'Windows', 
                    'System32', 
                    'SysWOW64', 
                    'Program Files', 
                    'Program Files (x86)',
                    '$Recycle.Bin',
                    'System Volume Information'
                ]
    
                if os.path.exists(drive_path):
                    for root, dirs, files in os.walk(drive_path, topdown=True):
                        try:
                            # Skip system directories
                            dirs[:] = [d for d in dirs if not d.startswith('$') 
                                     and d not in skip_dirs 
                                     and not any(skip_dir in root for skip_dir in skip_dirs)]
                            
                            # Check files with the specific extension
                            target_filename = f"{target_name}{target_type}" if target_type else target_name
                            if target_filename in files:
                                full_path = os.path.join(root, target_filename)
                                with lock:
                                    found_paths.append(full_path)
                                    return  # Exit after finding first match
                        except PermissionError:
                            continue
                        except Exception:
                            continue
    
            # Search all available drives
            for drive in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                if os.path.exists(f"{drive}:"):
                    t = threading.Thread(target=search_drive, args=(drive, name, filetype))
                    t.start()
                    threads.append(t)
    
            # Wait for all searches to complete
            for t in threads:
                t.join()
    
            if found_paths:
                os.startfile(found_paths[0])
                return f"Found and opened file: {os.path.basename(found_paths[0])}"
            else:
                return f"Could not find file: {name}{f'{filetype}' if filetype else ''}"
    
        except Exception as e:
            return f"Error while trying to open file: {str(e)}"

    def open_application(self, app_name):
        try:
            # Get all available drives
            drives = [f"{d}:\\" for d in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' if os.path.exists(f"{d}:\\")]
            
            # Common executable extensions
            extensions = ['.exe', '.msi', '.bat', '.cmd']
            
            # Common install directories relative to drive root
            common_dirs = [
                'Program Files',
                'Program Files (x86)',
                f'Users\\{os.getenv("USERNAME")}\\AppData\\Local',
                f'Users\\{os.getenv("USERNAME")}\\AppData\\Local\\Programs',
                f'Users\\{os.getenv("USERNAME")}\\AppData\\Roaming',
                'Windows',
                'Windows\\System32',
                'Windows\\SysWOW64',
                'Windows\\System',
                # dynamic path from differnt folder and 
            ]
    
            # First try direct launch
            try:
                if app_name.lower() == "vscode":
                    app_name = "Code.exe"
                subprocess.Popen(app_name)
                return f"Launched {app_name}"
            except:
                pass
    
            # Function to check if a filename matches our search
            def is_match(filename, search_term):
                name_lower = filename.lower()
                search_lower = search_term.lower()
                # Check if the search term appears as a whole word
                return (search_lower == name_lower or 
                       search_lower == name_lower.replace('.exe', '') or
                       search_lower in name_lower.split('-') or
                       search_lower in name_lower.split('_') or
                       search_lower in name_lower.split(' '))
    
            # Search in registry for installed applications
            try:
                import winreg
                registry_paths = [
                    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"),
                    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths")
                ]
                
                for hkey, reg_path in registry_paths:
                    try:
                        reg_key = winreg.OpenKey(hkey, reg_path, 0, winreg.KEY_READ)
                        for i in range(winreg.QueryInfoKey(reg_key)[0]):
                            try:
                                key_name = winreg.EnumKey(reg_key, i)
                                if is_match(key_name, app_name):
                                    try:
                                        app_key = winreg.OpenKey(reg_key, key_name)
                                        path = winreg.QueryValue(app_key, None)
                                        if path and os.path.exists(path):
                                            subprocess.Popen(path)
                                            return f"Launched {app_name} from registry"
                                    except:
                                        continue
                            except:
                                continue
                    except:
                        continue
            except:
                pass
    
            # Search in drives if registry search failed
            for drive in drives:
                for base_dir in common_dirs:
                    search_dir = os.path.join(drive, base_dir)
                    if not os.path.exists(search_dir):
                        continue
                        
                    for root, dirs, files in os.walk(search_dir):
                        # Check each file in the current directory
                        for file in files:
                            if any(file.lower().endswith(ext) for ext in extensions):
                                if is_match(file, app_name):
                                    try:
                                        app_path = os.path.join(root, file)
                                        subprocess.Popen(app_path)
                                        return f"Launched {file} from {root}"
                                    except:
                                        continue
    
            # Try Windows Run as last resort
            try:
                subprocess.run(['cmd', '/c', 'start', app_name], check=True)
                return f"Launched {app_name} using Windows Run"
            except:
                return f"Could not find or launch {app_name}. Please verify the application name."
    
        except Exception as e:
            return f"Error while trying to launch {app_name}: {str(e)}"

    def create_file(self, name, location=None, filetype=None):
        try:
            path = os.path.join(location or os.getcwd(), f"{name}.{filetype}" if filetype else name)
            
            with open(path, 'w') as f:
                pass
            return f"Created file: {os.path.basename(path)}"
        except Exception as e:
            return f"Could not create file: {str(e)}"

    def create_folder(self, name, location=None):
        try:
            path = os.path.join(location or os.getcwd(), name)
            os.makedirs(path, exist_ok=True)
            return f"Created folder: {name}"
        except Exception as e:
            return f"Could not create folder: {str(e)}"

    def open_folder(self, name):
        try:
            # Call the search function and show searching message
            self.message_queue.put(("status", "Searching for folder, please wait..."))
            
            # Search for the folder using the search_drive method from earlier
            found_paths = []
            lock = threading.Lock()
            threads = []
    
            def search_drive(drive, target_name):
                drive_path = f"{drive}:\\"
                skip_dirs = [
                    'Windows', 
                    'System32', 
                    'SysWOW64', 
                    'Program Files', 
                    'Program Files (x86)',
                    '$Recycle.Bin',
                    'System Volume Information'
                ]
    
                if os.path.exists(drive_path):
                    for root, dirs, files in os.walk(drive_path, topdown=True):
                        try:
                            dirs[:] = [d for d in dirs if not d.startswith('$') 
                                     and d not in skip_dirs 
                                     and not any(skip_dir in root for skip_dir in skip_dirs)]
                            
                            if target_name == os.path.basename(root):
                                with lock:
                                    found_paths.append(root)
                                    return  # Exit after finding first match
                        except PermissionError:
                            continue
                        except Exception:
                            continue
    
            # Search all available drives
            for drive in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                if os.path.exists(f"{drive}:"):
                    t = threading.Thread(target=search_drive, args=(drive, name))
                    t.start()
                    threads.append(t)
    
            # Wait for all searches to complete
            for t in threads:
                t.join()
    
            if found_paths:
                os.startfile(found_paths[0])
                return f"Found and opened folder: {name}"
            else:
                return f"Could not find folder: {name}"
    
        except Exception as e:
            return f"Error while trying to open folder: {str(e)}"

    def open_url(self, url):
        try:
            # if not url.startswith(('http://', 'https://')):
            #   url = f"https://{url}"
            webbrowser.open(url)
            return f"Opened {url} in your browser"
        except Exception as e:
            return f"Could not open URL: {str(e)}"

    def play_music(self, song_name):
        try:
            # Show searching message
            self.message_queue.put(("status", "This feature is not implemented yet"))
            return "Music playback functionality is not implemented yet. Please try again later."
        except Exception as e:
            return f"Could not play music: {str(e)}"

    def read_pdf(self, name, location=None):
        try:
            path = os.path.join(location or os.getcwd(), f"{name}.pdf")
            os.startfile(path)
            return f"Opened PDF: {name}"
        except Exception as e:
            return f"Could not read PDF: {str(e)}"

    def read_docx(self, name, location=None):
        try:
            path = os.path.join(location or os.getcwd(), f"{name}.docx")
            os.startfile(path)
            return f"Opened document: {name}"
        except Exception as e:
            return f"Could not read document: {str(e)}"

    def get_current_date(self):
        today = datetime.now().strftime("%A, %B %d, %Y")
        return f"Today's date is {today}"

    def get_weather(self, location=None):
        try:
            # OpenWeatherMap API configuration
            api_key = "8d2de98e089f1c28e1a22fc19a24ef04"  # Free API key
            loc = location or "Karachi"  # Default location if none provided
    
            # Clean up location data if it comes as a dict
            if isinstance(loc, dict):
                loc = loc.get("city", "Nairobi")
    
            # Make API request
            url = f"http://api.openweathermap.org/data/2.5/weather?q={loc}&appid={api_key}&units=metric"
            response = requests.get(url, timeout=10)
            data = response.json()
    
            if response.status_code == 200:
                # Extract weather information
                temp = data['main']['temp']
                humidity = data['main']['humidity']
                desc = data['weather'][0]['description']
                wind_speed = data['wind']['speed']
                city = data['name']
                country = data['sys']['country']
    
                # Format response
                weather_info = (
                    f"Weather in {city}, {country}:\n"
                    f"🌡️ Temperature: {temp:.1f}°C\n"
                    f"💧 Humidity: {humidity}%\n"
                    f"🌥️ Conditions: {desc.capitalize()}\n"
                    f"💨 Wind Speed: {wind_speed} m/s"
                )
                return weather_info
            else:
                return f"Could not get weather for {loc}: {data.get('message', 'Unknown error')}"
        
        except requests.Timeout:
            return "Weather service request timed out. Please try again."
        except requests.RequestException as e:
            return f"Network error while fetching weather data: {str(e)}"
        except Exception as e:
            return f"Could not get weather: {str(e)}"

    def get_location(self):
        try:
            # Use IP-based geolocation API
            response = requests.get('https://ipinfo.io/json', timeout=5)
            if response.status_code == 200:
                data = response.json()
                city = data.get('city', 'Unknown')
                region = data.get('region', 'Unknown')
                country = data.get('country', 'Unknown')
                loc = data.get('loc', '')
                return f"Your approximate location is: {city}, {region}, {country} (Coordinates: {loc})"
            else:
                return "Could not determine your location (API error)."
        except Exception as e:
            return f"Could not get location: {str(e)}"

    def close_folder(self, folder_name):
        try:
            # This is complex on Windows - would need to find window by title
            return f"I would close the '{folder_name}' folder if this was fully implemented"
        except Exception as e:
            return f"Could not close folder: {str(e)}"


if __name__ == "__main__":

    
    root = tk.Tk()
    app = DennisAssistantUI(root)
    root.mainloop()

# obj = DennisAssistant(None)
# # obj.open_file("Ct-22025-CCN-Lab9", "pdf")
# response = obj.generate_response("play music ishq")
# print(response)

