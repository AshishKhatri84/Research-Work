import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import wikipedia
import subprocess

# Default directory for searching files
SEARCH_DIRECTORY = "C:\\Users\\ASUS\\Downloads"  # Change this to your preferred folder

# Initialize TTS engine
engine = pyttsx3.init()
engine.setProperty('rate', 160)
engine.setProperty('volume', 1.0)

dark_mode = True  # Default theme

def speak(text):
    engine.say(text)
    engine.runAndWait()

def recognize_speech():
    threading.Thread(target=_recognize_speech, daemon=True).start()

def _recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        append_message("🎤 Listening...", "bot", clear_previous=True)
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        text = recognizer.recognize_google(audio)
        append_message(f"You: {text}", "user")
        process_command(text)
    except sr.UnknownValueError:
        append_message("🤔 I couldn't understand that. Please try again.", "bot")
    except sr.RequestError:
        append_message("⚠️ Google API is not responding. Check your internet connection.", "bot")

def process_command(text):
    text = text.lower().strip()

    responses = {
        "hello": "Hello! How can I assist you?",
        "how are you": "I'm doing great! Thanks for asking.",
        "your name": "I'm ChatBot, your voice assistant!",
        "bye": "Goodbye! Have a nice day. 😊",
        "thank you": "You're welcome! 😊",
        "who created you": "I was created by Ashish Khatri!",
        "what can you do": "I can recognize speech, open apps, search the web, and perform tasks!",
    }

    for key, response in responses.items():
        if key in text:
            append_message(response, "bot")
            speak(response)
            return

    if "open youtube and search for" in text:
        query = text.replace("open youtube and search for", "").strip()
        webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
        reply = f"🔍 Searching YouTube for '{query}'..."
    
    elif "open google and search for" in text:
        query = text.replace("open google and search for", "").strip()
        webbrowser.open(f"https://www.google.com/search?q={query}")
        reply = f"🔎 Searching Google for '{query}'..."
    
    elif "wikipedia search" in text:
        query = text.replace("wikipedia search", "").strip()
        try:
            summary = wikipedia.summary(query, sentences=2)
            reply = f"📖 Wikipedia Summary: {summary}"
        except wikipedia.exceptions.DisambiguationError as e:
            reply = f"❌ Multiple results found. Be more specific: {', '.join(e.options[:5])}."
        except wikipedia.exceptions.PageError:
            reply = f"❌ No results found for '{query}'."
        except Exception as e:
            reply = f"⚠️ An error occurred while searching Wikipedia: {str(e)}"
    
    elif "open calculator" in text:  # New feature added here
        open_application("calc", "Calculator")
    
    elif text.startswith("open file"):
        filename = text.replace("open file", "").strip()
        open_file_by_name(filename)
    
    elif "open a file" in text:
        append_message("📂 Please select a file to open...", "bot")
        open_file()
    
    elif "open notepad" in text:
        open_application("C:\\Windows\\System32\\notepad.exe", "Notepad")

    elif "open chrome" in text:
        open_application("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "Google Chrome")

    elif "open firefox" in text:
        open_application("C:\\Program Files\\Mozilla Firefox\\firefox.exe", "Mozilla Firefox")

    elif "open word" in text or "open microsoft word" in text:
        open_application("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE", "Microsoft Word")

    elif "open excel" in text or "open microsoft excel" in text:
        open_application("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE", "Microsoft Excel")

    else:
        reply = f"❌ I didn't understand that command: '{text}'. Try again."
    
    append_message(reply, "bot")
    speak(reply)

def send_text():
    user_text = user_input.get().strip()
    if user_text:
        append_message(f"You: {user_text}", "user")
        process_command(user_text)
        user_input.delete(0, tk.END)

def append_message(text, sender, clear_previous=False):
    chat_window.config(state=tk.NORMAL)
    if clear_previous:
        chat_window.delete("end-2l", tk.END)  # Remove the previous listening message
    chat_window.insert(tk.END, text + "\n\n", sender)
    chat_window.config(state=tk.DISABLED)
    chat_window.yview(tk.END)

def open_file():
    file_path = filedialog.askopenfilename(title="Select a file")
    if file_path:
        try:
            os.startfile(file_path)  # Opens the file using the default application
            append_message(f"📂 Opening file: {os.path.basename(file_path)}", "bot")
        except Exception as e:
            append_message(f"⚠️ Unable to open the file: {str(e)}", "bot")

def open_file_by_name(filename):
    """Search and open a file by name in the specified directory."""
    file_path = os.path.join(SEARCH_DIRECTORY, filename)
    if os.path.exists(file_path):
        try:
            os.startfile(file_path)
            append_message(f"📂 Opening file: {filename}", "bot")
        except Exception as e:
            append_message(f"⚠️ Unable to open the file: {str(e)}", "bot")
    else:
        append_message(f"❌ File '{filename}' not found in {SEARCH_DIRECTORY}.", "bot")

def open_application(app_name, app_display_name):
    """Attempt to open an installed application."""
    try:
        subprocess.Popen(app_name)  # Try to open the application
        append_message(f"🚀 Opening {app_display_name}...", "bot")
    except FileNotFoundError:
        append_message(f"⚠️ {app_display_name} not found on your system.", "bot")

def toggle_theme():
    global dark_mode
    dark_mode = not dark_mode
    bg_color = "#1e1e1e" if dark_mode else "#f5f5f5"
    fg_color = "white" if dark_mode else "black"
    input_bg = "#333" if dark_mode else "white"
    
    root.configure(bg=bg_color)
    chat_window.configure(bg=bg_color, fg=fg_color)
    user_input.configure(
        bg=input_bg, 
        fg=fg_color, 
        insertbackground=fg_color, 
        highlightthickness=0 if not dark_mode else 2,  # Remove border in light mode
        relief="flat"  # Ensure no border style in light mode
    )
    
    input_frame.configure(bg=bg_color)
    button_frame.configure(bg=bg_color)

    send_button.configure(bg="#28A745" if dark_mode else "#2196F3", fg="white")
    mic_button.configure(bg="#E91E63" if dark_mode else "#D81B60", fg="white")
    theme_button.configure(bg="#FF9800" if dark_mode else "#FFC107", fg="black")

# GUI Setup
root = tk.Tk()
root.title("🧠 AI Voice Chatbot")
root.geometry("500x600")
root.configure(bg="#1e1e1e" if dark_mode else "#ffffff")

# Chat Window
chat_window = scrolledtext.ScrolledText(root, state=tk.DISABLED, wrap=tk.WORD, height=20, font=("Arial", 12), bg="#252526", fg="white", bd=5, relief="flat")
chat_window.pack(pady=10, padx=10, fill=tk.BOTH)

# Welcome Message
welcome_message = (
    "👋 Welcome to the AI Voice Chatbot!\n\n"
    "- 🎤 Listen to voice commands.\n"
    "- 🔍 Search Google or YouTube.\n"
    "- 📖 Fetch summaries from Wikipedia.\n"
    "- 🌓 Switch between light/dark themes.\n"
    "- 📂 Open files or launch applications.\n\n"
)
append_message(welcome_message, sender="bot")

# Input Frame (Textbox + Buttons)
input_frame = tk.Frame(root, bg="#1e1e1e")
input_frame.pack(pady=5)

user_input = tk.Entry(input_frame, font=("Arial", 14), bd=5, relief="solid", bg="#333333", fg="white")
user_input.grid(row=0, column=0, padx=(0, 5), sticky="ew")
user_input.bind("<Return>", lambda event: send_text())

send_button = tk.Button(input_frame, text="📩 Send", command=send_text, font=("Arial", 14), width=8)
send_button.grid(row=0, column=1)

mic_button = tk.Button(input_frame, text="🎤 Speak", command=recognize_speech, font=("Arial", 14), width=8)
mic_button.grid(row=0, column=2)

button_frame = tk.Frame(root)
button_frame.pack(pady=5)

theme_button = tk.Button(button_frame, text="🌓 Theme Toggle", command=toggle_theme)
theme_button.pack()

toggle_theme()  # Apply initial theme settings

root.mainloop()
