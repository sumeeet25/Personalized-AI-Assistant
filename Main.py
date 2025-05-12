import datetime
import os
import webbrowser
import win32com.client
import speech_recognition as sr
from openai import OpenAI

# Initialize Speech Engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# OpenAI Client (Make sure your API key is in openai.py)
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="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",
)

def say(text):
    """Converts text to speech"""
    speaker.Speak(text)

def takeCommand():
    """Captures voice input from the user"""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand. Please repeat.")
            return ""
        except sr.RequestError:
            print("Sorry, there was an issue with the speech recognition service.")
            return ""

def ask_openai(prompt):
    """Sends user input to OpenAI and returns a response"""
    try:
        response = client.chat.completions.create(
            model="deepseek/deepseek-r1:free",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"

if __name__ == '__main__':
    welcome_message = "Hello! I am your personalized AI assistant. How can I assist you today?"
    print(welcome_message)
    say(welcome_message)

    while True:
        query = takeCommand()

        # If no input, continue listening
        if not query:
            continue

        # Open websites
        websites = {
            "google": "https://www.google.com",
            "youtube": "https://www.youtube.com",
            "facebook": "https://www.facebook.com",
            "amazon": "https://www.amazon.com",
            "wikipedia": "https://www.wikipedia.org",
            "twitter": "https://www.twitter.com",
            "instagram": "https://www.instagram.com",
            "linkedin": "https://www.linkedin.com",
            "reddit": "https://www.reddit.com",
            "whatsapp": "https://web.whatsapp.com",
            "netflix": "https://www.netflix.com",
            "spotify": "https://www.spotify.com",
            "github": "https://www.github.com",
            "stackoverflow": "https://www.stackoverflow.com",
            "quora": "https://www.quora.com",
            "flipkart": "https://www.flipkart.com",
            "microsoft": "https://www.microsoft.com",
            "apple": "https://www.apple.com",
            "yahoo": "https://www.yahoo.com",
            "bing": "https://www.bing.com",
            "zoom": "https://www.zoom.us",
            "canva": "https://www.canva.com",
            "tesla": "https://www.tesla.com",
        }

        for site, url in websites.items():
            if f"open {site}" in query:
                say(f"Opening {site}...")
                webbrowser.open(url)
                break  # Prevents multiple websites from opening at once

        # Play music
        if "play music" in query:
            music_path = r"C:\\Users\\Legion\\Downloads\\हमर सथ शर रघनथ.mp3"
            os.startfile(music_path)

        # Tell time
        elif "the time" in query:
            time_now = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"The current time is {time_now}")

        # Open Notepad
        elif "open notepad" in query:
            os.startfile(r"C:\Users\Public\Desktop\Notepad++.lnk")

        # Ask OpenAI for answers
        elif "what is" in query or "who is" in query or "explain" in query:
            ai_response = ask_openai(query)
            print(f"AI Assistant: {ai_response}")
            say(ai_response)

        # Exit the assistant
        elif "exit" in query or "goodbye" in query:
            print("Goodbye! Have a great day!")
            say("Goodbye! Have a great day!")
            break
