import win32com.client
import speech_recognition as sr
import webbrowser
import openai
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("how can i help you")
        r.adjust_for_ambient_noise(source, duration=0.2)
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='en-in')
            print(f"user said: {query}\n")
            return query.lower()
        except Exception as e:
            print("some error occurred")
            return ""


sites = [
    ["youtube", "https://www.youtube.com/"],
    ["google", "https://www.google.com/"],
    ["facebook", "https://www.facebook.com/"],
    ["instagram", "https://www.instagram.com/"],
    ["twitter", "https://www.twitter.com/"],
    ["linkedin", "https://www.linkedin.com/"],
    ["github", "https://www.github.com/"],
    ["stackoverflow", "https://stackoverflow.com/"],
    ["reddit", "https://www.reddit.com/"],
    ["wikipedia", "https://www.wikipedia.org/"],
    ["spotify", "https://www.spotify.com/"],
    ["netflix", "https://www.netflix.com/"],
    ["daraz", "https://www.daraz.pk/"],
    ["iobm lms", "https://lms.iobm.edu.pk/moodle/login/index.php"],
    ["iobm smartz", "https://smartz.iobm.edu.pk/StudentPortal/Student/EDU_EBS_STU_Login.aspx"]
]

while True:
    print("Listening...")
    text = takecommand()
    if "stop the program" in text:
        print("Exiting...")
        break

    for site in sites:
        if f"open {site[0]}".lower() in text.lower():
            webbrowser.open(site[1])
            break
    if "what is the time" in text or "what time is it" in text:
        time = datetime.datetime.now().strftime("%I:%M:%S %p")
        print(f"The time is {time}")
        speaker.Speak(f"The time is {time}")
        break

    if "what is the date" in text or "what date is it" in text:
        date = datetime.datetime.now().strftime("%Y-%m-%d")
        print(f"The date is {date}")
        speaker.Speak(f"The date is {date}")
        break

    if "what day is it" in text or "what day is todays" in text:
        day = datetime.datetime.now().strftime("%A")
        print(f"The day today is {day}")
        speaker.Speak(f"the day today is {day}")
        break
