import random
import time
import subprocess
import speech_recognition as sr
import webbrowser
import pyttsx3
import os
import email, getpass, imaplib, datetime
from bs4 import BeautifulSoup
import requests
import json
import psutil
import win32com.client

def recognize_speech_from_mic(recognizer, microphone):
    """Transcribe speech from recorded from `microphone`.

    Returns a dictionary with three keys:
    "success": a boolean indicating whether or not the API request was
               successful
    "error":   `None` if no error occured, otherwise a string containing
               an error message if the API could not be reached or
               speech was unrecognizable
    "transcription": `None` if speech could not be transcribed,
               otherwise a string containing the transcribed text
    """
    # check that recognizer and microphone arguments are appropriate type
    if not isinstance(recognizer, sr.Recognizer):
        raise TypeError("`recognizer` must be `Recognizer` instance")

    if not isinstance(microphone, sr.Microphone):
        raise TypeError("`microphone` must be `Microphone` instance")

    # adjust the recognizer sensitivity to ambient noise and record audio
    # from the microphone
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    # set up the response object
    response = {
        "success": True,
        "error": None,
        "transcription": None
    }

    # try recognizing the speech in the recording
    # if a RequestError or UnknownValueError exception is caught,
    #     update the response object accordingly
    try:
        response["transcription"] = recognizer.recognize_google(audio)
    except sr.RequestError:
        # API was unreachable or unresponsive
        response["success"] = False
        response["error"] = "API unavailable"
    except sr.UnknownValueError:
        # speech was unintelligible
        response["error"] = "Unable to recognize speech"

    return response


def process_mailbox(M):

    date = (datetime.date.today() - datetime.timedelta(1)).strftime("%d-%b-%Y")
    date='23-Jun-2018'
    #print(" Date is ", date)
    result, data = M.uid('search', None, "ALL", '(SENTON %s)' % date) # search all email and return uids
    if result == 'OK':
        for num in data[0].split():
            result, data = M.uid('fetch', num, '(RFC822)')
            if result == 'OK':
                email_message = email.message_from_bytes(data[0][1])    # raw email text including headers
                print(" ************************************************************")
                print('    From:' + email_message['From'])
                print('    Subject:' + email_message['Subject'])
                print('    Date:' + email_message['Date'])
                print(" ************************************************************")

    M.close()
    M.logout()

def fetchLatestMails() :
    user = "tushar.kadu1291@gmail.com"
    pwd = "mailjol@123"
    M = imaplib.IMAP4_SSL('imap.gmail.com')

    try:
        M.login(user, pwd)
        rv, mailboxes = M.list()
        if rv == 'OK':
            rv, data = M.select("INBOX")
            if rv == 'OK':
                #print (" Fetching todays mail...\n")
                process_mailbox(M) # ... do something with emails, see below ...
                M.close()
                M.logout()

    except imaplib.IMAP4.error as e :
        print (" MAIL ERROR!!! ",e)

def getNewsDetails() :
    
    # For URL
    url = "http://zeenews.india.com"
    page_link = url

    # fetch the content from url
    page_response = requests.get(page_link, timeout=5)

    # parse html
    page_content = BeautifulSoup(page_response.content, "html.parser")

    # Find all elements in body
    MainBody = page_content.find("div",{"class" : "view-content"})

    # Find all title details
    #productLinks = soup.findAll('a', attrs={'class' : 'on'})
    newsTitalDetails = MainBody.findAll('div', attrs={'class' : 'mini-con'})
   
    title_list = []
    count = 0
    print(" Latest  news details are \n ")
    # Get movie title details
    for headLines in newsTitalDetails:
        count = count + 1
        news = " " + str(count) + " " + headLines.a.text.strip()
        print(news)     



if __name__ == "__main__":
    
    # set the list of words, maxnumber of guesses, and prompt limit
    WORDS = ["apple", "banana", "grape", "orange", "mango", "lemon"]
    NUM_GUESSES = 100
    PROMPT_LIMIT = 102

    # create recognizer and mic instances
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()

    isStarted = "true" 
    while isStarted :
        # Sleep for 3 sec
        time.sleep(3)
        print("\n Oleg is ready to answer !!! (Say \"Start Execution\" to activate )")
        guess = recognize_speech_from_mic(recognizer, microphone)
        if guess["error"] :
            pass
        else :    
            #print("Guess is ", guess)
            # Listen voice
            startCommand = guess["transcription"].lower()
            # show the user the transcription
            print("    You said: {}".format(guess["transcription"]))
            
            # Start the execution only when someone is rady
            if "start execution" in startCommand:
                break

    # get a random word from the list
    word = random.choice(WORDS)

    # format the instructions string
    instructions = (
        "\n IMP :- You have {n} tries to give commands.\n"
    ).format(words=', '.join(WORDS), n=NUM_GUESSES)

    # show instructions and wait 3 seconds before starting the game
    print(instructions)
    time.sleep(3)

    for i in range(NUM_GUESSES):
        # get the guess from the user
        # if a transcription is returned, break out of the loop andcls
        #     continue
        # if no transcription returned and API request failed, break
        #     loop and continue
        # if API request succeeded but no transcription was returned,
        #     re-prompt the user to say their guess again. Do this up
        #     to PROMPT_LIMIT times
        for j in range(PROMPT_LIMIT):
            print('\n Guess {}. Speak Loudly!'.format(i+1))
            guess = recognize_speech_from_mic(recognizer, microphone)
            if guess["transcription"]:
                break
            if not guess["success"]:
                break
            print("I didn't catch that. What did you say?\n")

        # if there was an error, stop the game
        if guess["error"]:
            print("ERROR: {}".format(guess["error"]))
            break

        # show the user the transcription
        print(" You said: {}".format(guess["transcription"]))

        # determine if guess is correct and if any attempts remain
        guess_is_correct = guess["transcription"].lower() == word.lower()
        user_has_more_attempts = i < NUM_GUESSES - 1

        # take original text of user
        userText = guess["transcription"].lower()


        # determine if the user has won the game
        # if not, repeat the loop if user has more attempts
        # if no attempts left, the user loses the game
 
        if "eclipse" in userText:
            print("Opening eclipse \n ")
            subprocess.call(['C:/Users/gs-1157/Pictures/eclipse-jee-mars-2-win32-x86_64/eclipse/eclipse.exe'])
            #break
    
        elif "notepad" in userText:
            print("   Opening notepad ++ \n ")
            subprocess.call(['C:/Program Files (x86)/Notepad++/notepad++.exe'])
            #break
    
        elif "youtube" in userText:    
            print("   Opening youtube \n ")             
            controller = webbrowser.get()
            controller.open("https://www.youtube.com/")
            #break
 
        elif "rakesh" in userText:    
            engine = pyttsx3.init()
            engine.say("  Rakesh issss Assss HOLE")
            engine.setProperty('rate',500)  
            engine.runAndWait()

        elif "restart" in userText:    

            if "simulator" in userText:    
                print("   Restarting simulator .. .. .. \n ")
                # Stop process
                PROCNAME = "java.exe"        
                simulator_path = "C:\\Users\\gs-1157\\workspace\\atprofscsimfw\\simulator"
                for proc in psutil.process_iter():
                    if proc.name() == PROCNAME:
                        p = psutil.Process(proc.pid)
                        #print("   Process Details Are ", p.cwd())
                        simulator_process_path = str(p.cwd())
                        # print(simulator_process_path)
                        if (simulator_path == simulator_process_path) :
                            print("   Killing process ", proc.pid, " Path ", simulator_process_path)
                            proc.kill()
                # Start process
                os.system("start /B start cmd.exe @cmd /k  \"cd C:/Users/gs-1157/workspace/atprofscsimfw/simulator && java -Xmx2g -jar atomiton.tqlengine.jar\"")
                #break    

            if "engine" in userText:    
                print("   Restarting device engine \n ")
                # Stop process
                device_engine_path = "C:\\Users\\gs-1157\\workspace\\atprofscitcisco\\cdp\\deviceEngine"
                PROCNAME = "java.exe"        
                for proc in psutil.process_iter():
                    if proc.name() == PROCNAME:
                        p = psutil.Process(proc.pid)
                        #print("   Process Details Are ", p.cwd())
                        device_engine_process_path = str(p.cwd())
                        # print(simulator_process_path)
                        if (device_engine_path == device_engine_process_path) :
                            print("   Killing process ", proc.pid, " Path ", device_engine_process_path)
                            proc.kill()
              # Start process
                os.system("start /B start cmd.exe @cmd /k  \"cd C:/Users/gs-1157/workspace/atprofscitcisco/cdp/deviceEngine && java -Xmx2g -jar atomiton.tqlengine.jar\"")
                #break    

        elif "start" in userText:    
            if "simulator" in userText:    
                print("   Starting simulator .. .. .. \n ")
                os.system("start /B start cmd.exe @cmd /k  \"cd C:/Users/gs-1157/workspace/atprofscsimfw/simulator && java -Xmx2g -jar atomiton.tqlengine.jar\"")
                #break    

            if "engine" in userText:    
                print("   Starting device engine \n ")
                os.system("start /B start cmd.exe @cmd /k  \"cd C:/Users/gs-1157/workspace/atprofscitcisco/cdp/deviceEngine && java -Xmx2g -jar atomiton.tqlengine.jar\"")
                #break    
 
        elif "mail" in userText:    
            print("   Following are the todays mail.. \n ")
            fetchLatestMails()
            print("\n\n ")
            time.sleep(6)
            #break    

        elif "news" in userText:    
            print("   Following are the todays news.. \n ")
            getNewsDetails()

        elif "stop" in userText:    
            if "simulator" in userText:
                print("   Stopping simulator process .. .. .. \n ")
                simulator_path = "C:\\Users\\gs-1157\\workspace\\atprofscsimfw\\simulator"
                PROCNAME = "java.exe"        
                for proc in psutil.process_iter():
                    if proc.name() == PROCNAME:
                        p = psutil.Process(proc.pid)
                        #print("   Process Details Are ", p.cwd())
                        simulator_process_path = str(p.cwd())
                        # print(simulator_process_path)
                        if (simulator_path == simulator_process_path) :
                            print("   Killing process ", proc.pid, " Path ", simulator_process_path)
                            proc.kill()
            if "engine" in userText:    
                print("   Stopping device engine process .. .. .. \n ")
                device_engine_path = "C:\\Users\\gs-1157\\workspace\\atprofscitcisco\\cdp\\deviceEngine"
                PROCNAME = "java.exe"        
                for proc in psutil.process_iter():
                    if proc.name() == PROCNAME:
                        p = psutil.Process(proc.pid)
                        #print("   Process Details Are ", p.cwd())
                        device_engine_process_path = str(p.cwd())
                        # print(simulator_process_path)
                        if (device_engine_path == device_engine_process_path) :
                            print("   Killing process ", proc.pid, " Path ", device_engine_process_path)
                            proc.kill()

        elif "terminate" in userText:    
            print("\n Terminating OLEG .. .. .. \n ")
            print("  .. .. ")
            print("  .. ")
            print("  END")
            break    
                        
        else:
            print(" Sorry, OLEG is not able to understand your command !\n")
            #break