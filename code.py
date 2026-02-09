import pyttsx3
import datetime
import os
import psutil
import socket
import time
import subprocess
import importlib.util

# For brightness and volume
if importlib.util.find_spec("screen_brightness_control"):
    import screen_brightness_control as sbc
else:
    sbc = None

try:
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except Exception:
    AudioUtilities = None
    IAudioEndpointVolume = None
    CLSCTX_ALL = None


# --- Initialize Text-to-Speech Engine ---
try:
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)  # female voice
    engine.setProperty('rate', 185)
except Exception as e:
    print(f"Error initializing pyttsx3: {e}")
    engine = None


def speak(text):
    """Speak and print text."""
    print(text)
    if engine:
        try:
            engine.say(text)
            engine.runAndWait()
        except Exception as e:
            print(f"TTS error: {e}")


# --- Feature 1: Time & Date ---
def get_time_and_date():
    now = datetime.datetime.now()
    time_now = now.strftime("%I:%M %p")
    date_now = now.strftime("%B %d, %Y, %A")
    speak(f"The current time is {time_now}")
    speak(f"Today's date is {date_now}")


# --- Feature 2: Battery Status ---
def get_battery_status():
    try:
        battery = psutil.sensors_battery()
        if battery is None:
            speak("Battery information is not available on this device.")
            return
        percent = battery.percent
        charging = "and charging" if battery.power_plugged else "and not charging"
        speak(f"The battery is at {percent} percent {charging}.")
        if percent < 20 and not battery.power_plugged:
            speak("Warning! Battery is low, please plug in your charger.")
    except Exception as e:
        speak(f"Error while checking battery status: {e}")


# --- Feature 3: System Information ---
def system_info():
    try:
        hostname = socket.gethostname()
        ip_address = socket.gethostbyname(hostname)
        speak(f"This system name is {hostname}.")
    except Exception:
        speak("Unable to fetch system information at the moment.")


# --- Feature 4: Internet Connectivity + WiFi Name ---
def get_network_info():
    """Check if internet is connected and tell Wi-Fi name."""
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        connected = True
    except OSError:
        connected = False

    if connected:
        # Try to get Wi-Fi name (Windows command)
        try:
            result = subprocess.run(
                ['netsh', 'wlan', 'show', 'interfaces'],
                capture_output=True, text=True
            )
            wifi_name = "Unknown Network"
            for line in result.stdout.splitlines():
                if "SSID" in line and "BSSID" not in line:
                    wifi_name = line.split(":", 1)[1].strip()
                    break
            speak(f"Connected to Wi-Fi network named {wifi_name}.")
        except Exception:
            speak("Connected to internet, but unable to fetch Wi-Fi name.")
    else:
        speak("No internet connection detected.")


# --- Feature 5: Brightness Level ---
def get_brightness():
    """Check current screen brightness."""
    if sbc is None:
        speak("Brightness control module is not installed.")
        return
    try:
        brightness = sbc.get_brightness(display=0)
        if isinstance(brightness, list):
            brightness = brightness[0]
        speak(f"Screen brightness is {brightness} percent.")
    except Exception:
        speak("Unable to fetch brightness information.")

# --- Feature 7: CPU and RAM Status ---
def get_system_performance():
    try:
        cpu_usage = psutil.cpu_percent(interval=1)
        ram_usage = psutil.virtual_memory().percent
        speak(f"CPU usage is {cpu_usage:.1f} percent.")
        speak(f"RAM usage is {ram_usage:.1f} percent.")
    except Exception as e:
        speak(f"Error checking performance: {e}")


# --- MAIN PROGRAM ---
if __name__ == "__main__":
    speak("Let's begin your startup report.")

    get_time_and_date()
    get_battery_status()
    system_info()
    get_network_info()
    get_brightness()
    get_system_performance()
    speak("Goodbye!")
    time.sleep(2)
