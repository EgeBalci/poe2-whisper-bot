import sys
import requests
import logging
import time
import argparse
import re
import os
import psutil
import threading
import pythoncom
import win32com.client
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def find_poe_client_log():
    """Find the absolute path of the PathOfExile.exe or PathOfExileSteam.exe process."""
    try:
        for proc in psutil.process_iter(['name', 'exe']):
            if proc.info['name'] in ['PathOfExile.exe', 'PathOfExileSteam.exe']:
                 # Get the directory containing the exe
                exe_dir = os.path.dirname(proc.info['exe'])
                # Check for Client.txt in logs subdirectory
                client_log_path = os.path.join(exe_dir, "logs", "Client.txt")
                if os.path.isfile(client_log_path):
                    return client_log_path
        return None
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        logging.error("Error accessing process information")
        return None

def is_purchase_whisper(message):
    if "@From" in message and "I would like to buy your" in message:
        return True
    return False

def parse_purchase_whisper(message):
    """Parse a purchase whisper message and extract relevant information."""
    pattern = r'@From ([^:]+): Hi, I would like to buy your ([^$]+) listed for (\d+) (\w+) in \w+ \(stash tab "([^"]+)"; position: left (\d+), top (\d+)\)'
    match = re.search(pattern, message)
    if match:
        return {
            'sender': match.group(1),
            'item': match.group(2).strip(),
            'amount': int(match.group(3)),
            'currency': match.group(4),
            'tab': match.group(5),
            'position_left': int(match.group(6)),
            'position_top': int(match.group(7))
        }
    return None

def send_start_message(bot_token, chat_id):
    """Send a start message to the Telegram bot's chat."""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    message = """**🎮 Path of Exile 2 🎮**\n\nSuccessfully connected to POE2 Whisper bot\\! Your trade notifications from Path of Exile 2 will be sent to this chat\\.\n\nHappy trading, Exile\\! 💎"""
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "MarkdownV2"
    }
    try:
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            logging.info("Start message sent successfully")
        else:
            logging.error("Failed to send start message. Status: %s, Response: %s", 
                         response.status_code, response.text)
    except requests.exceptions.RequestException as e:
        logging.exception("Error sending start message: %s", e)
    except Exception as e:
        logging.exception("Unexpected error sending start message: %s", e)

def send_message_to_telegram(bot_token, chat_id, purchase_info):
    """Send a message to a Telegram bot's chat."""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    message = f"**🎮 Path of Exile 2 🎮**\n\n"
    message += f"👤 `{purchase_info['sender']}`\n📦 `{purchase_info['item']} `\n💰 `{purchase_info['amount']}/{purchase_info['currency']} `\n📍 `{purchase_info['tab']} - {purchase_info['position_left']}, {purchase_info['position_top']}`\n⏰ `{datetime.now().strftime('%H:%M:%S')}`"""
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "MarkdownV2"
    }
    try:
        response = requests.post(url, json=payload)
        if response.status_code != 200:
            logging.error("Failed to send message: %s, Status: %s, Response: %s", message, response.status_code, response.text)
    except requests.exceptions.RequestException as e:
        logging.exception("Error while sending message: %s", e)

def focus_poe_window():
    """Focus the Path of Exile window if it exists."""
    try:
        import win32gui
        import win32con

        def window_enum_handler(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if "Path of Exile" in window_title:
                    windows.append(hwnd)

        windows = []
        win32gui.EnumWindows(window_enum_handler, windows)

        if windows:
            # Get the first PoE window found
            hwnd = windows[0]
            
            # If window is minimized, restore it
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            # Bring window to front and focus it
            win32gui.SetForegroundWindow(hwnd)
            return True
        return False
    except ImportError:
        logging.error("win32gui module not found. Please install pywin32 package.")
        return False
    except Exception as e:
        logging.exception("Error focusing PoE window: %s", e)
        return False

def prevent_afk_state():
    """Periodically send 'x' key to PoE window to prevent AFK state."""
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        while True:
            if focus_poe_window():
                # Send 'x' key
                shell.SendKeys("x")
            time.sleep(60)  # Wait 5 minutes before next keystroke
    except ImportError:
        logging.error("win32com module not found. Please install pywin32 package.")
    except Exception as e:
        logging.exception("Error in anti-AFK thread: %s", e)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Monitor a file for Path of Exile trade whispers and forward them to Telegram.')
    parser.add_argument('bot_token', help='Telegram bot token')
    parser.add_argument('chat_id', help='Telegram chat ID')
    args = parser.parse_args()

    bot_token = args.bot_token
    chat_id = args.chat_id

    client_log_path = find_poe_client_log()
    logging.info("POE Client Log: %s", client_log_path)
    logging.info("Tailing file: %s... (Press Ctrl+C to stop)", client_log_path)
    logging.info("Script started with bot_token: %s, chat_id: %s, file_path: %s", bot_token, chat_id, client_log_path)
    send_start_message(bot_token, chat_id)
    
    # Start anti-AFK thread
    logging.info("Starting anti-AFK thread...")
    anti_afk_thread = threading.Thread(target=prevent_afk_state, daemon=True)
    anti_afk_thread.start()

    
    try:
        with open(client_log_path, "r") as file:
            # Move to the end of the file
            file.seek(0, 2)
            while True:
                line = file.readline()
                if line:
                    message = line.strip()
                    if message and is_purchase_whisper(message):
                        purchase_info = parse_purchase_whisper(message)
                        logging.info("New purchase whisper: %s", purchase_info)
                        send_message_to_telegram(bot_token, chat_id, purchase_info)
                else:
                    time.sleep(1)  # Wait for new lines
    except KeyboardInterrupt:
        logging.info("Script stopped by user.")
    except FileNotFoundError:
        logging.error("File not found: %s", client_log_path)
    except Exception as e:
        logging.exception("An error occurred: %s", e)
