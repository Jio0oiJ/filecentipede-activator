import os
import json
import psutil
import requests
import argparse
import logging
import pyperclip
import pygetwindow as gw
import pyautogui

from bs4 import BeautifulSoup
from datetime import datetime, timedelta, timezone, UTC as TIME_UTC
from pymsgbox import (
    alert as msg_alert,
    confirm as msg_confirm,
    YES_TEXT,
    NO_TEXT,
)


KEY_LANG = ["en_US", "zh_CN", "zh_TW", "ru_RU", "ko_KR"]
KEY_URL = f"http://filecxx.com/{KEY_LANG[0]}/activation_code.html"
AUTOMATIC_MODE = False


# configure Logging
def setup_logging(log_file=None):
    """Configures the logging system.

    Args:
        log_file (str, optional): Path to the log file. If provided, logging is enabled.
    """
    if log_file:
        logging.basicConfig(
            filename=log_file,
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )
    else:
        logging.disable(logging.CRITICAL)


# Window Utilities
class MessageBoxIcon:
    """
    Represents the possible icons to display in a message box.

    Note: The "Question" icon is deprecated and should be avoided.
    """

    ASTERISK = 64  # Lowercase 'i' in a circle
    ERROR = 16  # White 'X' in a red circle
    EXCLAMATION = 48  # Exclamation point in a yellow triangle
    HAND = 16  # Same as ERROR
    INFORMATION = 64  # Same as ASTERISK
    NONE = 0  # No symbol
    # QUESTION = 32  # DEPRECATED: Question mark in a circle
    STOP = 16  # Same as ERROR
    WARNING = 48  # Same as EXCLAMATION


def is_filecxx_running() -> bool:
    """
    Checks if FileCxx is running.

    Returns:
        bool: True if FileCxx is running, False otherwise.
    """
    for process in psutil.process_iter():
        try:
            if process.name().lower() == "fileu.exe":
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            logging.exception("Error while checking if FileCxx is running")
            break  # process is no longer active
    return False


def show_initial_messagebox() -> None:
    """
    Shows the initial message box.
    
    msg_alert(
        "Beware of Scams! This is free software.\nIf you paid for it, please refund immediately.",
        "FileCxx Automatic Activation Code",
        icon=MessageBoxIcon.INFORMATION,
    )
    """

# Main Functions
def parse_html_content(html_content) -> list:
    """
    Parses the HTML content and returns a list of activation codes.

    Args:
        html_content (str): The HTML content to parse.

    Returns:
        list: A list of activation codes or an empty list if no activation codes are found.
    """
    soup = BeautifulSoup(html_content, "html.parser")
    pre_element = soup.find("pre", {"id": "codes"})
    logging.debug(f"Found pre element: {pre_element}")

    if not pre_element:
        return []

    lines = pre_element.text.strip().split("\n")
    logging.debug(f"Found {len(lines)} lines")
    result = []
    current_date_range = None

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if " - " in line:
            current_date_range = line
        elif current_date_range:
            result.append({"date_range": current_date_range, "activation_code": line})
            current_date_range = None

    return result


def format_date(date_str) -> str:
    """
    Formats a date string in the format "YYYY-MM-DD HH:MM:SS".

    Args:
        date_str (str): The date string to format.

    Returns:
        str: The formatted date string.
    """
    return datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S").isoformat()


def convert_to_utc(date_string) -> datetime:
    """
    Converts a UTC+8 datetime string to UTC.

    Args:
        date_string (str): The datetime string to convert.

    Returns:
        datetime: The UTC datetime object.
    """
    date_obj = datetime.fromisoformat(date_string).replace(
        tzinfo=timezone(timedelta(hours=8))
    )
    utc_date = date_obj.astimezone(timezone.utc)
    return utc_date


def create_json_object(parsed_data) -> str:
    """
    Creates a JSON object from the parsed data.

    Args:
        parsed_data (object): The parsed data to create the JSON object from.

    Returns:
        str: The JSON object string.
    """
    formatted_data = []
    for item in parsed_data:
        start_date, end_date = item["date_range"].split(" - ")
        formatted_data.append(
            {
                "start_date": format_date(start_date),
                "end_date": format_date(end_date),
                "activation_code": item["activation_code"],
            }
        )
    return json.dumps(formatted_data, indent=2)


def build_cached_data() -> str | None:
    """
    Builds the cached data.

    Returns:
        str | None: The cached data or None if an error occurred.
    """
    try:
        response = requests.get(KEY_URL, headers={"User-Agent": "filecxx/2.82"})
        response.raise_for_status()

        logging.debug(f"Response status code: {response.status_code}")
        if response.status_code == 200:
            parsed_data = parse_html_content(response.text)
            json_object = create_json_object(parsed_data)
            return json_object

    except requests.exceptions.RequestException:
        logging.exception("Failed to connect to the server")
        msg_alert(
            "Failed to connect to the server. Please try again later.",
            "FileCxx Automatic Activation Code",
            icon=MessageBoxIcon.ERROR,
        )
        return None


def write_keys_to_file(keys_data) -> None:
    """
    Writes the keys to a file.

    Args:
        keys_data (str): The keys data to write to the file.

    Returns:
        None
    """
    try:
        with open("keys.json", "w") as f:
            f.write(keys_data)
    except Exception as e:
        logging.exception(f"Failed to write keys to file: {e}")
        msg_alert(
            "Failed to write keys to file. Check your file permissions.",
            "FileCxx Automatic Activation Code",
            icon=MessageBoxIcon.ERROR,
        )


def insert_key_automatically(correct_code) -> None:
    """
    Inserts the key automatically.

    Args:
        correct_code (str): The correct code to insert.

    Returns:
        None
    """
    if not is_filecxx_running():
        logging.error("FileCxx is not running")
        msg_alert(
            "FileCxx is not running. Please start FileCxx and try again.",
            "FileCxx Automatic Activation Code",
            icon=MessageBoxIcon.ERROR,
        )
        return

    # Grab activation window - Note: you add more language here please
    window_title_name = {
        "en": "File Centipede - Activation code",
    }
    window_list = gw.getWindowsWithTitle(window_title_name["en"])
    logging.debug(f"Window list: {window_list}")
    if len(window_list) == 0:
        logging.error("FileCxx Activation window not found")
        msg_alert(
            "Open FileCxx Activation window and try again.",
            "FileCxx Automatic Activation Code",
            icon=MessageBoxIcon.ERROR,
        )
        return

    window = window_list[0]
    logging.debug(f"Window: {window}")
    window.activate()

    # click on the activation window input text
    logging.debug(f"Attempting to click on {window}")
    pyautogui.click(x=window.left + 30, y=window.top + 95)

    # delete all text in the activation window
    logging.debug("Deleting all text in the activation window")
    pyautogui.hotkey("ctrl", "a")
    logging.debug("Deleting all text in the activation window")
    pyautogui.hotkey("backspace")

    # type the correct code
    logging.debug(f"Typing {correct_code}")
    pyautogui.typewrite(f"{correct_code}\n")


def main() -> None:
    global AUTOMATIC_MODE
    parser = argparse.ArgumentParser(description="FileCxx Automatic Activation Code")

    parser.add_argument(
        "-a",
        "--automatic",
        action="store_true",  # Store True if the flag is present, False otherwise
        help="Enable automatic mode",
    )
    parser.add_argument(
        "-l",
        "--log",
        metavar="<filename>",
        type=str,
        help="Enable logging and specify the log file name.",
    )
    args = parser.parse_args()
    AUTOMATIC_MODE = args.automatic
    setup_logging(args.log)

    show_initial_messagebox()

    cached_data = None
    if os.path.exists("./keys.json"):
       # overwrite_confirmation = msg_confirm(
       #     "keys.json already exists. Do you want to overwrite it?",
       #     "FileCxx Automatic Activation Code",
       #     buttons=(YES_TEXT, NO_TEXT),
       # )
       # if overwrite_confirmation == YES_TEXT:
       #     logging.debug("Overwriting keys.json...")
            temp_data = build_cached_data()
            if temp_data:
                write_keys_to_file(temp_data)
                cached_data = json.loads(temp_data)
    else:
        temp_data = build_cached_data()
        if temp_data:
            write_keys_to_file(temp_data)
            cached_data = json.loads(temp_data)

    # define utc+8 timezone
    current_time_utc = datetime.now(TIME_UTC)

    correct_code = None
    for item in cached_data:
        start_time = convert_to_utc(item["start_date"])
        end_time = convert_to_utc(item["end_date"])

        if current_time_utc >= start_time and current_time_utc <= end_time:
            correct_code = item["activation_code"]
            break

    if not correct_code:
        msg_alert(
            "No valid activation code found. Please try again later.",
            "FileCxx Automatic Activation Code",
            icon=MessageBoxIcon.ERROR,
        )
        return

    if AUTOMATIC_MODE:
        insert_key_automatically(correct_code)
        return
    else:
        with open("key.txt", "w") as f:
            f.write(correct_code)
        pyperclip.copy(correct_code)

  #      open_confirmation = msg_confirm(
  #          f"Your activation code is saved as key.txt\nDo you want to open ?",
  #          "FileCxx Automatic Activation Code",
  #          buttons=(YES_TEXT, NO_TEXT),
  #          icon=MessageBoxIcon.INFORMATION,
  #      )
  #      if open_confirmation == YES_TEXT:
  #      os.startfile("key.txt")
        msg_alert("已将文件蜈蚣的注册码复制到剪贴板")

if __name__ == "__main__":
    main()
