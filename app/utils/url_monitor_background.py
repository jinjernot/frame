import requests
import time
from datetime import datetime
import threading
import json

from config import TEAMS_WEBHOOK_URL, URLS_TO_MONITOR

CHECK_INTERVAL_SECONDS = 300
TIMEOUT_SECONDS = 10

# Shared dictionary to store URL statuses
url_statuses = {}
status_lock = threading.Lock() # To safely update the shared dictionary

def send_teams_message(subject, body):
    """Sends a message to Microsoft Teams via Incoming Webhook."""
    if not TEAMS_WEBHOOK_URL or TEAMS_WEBHOOK_URL == "TEAMS_WEBHOOK":
        # print(f"[{datetime.now()}] Teams Webhook URL not configured. Skipping notification.") # Debug print
        return

    # print(f"Attempting to send Teams notification: {subject}") # Debug print

    message_payload = {
        "text": f"**{subject}**\n\n{body}"
    }

    try:
        response = requests.post(
            TEAMS_WEBHOOK_URL,
            headers={"Content-Type": "application/json"},
            data=json.dumps(message_payload)
        )
        response.raise_for_status() # Raise an exception for HTTP errors (4xx or 5xx)
        # print(f"[{datetime.now()}] Teams notification sent: {subject} (Status: {response.status_code})") # Debug print
    except requests.exceptions.RequestException as e:
        # Keep this print for actual errors
        print(f"[{datetime.now()}] Error sending Teams notification: {e}. "
              f"Response Text: {getattr(e.response, 'text', 'N/A')}")

def check_url_status(url):
    """
    Checks the HTTP status of a given URL.
    Returns True if the status is 200 (OK), False otherwise.
    """
    try:
        response = requests.get(url, timeout=TIMEOUT_SECONDS)
        if response.status_code == 200:
            return True, f"UP (Status: {response.status_code})"
        else:
            return False, f"DOWN (Status: {response.status_code})"
    except requests.exceptions.Timeout:
        return False, f"DOWN (Timeout after {TIMEOUT_SECONDS} seconds)"
    except requests.exceptions.ConnectionError:
        return False, "DOWN (Connection Error)"
    except requests.exceptions.RequestException as e:
        return False, f"DOWN (Request Error: {e})"

def monitor_urls_background_task():
    """
    Continuously monitors the specified URLs and updates their statuses in the shared dictionary.
    Sends notifications if any URL status changes.
    """
    # print(f"[{datetime.now()}] Starting URL monitor background task...") # Debug print

    # --- Send "Status Checker Online" message ---
    subject_online = "URL Status Checker Online"
    body_online = (f"The URL monitoring service has just started.\n"
                   f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    send_teams_message(subject_online, body_online)


    # Initialize statuses on first run
    with status_lock:
        for url in URLS_TO_MONITOR:
            is_up, message = check_url_status(url)
            url_statuses[url] = {"is_up": is_up, "message": message, "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            if not is_up:
                # print(f"[{datetime.now()}] Initial check: {url} is {message}") # Debug print
                subject = f"Initial Alert: URL DOWN - {url}"
                body = (f"The following URL was found down during initial check:\n\n"
                        f"URL: {url}\n"
                        f"Status: {message}\n"
                        f"Time: {datetime.now()}")
                send_teams_message(subject, body)
            else:
                print(f"[{datetime.now()}] Initial check: {url} is {message}") # Debug print

    while True:
        current_statuses = {}
        for url in URLS_TO_MONITOR:
            is_up, message = check_url_status(url)
            current_statuses[url] = {"is_up": is_up, "message": message, "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

            with status_lock:
                # Check for status changes
                previous_is_up = url_statuses.get(url, {}).get("is_up")

                if is_up != previous_is_up:
                    if is_up:
                        subject = f"URL RECOVERED - {url}"
                        body = (f"The following URL has recovered:\n\n"
                                f"URL: {url}\n"
                                f"Status: {message}\n"
                                f"Time: {datetime.now()}")
                        send_teams_message(subject, body)
                    else:
                        subject = f"URGENT: URL DOWN - {url}"
                        body = (f"The following URL is currently unreachable or returning an error:\n\n"
                                f"URL: {url}\n"
                                f"Status: {message}\n"
                                f"Time: {datetime.now()}")
                        send_teams_message(subject, body)

                # Update the shared status
                url_statuses[url] = current_statuses[url]

            # print(f"[{datetime.now()}] {url} is {current_statuses[url]['message']}") # Debug print

        # print(f"[{datetime.now()}] Next check in {CHECK_INTERVAL_SECONDS} seconds...") # Debug print
        time.sleep(CHECK_INTERVAL_SECONDS)

# This global dictionary will be imported by the Flask app to display statuses
# and accessed by the monitoring thread to update.
# A threading.Lock is used to ensure thread-safe access.
_url_monitor_thread = None

def start_monitor_thread():
    """Starts the URL monitoring in a separate daemon thread."""
    global _url_monitor_thread
    if _url_monitor_thread is None or not _url_monitor_thread.is_alive():
        _url_monitor_thread = threading.Thread(target=monitor_urls_background_task, daemon=True)
        _url_monitor_thread.start()
        # print(f"[{datetime.now()}] URL monitor background thread started.") # Debug print
    else:
        # print(f"[{datetime.now()}] URL monitor background thread already running.") # Debug print
        pass # Or log to a file instead of printing if needed for information
