#!/usr/bin/env python

"""
Module: ZoomRoomScript.py
Sends messages to Microsoft Teams via Power Automate when we have a Zoom Room go offline.
This info is taken directly from the Zoom API.
"""

import base64
import threading
from math import ceil
from typing import *
import time
import requests
import sys
import json
import pystray
import os
from PIL import Image
from datetime import datetime
from zoneinfo import ZoneInfo
import tkinter as tk
from tkinter import ttk

__author__ = "Ben Thomas"
__credits__ = ["Ben Thomas", "Krish Jain"]
__license__ = "MIT"
__version__ = "0.1.0"
__maintainer__ = "Ben Thomas"
__email__ = "me@benjaminbt.com"
__status__ = "Development"

MSPOWERAUTOMATE_LINK = os.getenv('MSPOWERAUTOMATE_LINK').replace("^", "")
LOG_PATH = 'zoom_room_events.jsonl'
BOSTON_TZ = ZoneInfo("America/New_York")

#------------------------------------------------------------------------------------------------------
ZOOM_CLIENT_ID = os.getenv('ZOOM_CLIENT_ID')
ZOOM_CLIENT_SECRET = os.getenv('ZOOM_CLIENT_SECRET')
ZOOM_ACCOUNT_ID = os.getenv('ZOOM_ACCOUNT_ID')
#------------------------------------------------------------------------------------------------------
offline_room_duration = {}
offline_rooms = []
num_of_offline_rooms = 0
num_of_rooms = 0
times_ran = 0
gui_root = None
gui_is_open = False
COL_NAME, COL_ID, COL_DEVICES, COL_START, COL_STOP, COL_ELAPSED, COL_DATE = range(7)
#------------------------------------------------------------------------------------------------------
def get_zoom_access_token():
    '''
    this function will obtain a Zoom Server-to-Server OAuth2.0 token.
    '''
    pair = f"{ZOOM_CLIENT_ID}:{ZOOM_CLIENT_SECRET}"
    b64 = base64.b64encode(pair.encode("ascii")).decode("ascii")
    headers = {
        "Authorization": f"Basic {b64}",
        "Content-Type": "application/json"
    }
    url = f"https://zoom.us/oauth/token?grant_type=account_credentials&account_id={ZOOM_ACCOUNT_ID}"
    r = requests.post(url, headers=headers)
    if r.status_code != 200:
        print(f"Failed to get token: {r.status_code} {r.text}", file=sys.stderr)
        sys.exit(1)
    return r.json().get("access_token")


def send_teams_message(message_text: str, devices: str, send_email: bool = False):
    '''
    this function will send a teams message, the payload for the message requires a message,
    the list of devices as a string, and a boolean send_email which determines if an email is sent or not.
    this function sends this information to Power Automate which does the actual teams message sending.
    '''
    url = (
        MSPOWERAUTOMATE_LINK
    )
    headers = {
        "Content-Type": "application/json"
    }
    payload = {
        "msg": message_text,
        "devices": devices,
        "send_email": send_email,
        "status": "OK"
   }
    r = requests.post(url, headers=headers, json=payload)
    if r.status_code == 202:
        log_message_sent("Message Status", "Message Sent")
    else:
        print(r.text)
        log_error("Message Sending Error", f"{r.status_code}")
        log_message_sent("Message Status", "Message failed to send")

def get_room(token: str) -> List[Dict[str, str]]: ##This was made @4:53p and is the right method
    """

    :param token:
    :return: List[Dict[str, str]]
    """
    '''
    this function calls the zoom room api and returns the full list containing each room and its information 
    '''
    url = "https://api.zoom.us/v2/rooms" #Calling the rooms API
    headers = {
        "Authorization": f"Bearer {token}" #Declaring my headers
    }
    r = requests.get(url, headers=headers) # Saving my request to a variable
    if r.status_code != 200:
        print(f"Failed to find any rooms: {r.status_code} {r.text}", file=sys.stderr) #If the status code happens to be anything else but successful (!200) then something went wrong and we just print an error message
    return r.json()# return the full json


def is_room_online(user_token: str = None) -> None:
    '''
    this function is the main logic and will check to see if a room has become offline, if it has, it will
    call send_teams_message() to send a message stating that the room is now offline. this code will then log
    that information into a csv file. should the room go back online, a followup teams message will be sent stating
    that the room is now online (will only occur if all devices inside the room are online), this message will also
    contain the total duration that the room was offline.
    '''
    global num_of_rooms, offline_room_name, times_ran
    while True:
        current_rooms = get_room(user_token)['rooms']
        for i in range(len(current_rooms)):
            num_of_rooms += 1
        for index in range(num_of_rooms):
            if current_rooms[index]['status'] == 'Offline':
                offline_room_name = current_rooms[index]['name']
                offline_room_id = current_rooms[index]['room_id']
                if not (index in offline_rooms):
                    # print("added room to offline rooms:", current_rooms[index]['name'])
                    offline_rooms.append(index) # adding offline room to global offline_rooms list
                    start_timer(current_rooms[index]['name'])
                    send_teams_message(f"{offline_room_name} is down\n", list_of_devices_in_string(detailed_list_of_devices(offline_room_id), index), send_email=True)
                    log_event("Room Down", offline_room_name, offline_room_id, csv_format_devices(detailed_list_of_devices(offline_room_id), index))


        for offline_room_index in offline_rooms:
            if current_rooms[offline_room_index]['status'] != 'Offline':
                all_devices_online = True
                devices = detailed_list_of_devices(current_rooms[offline_room_index]['room_id'])
                for device in devices:
                    if device.get('status') == "Offline":
                        all_devices_online = False
                        break
                if all_devices_online:
                    send_teams_message(f"{current_rooms[offline_room_index]['name']} is back online! {stop_timer(current_rooms[offline_room_index]['name'])}\n", list_of_devices_in_string(detailed_list_of_devices(current_rooms[offline_room_index]['room_id']), offline_room_index))
                    # print("removing room:", current_rooms[offline_room_index]['name'])
                    offline_rooms.pop(offline_rooms.index(offline_room_index)) # removing room from offline_rooms list as it is now online
                    log_event("Room Up", offline_room_name, offline_room_id, csv_format_devices(devices, index))
        num_of_rooms = 0
        # Periodically update tray title with device status summary to limit API calls
        times_ran += 1
        if times_ran % 12 == 0:  # roughly once per minute (loop sleeps 5s)
            try:
                online_rooms, offline_rooms_count, total_rooms = compute_room_status_counts(current_rooms)
                icon.title = f"{online_rooms} - {total_rooms} Rooms Online 🟢\n{offline_rooms_count} Rooms Offline 🔴"
            except Exception:
                # Keep previous title if any error occurs
                pass
        time.sleep(5)

'''
this function will return the list of devices with a given room parameter.
'''
def detailed_list_of_devices(roomId: str) -> Dict[str, str]:
    url = f"https://api.zoom.us/v2/rooms/{roomId}/devices"
    headers = {"Authorization": f"Bearer {zoom_token}"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print("Invalid roomID -- Please try again.")
        print(f"You're passing: {roomId}\nPlease change this to a proper roomID")
        sys.exit(1)

    return r.json().get('devices')


def compute_room_status_counts(rooms: List[Dict[str, Any]]) -> Tuple[int, int, int]:
    '''
    Compute total online/offline room counts from a list of rooms.
    Returns (online_rooms, offline_rooms, total_rooms)
    '''
    if not rooms:
        return 0, 0, 0
    total_rooms = len(rooms)
    offline_rooms_count = 0
    for room in rooms:
        if room.get('status') == 'Offline':
            offline_rooms_count += 1
    online_rooms_count = total_rooms - offline_rooms_count
    return online_rooms_count, offline_rooms_count, total_rooms

def compute_device_status_counts_for_rooms(rooms: List[Dict[str, Any]]) -> Tuple[int, int, int]:
    '''
    Compute total online/offline device counts across all provided rooms.
    Returns (online_devices, offline_devices, total_devices)
    '''
    online_devices = 0
    offline_devices = 0
    total_devices = 0
    for room in rooms:
        try:
            room_id = room.get('room_id')
            if not room_id:
                continue
            devices = detailed_list_of_devices(room_id)
            if not devices:
                continue
            for device in devices:
                total_devices += 1
                if device.get('status') == 'Offline':
                    offline_devices += 1
                else:
                    online_devices += 1
        except Exception:
            # Skip rooms that error out
            continue
    return online_devices, offline_devices, total_devices

def list_of_devices_in_string(list_of_devices_passed: Dict[str, str], index_of_room_to_preview: int):
    '''
    this function will convert the list of devices into a properly formatted string that can be sent to Teams.
    It takes in a dictionary[str: str] and iterates through it checking the status key:value pair and if the value ==
    Offline, then it adds that devices to the final return string in a formated way.
    '''
    return_string = f"The devices in {get_room(zoom_token)['rooms'][index_of_room_to_preview]['name']}:\n"
    for index in range(len(list_of_devices_passed)):
        if list_of_devices_passed[index].get('status') == "Offline":
            emoji = "❌"
        else:
            emoji = "✅"
        return_string += f"\n\n •{emoji} {list_of_devices_passed[index].get('device_type')} — {list_of_devices_passed[index].get('status')}"
    return return_string



def csv_format_devices(list_of_devices_passed: Dict[str, str], index_of_room_to_preview: int):
    '''
    this function will convert the list of devices into a properly formatted string that can be exported to a csv file
    '''
    return_string = ""
    for index in range(len(list_of_devices_passed)):
        return_string += f"{list_of_devices_passed[index].get('device_type')} - {list_of_devices_passed[index].get('status')}, "
    return return_string[:-2]

def start_timer(room: str):
    '''
    this function will start a timer for a given room name
    '''
    offline_room_duration[room] = {
        "start": time.time(),
        "stop": None,
        "elapsed": None
    }

def stop_timer(room: str):
    '''
    this function will stop a timer for a given room name and return the total duration the room was offline
    '''
    if room in offline_room_duration and offline_room_duration[room]["start"]:
        offline_room_duration[room]["stop"] = time.time()
        offline_room_duration[room]["elapsed"] = ceil(offline_room_duration[room]["stop"] - offline_room_duration[room]["start"])
        secs = offline_room_duration[room]["elapsed"] % 60 # ex: 465 -> 45 seconds
        mins = offline_room_duration[room]["elapsed"] // 60 # ex: 465 -> 7 minutes
        if mins >= 60:
            hours = mins // 60
            mins -= hours * 60
            if hours >= 24:
                days = hours // 24
                days -= days * 24
                return f"\n\nOffline Duration: {days} days, {hours} hrs, {mins} min and {secs}sec\n" # if longer than a day
            return f"\n\nOffline Duration: {hours} hrs, {mins} min and {secs} sec\n" # if longer than an hour
        return f"\n\nOffline Duration: {mins} min and {secs} sec\n" # if less than an hour
    return "\n\nOffline Duration: NULL\n" # error has occured

def log_event(event_type: str = "Null", room_name: str = "Null", room_id: str = "Null", devices: List[Dict[str, Any]] = "Null", **extra: Any) -> None:
    """
    Append a JSON object per line to LOG_PATH for SIEM ingestion | dedicated for room status going on or offline
    """
    record = {
        'timestamp': datetime.now(BOSTON_TZ).isoformat(),
        'event': event_type,
        'room_name': room_name,
        'room_id': room_id,
        'devices': devices
    }
    record.update(extra)
    with open(LOG_PATH, 'a') as f:
        f.write(json.dumps(record) + '\n')

def log_error(event_type: str = "Null", error_message: str = "Null", **extra: Any) -> None:
    """
    Append a JSON object per line to LOG_PATH for SIEM ingestion | dedicated for logging errors
    """
    record = {
        'timestamp': datetime.now(BOSTON_TZ).isoformat(),
        'event': event_type,
        'error_message': error_message,
    }
    record.update(extra)
    with open(LOG_PATH, 'a') as f:
        f.write(json.dumps(record) + '\n')

def log_message_sent(event_type: str = "Null", message_status: str = "Null", **extra: Any):
    """"
    Append a JSON object per line to LOG_PATH for SIEM ingestion | dedicated for sending Teams Messages
    """
    record = {
        'timestamp': datetime.now(BOSTON_TZ).isoformat(),
        'event': event_type,
        'message_status': message_status,
    }
    record.update(extra)
    with open(LOG_PATH, 'a') as f:
        f.write(json.dumps(record) + '\n')


def handle_exit(icon=None, item=None):
    icon.stop()
    log_error("shutdown_requested", "script was shut off")
    send_teams_message(message_text="⚠ Zoom Room Monitoring System is Offline ⚠", devices="Please check the program on NEBAV02")
    sys.exit(0)


def open_gui(icon=None, item=None):
    """
    Open the GUI window on demand from the tray menu. Creates the window if needed,
    otherwise brings it to the foreground.
    """
    global gui_root, gui_is_open
    try:
        if gui_root is not None and gui_is_open:
            try:
                gui_root.deiconify()
                gui_root.lift()
                gui_root.focus_force()
                return
            except Exception:
                pass

        threading.Thread(target=_launch_gui, daemon=True).start()
    except Exception as e:
        log_error(event_type="GUI Error", error_message=str(e))


def _launch_gui():
    """
    Create and run the Tkinter GUI. Runs in a background thread.
    """
    global gui_root, gui_is_open
    try:
        gui_root = tk.Tk()
        gui_root.title("Zoom Room Monitor")
        gui_root.geometry("420x260")
        gui_root.protocol("WM_DELETE_WINDOW", _on_gui_close)

        container = ttk.Frame(gui_root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(container, text="Current Status", font=("Helvetica", 16, "bold"))
        title_label.pack(anchor=tk.W)

        sep = ttk.Separator(container, orient=tk.HORIZONTAL)
        sep.pack(fill=tk.X, pady=(8, 12))

        status_online_var = tk.StringVar(value="Rooms Online: —")
        status_offline_var = tk.StringVar(value="Rooms Offline: —")
        last_updated_var = tk.StringVar(value="Last updated: —")

        online_label = ttk.Label(container, textvariable=status_online_var, font=("Helvetica", 13))
        offline_label = ttk.Label(container, textvariable=status_offline_var, font=("Helvetica", 13))
        updated_label = ttk.Label(container, textvariable=last_updated_var, font=("Helvetica", 10))
        online_label.pack(anchor=tk.W, pady=(0, 4))
        offline_label.pack(anchor=tk.W, pady=(0, 8))
        updated_label.pack(anchor=tk.W)

        def refresh():
            try:
                rooms = get_room(zoom_token)['rooms']
                online_rooms, offline_rooms_count, total_rooms = compute_room_status_counts(rooms)
                status_online_var.set(f"Rooms Online: {online_rooms} / {total_rooms} 🟢")
                status_offline_var.set(f"Rooms Offline: {offline_rooms_count} 🔴")
                last_updated_var.set(f"Last updated: {datetime.now(BOSTON_TZ).strftime('%Y-%m-%d %H:%M:%S %Z')}")
            except Exception as e:
                last_updated_var.set(f"Error: {e}")
            finally:
                gui_root.after(5000, refresh)

        refresh()
        gui_is_open = True
        gui_root.mainloop()
    finally:
        gui_is_open = False


def _on_gui_close():
    """Hide the GUI instead of destroying it, so it can be re-opened quickly."""
    global gui_root, gui_is_open
    try:
        if gui_root is not None:
            gui_root.withdraw()
            gui_is_open = False
    except Exception:
        pass

icon = pystray.Icon(
    "Zoom Room ",
    icon=Image.open("CameraXandTeamsIcon.png"),
    title="Zoom Room Script",
    menu=pystray.Menu(
        pystray.MenuItem("Open", open_gui),
        pystray.MenuItem("Quit", handle_exit)
    )
)

if __name__ == "__main__":

    zoom_token = get_zoom_access_token()
    info = get_room(zoom_token)['rooms']
    try:
        init_online, init_offline, init_total = compute_room_status_counts(info)
        icon.title = f"{init_online} - {init_total} Rooms Online 🟢\n{init_offline} Rooms Offline 🔴"
    except Exception:
        pass
    def worker():
        while(True):
            try:
                is_room_online(zoom_token)  # MAIN SCRIPT RUN
            except Exception as e:
                log_error(event_type="Error", error_message=str(e))
                handle_exit(icon=icon, item=info)
                break
            except KeyboardInterrupt:
                log_error(event_type="Error", error_message=str(KeyboardInterrupt))
                handle_exit(icon=icon, item=info)
                break
            except ConnectionError:
                log_error(event_type="Error", error_message=str(ConnectionError))
                handle_exit(icon=icon, item=info)
                break


    threading.Thread(target=worker, daemon=True).start()
    icon.run()