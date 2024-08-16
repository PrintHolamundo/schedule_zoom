import subprocess
import time
import schedule
import pyautogui
import psutil
import pandas as pd
import webbrowser


def close_zoom():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'Zoom.exe':
            proc.kill()


def join_zoom_meeting(meeting_id, passcode):
    close_zoom()
    time.sleep(2)
    subprocess.Popen(['C:\\Users\\davidharo\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'])
    time.sleep(10)
    pyautogui.press('tab', presses=11, interval=0.1)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.write(str(meeting_id))
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.write(str(passcode))
    pyautogui.press('enter')


def join_google_meet(meeting_url):
    webbrowser.open(meeting_url)


def join_meeting(meeting_type, meeting_id, passcode=None):
    if meeting_type == 'zoom':
        join_zoom_meeting(meeting_id, passcode)
    elif meeting_type == 'google_meet':
        join_google_meet(meeting_id)
    else:
        print(f"Unknown meeting type: {meeting_type}")


df = pd.read_excel('meetings.xlsx', dtype={'time': str, 'meeting_id': str, 'passcode': str, 'type': str})

for _, row in df.iterrows():
    schedule.every().day.at(row['time']).do(join_meeting, row['type'], row['meeting_id'], row.get('passcode'))

while True:
    schedule.run_pending()
    time.sleep(1)
