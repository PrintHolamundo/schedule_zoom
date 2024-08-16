import subprocess
import time
import schedule
import pyautogui
import psutil
import pandas as pd

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

df = pd.read_excel('meetings.xlsx', dtype={'time': str, 'meeting_id': str, 'passcode': str})

for _, row in df.iterrows():
    schedule.every().day.at(row['time']).do(join_zoom_meeting, row['meeting_id'], row['passcode'])

while True:
    schedule.run_pending()
    time.sleep(1)

