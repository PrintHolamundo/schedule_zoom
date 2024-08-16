# Zoom Auto Joiner

## Description
Zoom Auto Joiner automates the process of joining Zoom meetings at scheduled times based on data from an Excel file. It closes Zoom if it is running, restarts it, and uses `pyautogui` to enter the meeting ID and passcode.

## Components
- **Close Zoom**: Terminates the Zoom process using `psutil` if it is running.
- **Join Meeting**: Launches Zoom, navigates the interface, and inputs the meeting ID and passcode using `pyautogui`.
- **Scheduling**: Executes the join meeting function at specified times using `schedule` based on an Excel file.

## Requirements
- **Python**: 3.x and above
- **Python Packages**: Install required packages using:

  ```bash
  pip install -r requirements.txt
  ```

## Usage
1. **Prepare the Excel File**: Ensure your Excel file (`meetings.xlsx`) contains columns `time`, `meeting_id`, and `passcode`.
2. **Run the Script**: Execute the script to keep it running and automatically join meetings as scheduled.


## Author
[David Haro](https://davidharo.net/)
