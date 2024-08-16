# Zoom Auto Joiner

## Description

Zoom Auto Joiner automates the process of joining scheduled meetings for both Zoom and Google Meet based on data from an
Excel file. It handles Zoom by closing any running instance, restarting it, and using `pyautogui` to enter the meeting
ID and passcode. For Google Meet, it opens the meeting URL directly in your default web browser.

## Components

- **Close Zoom**: Terminates the Zoom process using `psutil` if it is running.
- **Join Zoom Meeting**: Launches Zoom, navigates the interface, and inputs the meeting ID and passcode
  using `pyautogui`.
- **Join Google Meet**: Opens the Google Meet URL in your default web browser.
- **Scheduling**: Executes the join meeting function at specified times using `schedule` based on an Excel file.

## Requirements

- **Python**: 3.x and above
- **Python Packages**: Install required packages using:

  ```bash
  pip install -r requirements.txt
  ```

## Usage

1. **Prepare the Excel File**: Ensure your Excel file (meetings.xlsx) contains columns time, meeting_id, passcode, and
   type. The type column should specify zoom or google_meet
2. **Run the Script**: Execute the script to keep it running and automatically join meetings as scheduled.

## Author

[David Haro](https://davidharo.net/)
