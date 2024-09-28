# The modified original script with the ability to spot already scheduled meetings,
# and schedule meetings for another day if time passed
# Using GUI
# WIth file generation


import requests
import json
from datetime import datetime, timedelta
import openpyxl  # For reading Excel files
from tkinter import Tk
from tkinter.filedialog import askopenfilename


# Zoom API credentials
ACCOUNT_ID = ''
CLIENT_ID = ''
CLIENT_SECRET = ''

# Zoom API base URL
BASE_URL = 'https://api.zoom.us/v2'
AUTH_TOKEN_URL = 'https://zoom.us/oauth/token'


def get_access_token():
    auth_token_url = AUTH_TOKEN_URL
    auth_response = requests.post(auth_token_url, data={
        'grant_type': 'account_credentials',
        'account_id': ACCOUNT_ID,
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET
    })

    if auth_response.status_code == 200:
        return auth_response.json()['access_token']
    else:
        raise Exception(f"Failed to get access token: {auth_response.text}")


def check_existing_meeting(user_identifier, topic, start_time):
    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Convert start_time to ISO 8601 format
    start_time_str = start_time.strftime("%Y-%m-%dT%H:%M:%S")

    # Get user's scheduled meetings
    response = requests.get(f'{BASE_URL}/users/{user_identifier}/meetings', headers=headers, params={
        'type': 'scheduled',
        'page_size': 100  # Adjust as needed
    })

    if response.status_code == 200:
        meetings = response.json().get('meetings', [])
        for meeting in meetings:
            # Convert Zoom's meeting start time (which includes 'Z' for UTC) to match your start_time format
            zoom_meeting_start = meeting['start_time'].replace('Z', '')

            if meeting['topic'] == topic and zoom_meeting_start == start_time_str:
                return True, meeting
    else:
        print(f"Failed to fetch meetings. Status code: {response.status_code}")
        print(f"Response: {response.text}")

    return False, None


def schedule_meeting(user_identifier, topic, start_time, duration, log_file):
    # First, check if the meeting already exists
    meeting_exists, existing_meeting = check_existing_meeting(user_identifier, topic, start_time)

    if meeting_exists:
        log_file.write(f"A meeting with the same topic and start time already exists:")
        # Safely check if 'host_email' exists before printing it
        if 'host_email' in existing_meeting:
            log_file.write(f"Host Email: {existing_meeting['host_email']}")
        else:
            log_file.write("Host Email not available")
        log_file.write(f"Join URL: {existing_meeting['join_url']}")
        log_file.write(f"Meeting ID: {existing_meeting['id']}")
        log_file.write(f"Start Time: {existing_meeting['start_time']}")
        log_file.write("------------------------------------------------------------")
        return

    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Convert start_time to ISO 8601 format
    start_time_str = start_time.strftime("%Y-%m-%dT%H:%M:%S")

    data = {
        'topic': topic,
        'type': 2,  # Scheduled meeting
        'start_time': start_time_str,
        'duration': duration,
        'timezone': 'UTC'  # You can change this to the user's timezone if needed
    }

    response = requests.post(f'{BASE_URL}/users/{user_identifier}/meetings', headers=headers, data=json.dumps(data))

    if response.status_code == 201:
        meeting_info = response.json()
        log_file.write(f"Meeting scheduled successfully!\n")
        log_file.write(f"Host Email: {meeting_info['host_email']}\n")
        log_file.write(f"Join URL: {meeting_info['join_url']}\n")
        log_file.write(f"Meeting ID: {meeting_info['id']}\n")
        log_file.write(f"Start Time: {meeting_info['start_time']}\n")
        log_file.write("------------------------------------------------------------\n")
    else:
        log_file.write(f"Failed to schedule meeting. Status code: {response.status_code}")
        log_file.write(f"Response: {response.text}")


def get_date_from_day_string(day_str, hour, minute):
    """
    This function converts a weekday string (like 'Monday') into a datetime object
    representing the next occurrence of that day in the current month.

    Args:
    day_str (str): The day of the week as a string, e.g., "Monday", "Tuesday".

    Returns:
    datetime: A datetime object representing the next occurrence of that day in the current month.
    """
    # Dictionary to map days of the week to integer values (Monday is 0 and Sunday is 6)
    days_of_week = {
        "Monday": 0,
        "Tuesday": 1,
        "Wednesday": 2,
        "Thursday": 3,
        "Friday": 4,
        "Saturday": 5,
        "Sunday": 6
    }

    # Ensure the day_str is valid and not a number
    if day_str not in days_of_week:
        raise ValueError(f"Invalid day string: {day_str}")

    # Get today's date
    today = datetime.now()

    # Get the current month
    current_month = today.month

    # Get the current year
    current_year = today.year

    # Get the integer value of the target day (e.g., "Monday" -> 0)
    target_day = days_of_week[day_str]

    # Find the difference in days between today and the target day
    current_weekday = today.weekday()  # Returns 0 for Monday, 6 for Sunday
    days_ahead = target_day - current_weekday

    # If the target day is earlier in the week than today, schedule for next week
    if days_ahead < 0:
        days_ahead += 7

    # Check if today is the target day and the time has passed
    if days_ahead == 0:
        target_time_today = today.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if today > target_time_today:
            days_ahead += 7  # Schedule for the next week

    # Calculate the target date
    target_date = today + timedelta(days=days_ahead)

    # Ensure the target date is within the current month
    if target_date.month != current_month:
        target_date = datetime(current_year, current_month, 1) + timedelta(days=(target_day - target_date.weekday()))

    # Return the target date as a datetime object
    return target_date


# Don't forget to make it max rows and columns
def process_excel_data(file_path):
    # Open the file in write mode
    with open('meetings.txt', 'w') as log_file:

        # Load the workbook and select the active sheet
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        # Process rows 2 and 3 (index 1 and 2 in Python)
        for row in sheet.iter_rows(min_row=11, max_row=14, values_only=True):
            # Extract data from the row
            email = row[1]  # Column B: Email
            topic = row[2]  # Column C: Subject
            day1 = row[3]  # Column F: Day 1
            day2 = row[4]  # Column G: Day 2
            hour = row[5]   # Column D: Hour
            minute = row[6]  # Column E: Minute

            # Create datetime objects for both days
            #current_year = datetime.now().year
            #current_month = datetime.now().month

            # Function to create datetime object
            def create_datetime(day):
                return datetime(day.year, day.month, day.day, hour, minute)

            # Schedule meetings for both days
            for day in [day1, day2]:
                if day:  # Check if day is not None or empty
                    date = get_date_from_day_string(day, hour, minute)
                    meeting_datetime = create_datetime(date)
                    schedule_meeting(email, topic, meeting_datetime, 60, log_file)  # Assuming 60 minutes duration


if __name__ == "__main__":
    # Initialize Tkinter and hide the root window
    Tk().withdraw()

    # Use Tkinter file dialog to allow the user to select an Excel file
    print("Please select the Excel file to process.")
    excel_file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    if excel_file_path:
        print(f"Processing file: {excel_file_path}")
        process_excel_data(excel_file_path)
    else:
        print("No file selected. Exiting.")