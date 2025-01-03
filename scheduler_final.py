import requests
import json
from datetime import datetime, timedelta
import openpyxl  # For reading Excel files
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import io


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
        log_file.write(f"A meeting with the same topic and start time already exists: {existing_meeting['start_time']}\n")
        if 'host_email' in existing_meeting:
            log_file.write(f"Host Email: {existing_meeting['host_email']}\n")
        log_file.write(f"Join URL: {existing_meeting['join_url']}\n")
        log_file.write(f"Meeting ID: {existing_meeting['id']}\n")
        log_file.write("------------------------------------------------------------\n")
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
        # Extract meeting info
        meeting_url = meeting_info['join_url']
        meeting_id = meeting_info['id']
        host_email = meeting_info.get('host_email', 'N/A')
        meeting_time = meeting_info['start_time']
        formatted_time = start_time.strftime("%Hh%M")  # Format time as "HourhMinutes"

        # Calculate the time plus 2 hours
        time_plus_2_hours = (start_time + timedelta(hours=1)).strftime("%Hh%M")

        # Write the formatted text to the log file
        log_file.write(f"-----------------Host Email: {host_email}-----------------\n")  # Include host email at the top
        log_file.write(f"-----------------Meeting Time: {meeting_time}-----------------\n")
        log_file.write(u"\n🔴 Please use the link below for today's class\n\n")
        log_file.write(u"If anyone is going to be late or absent, please let the group know well in advance, thank you!\n\n")
        log_file.write(u"WHEN\n\n")
        log_file.write(f"🚦 {formatted_time} GMT\n")
        log_file.write(f"🚦 {time_plus_2_hours} Europe\n\n")
        log_file.write(u"WHERE\n\n")
        log_file.write(f"CLICK HERE\n{meeting_url}\n\n")
        log_file.write(f"Meeting: {meeting_id}\n\n")
        log_file.write(u"HOW\n\n")
        log_file.write(u"Si c'est votre première fois à utiliser l'application gratuite ZOOM, veuille cliquer sur le lien zoom ci-dessus "
                       u"et téléchargez l’application ZOOM avant le cours. A l’heure du cours, cliquez de nouveau sur le lien zoom "
                       u"pour lancer l’application zoom et s’il vous demande le numéro de réunion, veuillez entrer le numéro unique de "
                       u"réunion à 11 chiffres afin de pouvoir vous connecter à la classe en direct. Lorsque vous vous connectez, "
                       u"assurez-vous que vos microphones sont activés et changez vos noms dans le ZOOM en vos prénoms et nom de famille. "
                       u"N’activez pas la vidéo - c’est des cours en audio. On se parle en classe !\n")
        log_file.write(u"------------------------------------------------------------\n")
    else:
        log_file.write(f"Failed to schedule meeting. Status code: {response.status_code}\n")
        log_file.write(f"Response: {response.text}\n")


# Updated function to calculate the next weekday occurrence
# Updated function to calculate the next weekday occurrence, allowing today if it's the target weekday
# and the meeting time has not yet passed
def get_next_weekday(start_date, weekday_name, include_today):
    weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    weekday_target = weekdays.index(weekday_name)

    # Calculate days ahead to the target weekday
    days_ahead = weekday_target - start_date.weekday()

    # Adjust if necessary
    if include_today and days_ahead == 0:
        # Check if the meeting can still be scheduled today (time has not passed)
        return start_date
    elif days_ahead <= 0:  # if the target day has passed for this week, schedule for the next week
        days_ahead += 7

    # Return the next occurrence of the weekday
    return start_date + timedelta(days=days_ahead)


# Use the new get_next_weekday logic for calculating meeting days
def process_excel_data(file_path):
    # Open the file in write mode with UTF-8 encoding
    with io.open('meetings.txt', 'w', encoding='utf-8') as log_file:
        # Load the workbook and select the active sheet
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        # Process rows 2 and 3 (index 1 and 2 in Python)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Extract data from the row
            email = row[1]  # Column B: Email
            topic = row[2]  # Column C: Subject
            day1 = row[3]  # Column F: Day 1
            day2 = row[4]  # Column G: Day 2
            hour = row[5]   # Column D: Hour
            minute = row[6]  # Column E: Minute

            # Create datetime objects for both days
            def create_datetime(day):
                return datetime(day.year, day.month, day.day, hour, minute)

            # Schedule meetings for both days
            for day in [day1, day2]:
                if day:  # Check if day is not None or empty
                    today = datetime.now()
                    include_today = True  # If you want to consider today
                    next_day = get_next_weekday(today, day, include_today)
                    meeting_datetime = create_datetime(next_day)
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
