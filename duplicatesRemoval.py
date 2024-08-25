from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import os.path
import pickle

# Define the scope for Google Calendar API
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Load credentials
creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If no valid credentials, get new ones
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret_244626337929-041gr73sel2dclu2noffdvbasenk4upc.apps.googleusercontent.com.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

# Build the service
service = build('calendar', 'v3', credentials=creds)

# List all calendars to find the "Work" calendar ID
calendar_list = service.calendarList().list().execute()
work_calendar_id = None
for calendar in calendar_list['items']:
    if calendar['summary'] == 'Work':
        work_calendar_id = calendar['id']
        break

if not work_calendar_id:
    print("Work calendar not found.")
    exit()

print(f"Using calendar ID: {work_calendar_id}")

# Define the time range for the current week and the previous week
now = datetime.now(timezone.utc)
start_of_current_week = now - timedelta(days=now.weekday())
end_of_current_week = start_of_current_week + timedelta(days=6)

start_of_previous_week = start_of_current_week - timedelta(weeks=1)
end_of_previous_week = end_of_current_week - timedelta(weeks=1)


def format_datetime(dt):
    return dt.isoformat()


# Get events from the current week and the previous week
time_min_current = format_datetime(start_of_current_week)
time_max_current = format_datetime(end_of_current_week)

time_min_previous = format_datetime(start_of_previous_week)
time_max_previous = format_datetime(end_of_previous_week)

# Fetch events from the current week
events_result_current = service.events().list(
    calendarId=work_calendar_id,
    timeMin=time_min_current,
    timeMax=time_max_current,
    singleEvents=True,
    orderBy='startTime'
).execute()

# Fetch events from the previous week
events_result_previous = service.events().list(
    calendarId=work_calendar_id,
    timeMin=time_min_previous,
    timeMax=time_max_previous,
    singleEvents=True,
    orderBy='startTime'
).execute()

events_current = events_result_current.get('items', [])
events_previous = events_result_previous.get('items', [])

all_events = events_current + events_previous
print(f"Found {len(all_events)} events in the date ranges.")

# Dictionary to track events by description and their IDs
events_by_description = {}

# Identify duplicates and track them
for event in all_events:
    summary = event.get('summary', 'No Title').strip()
    description = event.get('description', '').strip()
    event_id = event.get('id')
    if description:
        if description in events_by_description:
            events_by_description[description].append(event_id)
        else:
            events_by_description[description] = [event_id]

# Keep one instance and delete the rest
for description, ids in events_by_description.items():
    if len(ids) > 1:
        print(f"Found duplicates for description '{description}':")
        # Keep the first event and delete the rest
        to_delete = ids[1:]
        for event_id in to_delete:
            try:
                service.events().delete(calendarId=work_calendar_id, eventId=event_id).execute()
                print(f"Deleted event with ID: {event_id}")
            except Exception as e:
                print(f"An error occurred while deleting event with ID {
                      event_id}: {e}")

print("Duplicate removal process completed.")
