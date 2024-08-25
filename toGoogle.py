from datetime import datetime, timedelta, timezone
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import pandas as pd
import os
import pytz

# Define the path to your credentials file
TOKEN_PATH = 'token.json'

# Check if the token.json file exists
if not os.path.exists(TOKEN_PATH):
    print(f"Error: '{TOKEN_PATH}' not found.")
    exit(1)

# Load your credentials from a file or environment variable
try:
    creds = Credentials.from_authorized_user_file(TOKEN_PATH)
except Exception as e:
    print(f"Error loading credentials: {e}")
    exit(1)

# Create the Google Calendar API service object
try:
    service = build('calendar', 'v3', credentials=creds)
except Exception as e:
    print(f"Error building the Google Calendar service: {e}")
    exit(1)

# Define the Poland time zone
poland_tz = pytz.timezone('Europe/Warsaw')

# Define the time range for fetching events with timezone-aware datetime objects
timeMin = datetime.now(poland_tz).isoformat()
timeMax = (datetime.now(poland_tz) + timedelta(days=30)).isoformat()

# Fetch the events within the specified time range
try:
    events_result = service.events().list(
        calendarId='primary', timeMin=timeMin, timeMax=timeMax,
        maxResults=10, singleEvents=True,
        orderBy='startTime').execute()
except Exception as e:
    print(f"Error fetching events: {e}")
    exit(1)

events = events_result.get('items', [])

# Load the events from the Excel file
try:
    df = pd.read_excel('outlook_calendar.xlsx')
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit(1)

# Loop through the events in the dataframe and create them in Google Calendar
for index, row in df.iterrows():
    # Convert the start and end times to Poland timezone
    start_local = poland_tz.localize(row['Start'])
    end_local = poland_tz.localize(row['End'])

    event = {
        'summary': row['Subject'],
        'location': row['Location'] if pd.notnull(row['Location']) else '',
        'description': row['Description'],
        'start': {
            'dateTime': start_local.isoformat(),
            'timeZone': 'Europe/Warsaw',
        },
        'end': {
            'dateTime': end_local.isoformat(),
            'timeZone': 'Europe/Warsaw',
        },
    }

    # Check if the event already exists
    existing_event = None
    for e in events:
        if e.get('description') == event['description']:
            existing_event = e
            break

    if not existing_event:
        try:
            created_event = service.events().insert(
                calendarId='primary', body=event).execute()
            print(f"Event created: {created_event['htmlLink']}")
        except Exception as e:
            print(f"Error creating event: {e}")
    else:
        print(f"Event with description '{
              event['description']}' already exists, skipping...")

print("Google Calendar sync completed.")
