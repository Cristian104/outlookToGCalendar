import win32com.client
import pandas as pd
from datetime import datetime, timedelta
import pytz
import hashlib


def get_current_week_range():
    today = datetime.now(pytz.timezone('Europe/Warsaw'))
    start_of_week = today - timedelta(days=today.weekday())
    end_of_week = start_of_week + timedelta(days=6)
    return start_of_week, end_of_week


def generate_unique_key(subject, start):
    key_string = f"{subject}-{start.strftime('%Y-%m-%d %H:%M:%S')}"
    return hashlib.md5(key_string.encode()).hexdigest()


def get_outlook_calendar_events():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)
    items = calendar.Items

    start_of_week, end_of_week = get_current_week_range()
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
        start_of_week.strftime("%m/%d/%Y %I:%M %p"),
        end_of_week.strftime("%m/%d/%Y %I:%M %p")
    )
    restricted_items = items.Restrict(restriction)

    events = []
    poland_tz = pytz.timezone('Europe/Warsaw')

    for item in restricted_items:
        try:
            start = item.Start
            end = item.End
            if isinstance(start, str):
                start = datetime.strptime(start, '%m/%d/%Y %I:%M %p')
            if isinstance(end, str):
                end = datetime.strptime(end, '%m/%d/%Y %I:%M %p')

            if start.tzinfo is None:
                start = poland_tz.localize(start)
            if end.tzinfo is None:
                end = poland_tz.localize(end)

            start = start.replace(tzinfo=None, second=0, microsecond=0)
            end = end.replace(tzinfo=None, second=0, microsecond=0)

            unique_key = generate_unique_key(item.Subject, start)

            event = {
                "Subject": item.Subject,
                "Start": start,
                "End": end,
                "Location": item.Location,
                "Description": unique_key
            }
            events.append(event)
        except Exception as e:
            print(f"An event named '{item.Subject}' has an error: {
                  e} and will be skipped.")

    return pd.DataFrame(events)


events_df = get_outlook_calendar_events()
events_df['Start'] = pd.to_datetime(
    events_df['Start'].dt.strftime('%Y-%m-%d %H:%M'), errors='coerce')
events_df['End'] = pd.to_datetime(
    events_df['End'].dt.strftime('%Y-%m-%d %H:%M'), errors='coerce')
events_df.dropna(subset=['Start', 'End'], inplace=True)

try:
    events_df.to_excel("outlook_calendar.xlsx", index=False)
    print("Events successfully exported to outlook_calendar.xlsx")
except Exception as e:
    print(f"Error exporting to Excel: {e}")
