from __future__ import print_function
import datetime
import os.path
from collections import Counter
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def authenticate_google_calendar():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=8080)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def fetch_calendar_events(service, calendar_id, start_time, end_time):
    events_result = service.events().list(
        calendarId=calendar_id, timeMin=start_time, timeMax=end_time,
        singleEvents=True, orderBy='startTime').execute()
    return events_result.get('items', [])

def calculate_event_duration(start, end):
    start_time = datetime.datetime.fromisoformat(start.replace('Z', '+00:00'))
    end_time = datetime.datetime.fromisoformat(end.replace('Z', '+00:00'))
    return (end_time - start_time).total_seconds() / 3600  # Duration in hours

def main():
    email = input("Enter the email address: ")  # Prompt user for the email
    
    creds = authenticate_google_calendar()
    service = build('calendar', 'v3', credentials=creds)
    
    # Define Q2 2024 start and end times
    q2_start = '2024-04-01T00:00:00Z'
    q2_end = '2024-06-30T23:59:59Z'
    
    print(f"Reviewing calendar events from {q2_start} to {q2_end}")
    
    events = fetch_calendar_events(service, email, q2_start, q2_end)
    
    event_count = Counter(event.get('summary', 'No Title') for event in events)
    skip_titles = [title for title, count in event_count.items() if count > 5]

    total_duration = 0
    aggregated_data = []

    for event in events:
        event_name = event.get('summary', 'No Title')
        if event_name in skip_titles:
            continue

        start_time = event['start'].get('dateTime', event['start'].get('date'))
        end_time = event['end'].get('dateTime', event['end'].get('date'))
        duration = calculate_event_duration(start_time, end_time)
        total_duration += duration

        aggregated_data.append({
            'name': event_name,
            'duration_hours': duration
        })

    print("Skipped meeting titles:")
    for title in skip_titles:
        print(f" - {title}")

    print(f"\nTotal time spent in meetings: {total_duration:.2f} hours\n")

    for data in aggregated_data:
        print(f"Event: {data['name']}, Duration: {data['duration_hours']} hours")

if __name__ == '__main__':
    main()
