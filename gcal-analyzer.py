from __future__ import print_function
import datetime
import os.path
import argparse
from collections import Counter
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
INTERNAL_DOMAIN = 'dlc.link'

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
    return (end_time - start_time).total_seconds() / 60  # Duration in minutes

def read_titles_from_file(file_path):
    with open(file_path, 'r') as file:
        titles = [line.strip() for line in file.readlines()]
    return titles

def is_internal_meeting(attendees):
    if not attendees:
        return False
    return all(INTERNAL_DOMAIN in attendee['email'] for attendee in attendees if 'email' in attendee)

def main():
    parser = argparse.ArgumentParser(description="Google Calendar Analyzer")
    parser.add_argument('email', type=str, help='Email address to analyze calendar events for')
    args = parser.parse_args()
    email = args.email

    creds = authenticate_google_calendar()
    service = build('calendar', 'v3', credentials=creds)
    
    # Define Q2 2024 start and end times
    q2_start = '2024-04-01T00:00:00Z'
    q2_end = '2024-06-30T23:59:59Z'
    
    events = fetch_calendar_events(service, email, q2_start, q2_end)
    
    event_count = Counter(event.get('summary', 'No Title') for event in events)
    recurring_titles = [title for title, count in event_count.items() if count > 5]
    skip_titles_file = 'skip_titles.txt'
    forced_inclusion_file = 'forced_inclusion.txt'
    external_skip_titles = read_titles_from_file(skip_titles_file)
    forced_inclusion_titles = read_titles_from_file(forced_inclusion_file)

    total_duration = 0
    aggregated_data = []
    skipped_meetings = {
        'high_recurrence': set(),
        'external_skip': set(),
        'internal_meeting': set()
    }

    for event in events:
        event_name = event.get('summary', 'No Title')
        attendees = event.get('attendees', [])

        if event_name in forced_inclusion_titles:
            start_time = event['start'].get('dateTime', event['start'].get('date'))
            end_time = event['end'].get('dateTime', event['end'].get('date'))
            duration = calculate_event_duration(start_time, end_time)
            total_duration += duration

            aggregated_data.append({
                'name': event_name,
                'duration_minutes': duration
            })
            continue

        if event_name in recurring_titles:
            skipped_meetings['high_recurrence'].add(event_name)
            continue
        if event_name in external_skip_titles:
            skipped_meetings['external_skip'].add(event_name)
            continue
        if is_internal_meeting(attendees):
            skipped_meetings['internal_meeting'].add(event_name)
            continue

        start_time = event['start'].get('dateTime', event['start'].get('date'))
        end_time = event['end'].get('dateTime', event['end'].get('date'))
        duration = calculate_event_duration(start_time, end_time)
        total_duration += duration

        aggregated_data.append({
            'name': event_name,
            'duration_minutes': duration
        })

    print("\nSkipped meeting titles due to high recurrence:")
    for title in skipped_meetings['high_recurrence']:
        print(f"{title}")

    print("\nSkipped meeting titles due to external file input:")
    for title in skipped_meetings['external_skip']:
        print(f"{title}")

    print("\nSkipped meeting titles due to being internal meetings:")
    for title in skipped_meetings['internal_meeting']:
        print(f"{title}")

    for data in aggregated_data:
        print(f"{data['name']} ({data['duration_minutes']}m)")

    # Print total time spent in meetings and review period at the bottom
    print(f"\nReviewing calendar events from {q2_start} to {q2_end}")
    print(f"Total time spent in meetings: {total_duration / 60:.2f} hours")

if __name__ == '__main__':
    main()
