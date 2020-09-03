from __future__ import print_function
import pickle
import os.path
import string
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime, timedelta
import dateutil.parser as parser

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/gmail.modify']

# Function that takes a calendar service and message snippet, parses the email and adds a Calendar Event
def add_cal_event(cal, msg):
    # Take email snippet, take student name and other important information into variables
    words = msg.split(' ')
    student_name = words[7] + ' ' + words[8]
    zoomlink = words[26][:-1]
    timezone = ' -0700'
    
    # Form string we can parse into isoformat time object, format string for start time
    start_raw = ' '.join(words[13:19]) + timezone
    start_date = parser.parse(start_raw)
    start_string = start_date.isoformat()
    start_timeobj = datetime.strptime(start_string,'%Y-%m-%dT%H:%M:%S-07:00')
    
    # Create a new datetime object and add one hour, then format that to isoformat for end time
    one_hour = timedelta(hours=1)
    end_timeobj = start_timeobj + one_hour
    end_string = end_timeobj.isoformat() + '-07:00'

    # Add data into new nested dictionary for new calendar event 
    event = {
        'summary': student_name,
        'creator' : 'self',
        'location': 'Remote via Zoom',
        'description': ('Zoom link: ' + zoomlink),
        'start': {
            'dateTime': start_string,
            'timeZone': 'America/Los_Angeles',
        },
        'end': {
            'dateTime': end_string,
            'timeZone': 'America/Los_Angeles',
        },
        'recurrence': [],
        'attendees': [
            {'email': 'ayang@idtech.com'},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 10},
            ],
        },
        'status': 'confirmed'
    }

    # Add event to the Google Calendar and print to console to notify user! 
    event = cal.events().insert(calendarId='primary', body=event).execute()
    print('Lesson for %s added' %student_name)

# Function that accesses Gmail and goes through a list of email snippets, calling parsing function each time
def get_booking_emails(cal, mail):
    # Start with default search query string, then add to string based on options 
    search_tokens = 'from:iD Tech subject:iD Tech Online - New Lesson Scheduled with '
    search_num = input('Up to how many emails should I try to process? Enter a number: ')
    search_read_emails = raw_input('Search through read emails too? ')
    
    # Check user response, then add 'unread' label if we don't want read emails 
    search_tokens = search_tokens if re.search('[yY].*', search_read_emails) else search_tokens + 'is:unread'
    
    # Get a list of all the email IDs we will parse 
    results = mail.users().messages().list(userId='me', maxResults=search_num, pageToken=None, q=search_tokens).execute()
    messageIDs = results.get('messages', [])

    # Check if we have any messages to process. If there are, get the message's snippet content, then use it to add to calendar
    if not messageIDs:
        print('No messages found.')
    else:
        for ID in messageIDs:
            myId = ID['id']
            email = mail.users().messages().get(userId='me', id=myId).execute()
            mail.users().messages().modify(userId='me', id=myId, body={'removeLabelIds':['UNREAD']}).execute()
            add_cal_event(cal, email['snippet'])

def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Request gcal and gmail services
    gcal = build('calendar', 'v3', credentials=creds)
    gmail = build('gmail', 'v1', credentials=creds)
    # Process email list 
    get_booking_emails(gcal, gmail)
    

if __name__ == '__main__':
    main()
