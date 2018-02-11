#!/usr/bin/env python

########################################
#Author: Dr. Ariadna Blanca Romero
#        Data Scientist
#        ariadana@starignus.com
#        https://github.com/Starignus
#######################################

"""
This script gets the information of Google Calendar (The calendar was sync with the Outlook) from the
last Monday to the following Monday. The categories map the events and get the hours spend in
the week in each of them. It exports the results in a CSV, and each week it will append the
new info for the following week. Since there is a problem with the sync from Outlook and Google it
will be used the syncal.py.

Info:
1. Installing the API client library: https://developers.google.com/google-apps/calendar/quickstart/python#step_2_install_the_google_client_library
2. Options of API: https://developers.google.com/google-apps/calendar/v3/reference/events
3. API javascript: https://script.google.com/d/1RhOVM9boIIWL9rmrmrtiY2fEiNMtCDP0d8t_mou0zW3vTPQkXtmeptuh/edit?splash=yes
"""

from __future__ import print_function
import httplib2
import os
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime
from pytz import timezone
import pandas as pd
from string import Template


class DeltaTemplate(Template):
    """
    This modifies the format for the time delta function to
    use the delimiter %.
    """
    delimiter = "%"

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

"""
Global variables
"""

PATH = os.path.dirname(__file__)
PATH_CSV = os.path.join(PATH, 'time_spend_summary_sd.csv')

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'

# List with day of the week
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday',
            'Friday', 'Saturday', 'Sunday']

# Categories to map
category_bucket = {'PSM': 'Project Specific Meetings',
                   'OM': 'Other Meetings',
                   'PW': 'Project Work',
                   'OW': 'Other Work',
                   'TR': 'Transformation',
                   'IN': 'Innovation',
                   'PD': 'Personal Development',
                   'MT': 'Mentoring',
                   'EH': 'Exempt Hours',
                   'AL': 'Annual Leave',
                   'UN': 'Unavailable',
                   'OT': 'Overtime',
                   'OTH': 'Other',
                   'OTHN': 'Other non mine'
                   }

# Categories to count time
category_hours = {'PSM': datetime.timedelta(0),
                  'OM': datetime.timedelta(0),
                  'PW': datetime.timedelta(0),
                  'OW': datetime.timedelta(0),
                  'TR': datetime.timedelta(0),
                  'IN': datetime.timedelta(0),
                  'PD': datetime.timedelta(0),
                  'MT': datetime.timedelta(0),
                  'EH': datetime.timedelta(0),
                  'AL': datetime.timedelta(0),
                  'UN': datetime.timedelta(0),
                  'OT': datetime.timedelta(0),
                  'OTH': datetime.timedelta(0),
                  'OTHN': datetime.timedelta(0)
                  }


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_previous_byday(dayname, start_date=None):
    """
    Function to get the target date the previous week from
    :param dayname: string: 'Monday', 'Tuesday', etc
    :param start_date: datetime object or None to take the date today.
    :return: datetime object with the target date
    """
    if start_date is None:
        start_date = datetime.date.today()
    day_num = start_date.weekday()
    day_num_target = weekdays.index(dayname)
    days_ago = (7 + day_num - day_num_target) % 7
    if days_ago == 0:
        days_ago = 7
    d = start_date - datetime.timedelta(days=days_ago)
    target_date = datetime.datetime(d.year, d.month, d.day)
    target_date = timezone("Europe/London").localize(target_date)
    return target_date


def event_sum_timedelta(event, time_spend):
    """
    Function to count the total hours spend weekly
    :param event: string with event
    :param time_spend: time delta value
    :return: datetime object with total number of hours
    """
    event_type = event.split(':')
    if event_type[0] in category_hours:
        category_hours[event_type[0]] += time_spend
    else:
        category_hours['OTHN'] += time_spend
    category_hours_copy = category_hours.copy()
    del category_hours_copy['OTHN']
    weekly_hours = sum(category_hours_copy.values(), datetime.timedelta(0))
    return weekly_hours


def strfdelta(tdelta, fmt):
    """
    Function that formats dates for timedelta objects.
    :param tdelta: timedelta object
    :param fmt: string. With format. e.g "%D days %H:%M:%S"
    :return: string. Date in that format
    """
    d = {"D": tdelta.days}
    d["H"], rem = divmod(tdelta.seconds, 3600)
    d["M"], d["S"] = divmod(rem, 60)
    t = DeltaTemplate(fmt)
    return t.substitute(**d)


def data_to_csv(weekly_days_hours, weekly_hours, start_event):
    """
    Function to record the time spend in each category and the total on weekly basis in csv.
    If the file does not exist, it is created. If it exist, the record it is appended
    :param weekly_days_hours: string. total hours with strfdelta format.
    :param weekly_hours: datetime total hours.
    :param start_event: string with start event date (the Monday before)
    :return:
    """
    weekly_record = {category_bucket[name]: strfdelta(val, "%D days %H:%M:%S") for name, val in category_hours.items()}
    weekly_record['weekly_days_hours'] = weekly_days_hours
    weekly_record['weekly_hours'] = weekly_hours
    df_weekly_record = pd.DataFrame(weekly_record, index=[start_event])
    if not os.path.exists(PATH_CSV):
        with open(PATH_CSV, 'w') as f:
            df_weekly_record.to_csv(f)
    else:
        with open(PATH_CSV, 'a') as f:
            df_weekly_record.to_csv(f, header=False)


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of
    events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    calendarId = json.load(open('calendarid.json'))

    future_monday = (get_previous_byday('Monday') + datetime.timedelta(days=7)).isoformat()
    past_monday = get_previous_byday('Monday').isoformat()
    print('Getting the upcoming events')
    eventsResult = service.events().list(
        calendarId=calendarId["calendarId"][0]["id"], timeMin=past_monday,
        timeMax=future_monday, singleEvents=True, orderBy='startTime').execute()
    events = eventsResult.get('items', [])


    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime')
        end = event['end'].get('dateTime')
        if start is not None:
            print('START', start, event['summary'])
            print('END', end, event['summary'])
            dt_start = datetime.datetime.strptime(start, '%Y-%m-%dT%H:%M:%SZ')
            df_end = datetime.datetime.strptime(end, '%Y-%m-%dT%H:%M:%SZ')
            time_spend = df_end - dt_start
            print(time_spend)
            weekly_hours_time = event_sum_timedelta(event['summary'], time_spend)
    # Sum the total weekly number of hours
    weekly_hours = weekly_hours_time.total_seconds() / float(3600)
    weekly_hours_formated = strfdelta(weekly_hours_time, "%D days %H:%M:%S")
    data_to_csv(weekly_hours_formated, weekly_hours, start_event=past_monday.split('T')[0])


if __name__ == '__main__':
    main()
    #print(get_previous_byday('Monday').isoformat())
    #print((get_previous_byday('Monday') + datetime.timedelta(days=7)).isoformat())