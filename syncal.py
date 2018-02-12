#!/usr/bin/env python

########################################
#Author: Dr. Ariadna Blanca Romero
#        Data Scientist
#        ariadana@starignus.com
#        https://github.com/Starignus
#######################################

"""
 By Getting the ics from Outlook calendar, we can get the events and
 duration of the events in the calendar accordingly with the defined categories.

 Information of the API used in this script:
 1) API Reference: https://icalendar.readthedocs.io/en/latest/api.html
 2) ICS.py ICalendar for Humans: http://icspy.readthedocs.io/en/v0.3/
 3) Events API: http://icspy.readthedocs.io/en/latest/api.html#ics.event.Event.begin

"""

from __future__ import print_function
import ics
import arrow
import urllib2
import shutil
import json
import datetime
from pytz import timezone
import pandas as pd
import os
from string import Template


class DeltaTemplate(Template):
    """
    This modifies the format for the time delta function to
    use the delimiter %.
    """
    delimiter = "%"



"""
Global variables
"""
credentials = json.load(open('calendarid.json'))

PATH = os.path.dirname(__file__)
PATH_CSV = os.path.join(PATH, 'time_spend_summary_sd.csv')

URL = credentials["calendarId"][0]["URL"]

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


def update_calendar():
    """
    Function to  updte calendar and download the cal.ics file
    """
    src = urllib2.urlopen(URL)
    with open('cal.ics', 'wb') as dst:
        shutil.copyfileobj(src, dst)


def list_events():
    """
    Function that outputs a list of events on the user's calendar.
    Calls the methods to calculate the time spend in each category.
    """
    # Future Monday
    end = (get_previous_byday('Monday') + datetime.timedelta(days=7))
    # Past Monday
    start = get_previous_byday('Monday')
    #start = arrow.get('2018-02-04T00:00:00Z')
    #end = arrow.get('2018-02-11T00:00:00Z')
    with open('cal.ics', 'rb') as f:
        c = ics.Calendar(f.read().decode('utf-8'))
        events = c.events[start:end]
        if not events:
            print('No upcoming events found.')
        for event in events:
            print(event.begin, event.name, event.duration)
            weekly_hours_time = event_sum_timedelta(event.name, event.duration)
    # Sum the total weekly number of hours
    weekly_hours = weekly_hours_time.total_seconds() / float(3600)
    weekly_hours_formated = strfdelta(weekly_hours_time, "%D days %H:%M:%S")
    data_to_csv(weekly_hours_formated, weekly_hours, start_event=str(start).split('T')[0])


def main():
    """
    Function to run the update calendar and list of events that will be taken into account
    for the start and ending date.
    """
    # If the Calendar has been updated, we can comment it out. Otherwise we leave it.
    update_calendar()
    list_events()


if __name__ == '__main__':
    main()
