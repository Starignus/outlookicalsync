## Log of time spends over a week using info from Outlook calendar through Google API or iCalendar API. 


In this Git will be different script to make attempts to get events from Outlook Calendar (Exchange).  
The idea is to track the time spent in each task over a week accordingly with defined categories where the time has bucket. E.g.

```
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
```
                   
At the end of the week, you can edit the starting text of each event with the corresponding category.  
This script gets the information from Google Calendar or iCalendar  (a .ics file is downloaded locally) 
from the last Monday to the following Monday. The categories map the events and get the hours spend in 
the week in each of them. It exports the results in a CSV, and each week it will append the new info 
for the following week. Since there is a problem with the sync between Outlook and Google, it will 
be used the ```syncal.py``` which uses the iCalendar method (downloads a cal.ics in your local path).


The are other two script that tries to use the exchange libraries, but credential authentication was not successful: 
``CalendarExhange.py`` and ``CalendarExcahngelib.py``.

## syncal.py

By Getting the ics from Outlook calendar, we can get the events and duration of the events in the calendar accordingly 
with the defined categories.

 1) [API Reference](https://icalendar.readthedocs.io/en/latest/api.html)
 2) [ICS.py ICalendar for Humans](http://icspy.readthedocs.io/en/v0.3/)
 3) [Events API](http://icspy.readthedocs.io/en/latest/api.html#ics.event.Event.begin)

####  Requirements :

1. ics
2. urllib2
3. shutil
4.  json
5. datetime
6. timezone
7. pandas
8. os
9. Template

## CalendarGoogleAPI

1 . Sync Office 365 calendar (Outlook) to Google calendar. 
Follow the instructions in this [link](https://webapps.stackexchange.com/questions/89473/sync-office-365-calendar-to-google-calendar). 

2 . Turn on the Google Calendar API and install [the Google Client Library](https://developers.google.com/google-apps/calendar/quickstart/python#step_2_install_the_google_client_library). In this link there is a python sample code which we used to in our code. 

3 . More information about the [API options](https://developers.google.com/google-apps/calendar/v3/reference/events).

More information about the [Class Calendar events API](https://developers.google.com/apps-script/reference/calendar/calendar-event-series).
You can play with the API using java script in the next [link](https://script.google.com/home).

####  Requirements :

1. Python 2.7.x.
2. httplib2.
3. discovery.
4. oauth2client.
5. timezone.
6. Template.
7. argparse.
8. os.
9. ics.

## Credentials file

Construct a json file with the credentials from Outlook `calendarid.json`, they can be used for both,
`syncal.py`  and `CalendarGoogleAPI.py`:

```json
{"calendarId": [
  {"id": "as7d7sd7f7dfgh8fg9f@import.calendar.google.com",
  "URL": "https://outlook.office365.com/owa/calendar/sdsfd8f88d99fdsf00s0d9s99sf@shopdirect.com/sdfd8f89d90s0dsdfsdff0s88dfs/calendar.ics"
  }],
"exchange": [
    {"email": "xxxxxx@email.com",
    "URL": "https://outlook.office365.com/EWS/Exchange.asmx"}
  ]}
```


