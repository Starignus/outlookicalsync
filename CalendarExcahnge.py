#!/usr/bin/env python

########################################
#Author: Dr. Ariadna Blanca Romero
#        Data Scientist
#        ariadana@starignus.com
#        https://github.com/Starignus
#######################################

"""
 Credentials
"""

import keyring
from pytz import timezone
from datetime import datetime
from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection
import json

credentials = json.load(open('calendarid.json'))


USERNAME = credentials["exchange"][0]["email"]
PASSWORD = keyring.get_password("ExchangeCalShopDirect", USERNAME)
URL = credentials["exchange"][0]["URL"]

print(PASSWORD)
# Set up the connection to Exchange
connection = ExchangeNTLMAuthConnection(url=URL,
                                        username="ariadnablanca.romero@shopdirect.com",
                                        password=PASSWORD)

service = Exchange2010Service(connection)

my_calendar = service.calendar()

events = my_calendar.list_events(
    start=timezone("Europe/London").localize(datetime(2018, 01, 5)),
    end=timezone("Europe/London").localize(datetime(2018, 01, 10)),
    details=True
)

for event in events:
    print "{start} {stop} - {subject}".format(
        start=event.start,
        stop=event.end,
        subject=event.subject
    )


