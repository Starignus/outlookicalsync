#!/usr/bin/env python

########################################
#Author: Dr. Ariadna Blanca Romero
#        Data Scientist
#        ariadana@starignus.com
#        https://github.com/Starignus
#######################################

from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version
import keyring
import json

credentials = json.load(open('calendarid.json'))


USERNAME = credentials["exchange"][0]["email"]
# "ExchangeCalShopDirect" is the name saved in keyring
PASSWORD = keyring.get_password("ExchangeCalShopDirect", USERNAME)

credentials = Credentials(username=USERNAME, password=PASSWORD)

my_account = Account(primary_smtp_address=USERNAME, credentials=credentials,
                     autodiscover=True, access_type=DELEGATE)

# Build a list of calendar items
tz = EWSTimeZone.timezone('Europe/London')

items_for_2018 = my_account.calendar.filter(start__range=(
    tz.localize(EWSDateTime(2018, 1, 5)),
    tz.localize(EWSDateTime(2018, 1, 10))
))  # Filter by a date range

for item in items_for_2018:
    print(item.start, item.end, item.subject, item.body, item.location)