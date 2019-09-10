# Requirements for Google API
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Requirements for ExchangeLib
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, FaultTolerance, \
    Configuration, NTLM, GSSAPI, SSPI, Build, Version, CalendarItem, EWSDateTime

from exchangelib.folders import Calendar
from exchangelib.items import MeetingRequest, MeetingCancellation, SEND_TO_ALL_AND_SAVE_COPY
from exchangelib.protocol import BaseProtocol

from urllib.parse import urlparse
import requests.adapters

# Used to decrypt password from config file
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode

# Misc imports
from datetime import datetime
import dateutil.parser as parser
import pytz

from config import strUsernameCrypted, strPasswordCrypted, strEWSHost, strPrimarySMTP, dictCalendars, iMaxGoogleResults, iMaxExchangeResults, strKey

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

# Set our time zone
local_tz = pytz.timezone('US/Eastern')

# This needs to be the key used to stash the username and password values stored in config.py
f = Fernet(strKey)
strUsername = f.decrypt(strUsernameCrypted)
strUsername = strUsername.decode("utf-8")

strPassword = f.decrypt(strPasswordCrypted)
strPassword = strPassword.decode("utf-8")

def createExchangeItem(objExchangeAccount, strTitle, strLocn, strStartDate, strEndDate, strInviteeSMTP):
    print("Creating item {} which starts on {} and ends at {}".format(strTitle, strStartDate,strEndDate))
    objStartDate = parser.parse(strStartDate)
    objEndDate = parser.parse(strEndDate)
    
    item = CalendarItem(
        account=objExchangeAccount,
        folder=objExchangeAccount.calendar,
        start=objExchangeAccount.default_timezone.localize(EWSDateTime(objStartDate.year, objStartDate.month, objStartDate.day, objStartDate.hour, objStartDate.minute)),
        end=objExchangeAccount.default_timezone.localize(EWSDateTime(objEndDate.year,objEndDate.month,objEndDate.day,objEndDate.hour,objEndDate.minute)),
        subject=strTitle,
        reminder_minutes_before_start=30,
        reminder_is_set=True,
        location=strLocn,
        body="",
        required_attendees=[strInviteeSMTP]
    )
    item.save(send_meeting_invitations=SEND_TO_ALL_AND_SAVE_COPY )

def utc_to_local(utc_dt):
    local_dt = utc_dt.replace(tzinfo=pytz.utc).astimezone(local_tz)
    return local_tz.normalize(local_dt) 

def main():
    class RootCAAdapter(requests.adapters.HTTPAdapter):
        # An HTTP adapter that uses a custom root CA certificate at a hard coded location
        def cert_verify(self, conn, url, verify, cert):
            cert_file = {
                'exchange01.rushworth.us': './ca.crt'
            }[urlparse(url).hostname]
            super(RootCAAdapter, self).cert_verify(conn=conn, url=url, verify=cert_file, cert=cert)

    #Use this SSL adapter class instead of the default
    BaseProtocol.HTTP_ADAPTER_CLS = RootCAAdapter

    # Get Exchange calendar events and save to dictEvents 
    dictEvents = {}

    credentials = Credentials(username=strUsername, password=strPassword)
    config = Configuration(server=strEWSHost, credentials=credentials)
    account = Account(primary_smtp_address=strPrimarySMTP, config=config,
                    autodiscover=False, access_type=DELEGATE)

    for item in account.calendar.all().order_by('-start')[:iMaxExchangeResults]:
        if item.start:
            objEventStartTime = parser.parse(str(item.start))
            objEventStartTime = utc_to_local(objEventStartTime)

            strEventKey = "{}{:02d}-{:02d}-{:02d}".format(str(item.subject), int(objEventStartTime.year), int(objEventStartTime.month), int(objEventStartTime.day))
            strEventKey = strEventKey.replace(" ","")
            dictEvents[strEventKey]=1

    # Get Google calendar events & create if not already in dictEvents
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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    for strCalendarName, strCalendarID in dictCalendars.items():
        # Call the Calendar API
        now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
        events_result = service.events().list(calendarId=strCalendarID, timeMin=now,
                                            maxResults=iMaxGoogleResults, singleEvents=True,
                                            orderBy='startTime').execute()
        events = events_result.get('items', [])

        for event in events:
            dateStart = event['start'].get('dateTime', event['start'].get('date'))
            dateEnd = event['end'].get('dateTime', event['end'].get('date'))
            strLocation = event.get('location')
            
            strSummary = "{}: {}".format(strCalendarName, event.get('summary'))
            strThisEventKey = strSummary + (str(dateStart).split('T'))[0]
            strThisEventKey = strThisEventKey.replace(" ","")

            if strThisEventKey not in dictEvents:
                createExchangeItem(account, strSummary, strLocation, dateStart, dateEnd, 'lisa@rushworth.us')
            else:
                print("The event {} on {} already exists in the calendar.".format(strThisEventKey, str(dateStart)))

if __name__ == '__main__':
    main()
