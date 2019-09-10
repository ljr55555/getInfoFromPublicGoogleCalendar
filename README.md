# getInfoFromPublicGoogleCalendar
Get events from a Google calendar and create corresponding items in an Exchange calendar

1. [Register the app with Google to use the API](https://developers.google.com/docs/api/quickstart/python)
1. [Install exchangelib](https://github.com/ecederstrand/exchangelib)
1. Copy config.sample to config.py and add personal information
1. Create a ca.crt file with the CA signing key for your Exchange server (or remove the custom adapter if your server cert is signed by a public key)
1. Run getCalendarEvents.py and follow the URL to authorize access to your calendar

