import datetime
import os.path
import pickle

import win32com.client
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

TIMEZONE = "Asia/Tokyo"
BACK_DAYS = 7
AHEAD_DAYS = 30


class Event:
    MAX_DETAIL_PRINT_LENGTH = 50

    def __init__(self, name, location, start_time, end_time, detail, event_id=None):
        self.name = name + ("" if event_id else " (Outlook)")
        self.location = location
        self.start_time = start_time.replace(tzinfo=None)
        self.end_time = end_time.replace(tzinfo=None)
        self.detail = detail
        self.event_id = event_id

    def __eq__(self, other):
        return (
            isinstance(other, Event)
            and self.name == other.name
            and self.location == other.location
            and self.start_time.isoformat() == other.start_time.isoformat()
            and self.end_time.isoformat() == other.end_time.isoformat()
            and self.detail == other.detail
        )

    def __ne__(self, other):
        return not self == other

    def __str__(self):
        return (
            "name      : " + self.name + "\n"
            "location  : " + self.location + "\n"
            "start_time: " + self.start_time.isoformat() + "\n"
            "end_time  : " + self.end_time.isoformat() + "\n"
            "detail    : "
            + " ".join(self.detail.splitlines())[: Event.MAX_DETAIL_PRINT_LENGTH]
            + "\n"
            "event_id  : " + str(self.event_id) + "\n"
        )


class OutlookCalendar:
    @classmethod
    def connect(cls):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)

        return cls(calendar)

    def __init__(self, calendar):
        self.calendar = calendar

    def get_events(self, start_date, end_date):
        start_date = start_date.date().strftime("%m/%d/%Y 0:00")
        end_date = end_date.date().strftime("%m/%d/%Y 0:00")

        items = self.calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = "True"
        items = items.Restrict(
            "[Start] >= '" + start_date + "' AND [Start] <= '" + end_date + "'"
        )

        events = []

        for item in items:
            event = Event(
                item.subject,
                item.location,
                item.start,
                item.end,
                item.body,
            )
            events.append(event)

            # print(event)

        return events


class GoogleCalendar:
    @classmethod
    def connect(cls):
        SCOPES = ["https://www.googleapis.com/auth/calendar"]

        creds = None

        if os.path.exists("token.pickle"):
            with open("token.pickle", "rb") as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server()

            with open("token.pickle", "wb") as token:
                pickle.dump(creds, token)

        service = build("calendar", "v3", credentials=creds)

        return cls(service)

    def __init__(self, service):
        self.service = service

    def get_events(self, start_date, end_date):
        # start_date = datetime.datetime.combine(
        #    datetime.datetime(2021, 1, 22), datetime.time()
        # )
        # end_date = datetime.datetime.combine(
        #    datetime.datetime(2021, 1, 30), datetime.time()
        # )
        start_date = start_date.isoformat() + "Z"
        end_date = end_date.isoformat() + "Z"

        events_result = (
            self.service.events()
            .list(
                calendarId="primary",
                maxResults=9999,
                timeMin=start_date,
                timeMax=end_date,
                timeZone=TIMEZONE,
            )
            .execute()
        )
        gglevs = events_result.get("items", [])

        events = []

        for gglev in gglevs:
            if gglev["status"] == "cancelled":
                continue

            start = gglev["start"].get("dateTime") or gglev["start"].get("date")
            end = gglev["end"].get("dateTime") or gglev["end"].get("date")

            event = Event(
                gglev["summary"],
                gglev.get("location", ""),
                datetime.datetime.fromisoformat(start),
                datetime.datetime.fromisoformat(end),
                gglev.get("description", ""),
                gglev["id"],
            )
            events.append(event)

            # print(event)

        return events

    def add_event(self, event):
        if (
            event.start_time.time() == "00:00:00"
            and event.end_time.time() == "00:00:00"
        ):
            start = {
                "date": event.start_time.date().isoformat(),
                "timeZone": TIMEZONE,
            }
            end = {
                "date": event.end_time.date().isoformat(),
                "timeZone": TIMEZONE,
            }
        else:
            start = {
                "dateTime": event.start_time.isoformat(),
                "timeZone": TIMEZONE,
            }
            end = {
                "dateTime": event.end_time.isoformat(),
                "timeZone": TIMEZONE,
            }

        body = {
            "summary": event.name,
            "location": event.location,
            "start": start,
            "end": end,
            "description": event.detail,
        }

        event = self.service.events().insert(calendarId="primary", body=body).execute()

    def remove_event(self, event):
        if event.event_id:
            self.service.events().delete(
                calendarId="primary", eventId=event.event_id
            ).execute()


def copy_schedule(back_days, ahead_days):
    today = datetime.datetime.combine(datetime.datetime.now().date(), datetime.time())
    start_date = today - datetime.timedelta(days=back_days)
    end_date = today + datetime.timedelta(days=ahead_days + 1)
    print("copy schedule from {} to {}".format(start_date, end_date))

    outlook_calendar = OutlookCalendar.connect()
    print("connected to Outlook Calendar")

    outlook_events = outlook_calendar.get_events(start_date, end_date)
    print("obtained {} events from Outlook Calendar".format(len(outlook_events)))

    google_calendar = GoogleCalendar.connect()
    print("connected to Google Calendar")

    google_events = google_calendar.get_events(start_date, end_date)
    print("obtained {} events from Google Calendar".format(len(google_events)))

    add_events = []
    remove_events = [x for x in google_events if x.name.endswith(" (Outlook)")]

    for event in outlook_events:
        if event in remove_events:
            remove_events.remove(event)
        else:
            add_events.append(event)

    for event in add_events:
        google_calendar.add_event(event)
    print("added {} events to Google Calendar".format(len(add_events)))

    for event in remove_events:
        google_calendar.remove_event(event)
    print("removed {} events from Google Calendar".format(len(remove_events)))


copy_schedule(BACK_DAYS, AHEAD_DAYS)
