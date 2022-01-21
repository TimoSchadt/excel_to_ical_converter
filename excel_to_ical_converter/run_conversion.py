import pandas as pd
from tqdm import tqdm
from icalendar import Calendar, Event
from datetime import datetime, time, timedelta

DATE_COL = 1
TIME_COL = 2
LOCATION_COL = -1
TITLE_COL = 3
DURATION_COL = -1

DEFAULT_DURATION_HOURS = 1
DEFAULT_DURATION_MINUTES = 0

data = pd.read_excel("in.xlsx")
cal = Calendar()

for row in tqdm(data.values):
    if type(row[DATE_COL]) == pd._libs.tslibs.timestamps.Timestamp:
        row[DATE_COL] = row[DATE_COL].to_pydatetime()
    if type(row[DATE_COL]) == datetime:
        start_datetime = row[DATE_COL]
        end_datetime = row[DATE_COL]
        if type(row[TIME_COL]) == time:
            start_time = row[TIME_COL]
            start_datetime = start_datetime.replace(hour=start_time.hour)
            start_datetime = start_datetime.replace(minute=start_time.minute)
            if DURATION_COL >= 0 and type(row[DURATION_COL]) == int:
                end_datetime = start_datetime + timedelta(hours=row[DURATION_COL])
            else:
                end_datetime = start_datetime + timedelta(hours=DEFAULT_DURATION_HOURS, minutes=DEFAULT_DURATION_MINUTES)
    elif type(row[DATE_COL]) == str and len(row[DATE_COL].split("-")) == 2:
        parts = row[DATE_COL].split("-")
        start_parts = parts[0].split(".")
        while "" in start_parts:
            start_parts.remove("")
        end_parts = parts[1].split(".")
        for idx, part in enumerate(end_parts):
            if idx >= len(start_parts):
                start_parts.append(part)
        start_datetime = datetime.strptime(f"{start_parts[0]}.{start_parts[1]}.{start_parts[2]}".replace(" ",""), '%d.%m.%Y')
        end_datetime = datetime.strptime(parts[1].replace(" ",""), '%d.%m.%Y')
    else:
        continue
    event = Event()
    event.add('summary', row[TITLE_COL])
    event.add('dtstart', start_datetime)
    event.add('dtend', end_datetime)
    if LOCATION_COL >= 0:
        event.add('location', row[LOCATION_COL])
    #event.add('description', f"Teilnehmer: {row[4]}\nAusbilder: {row[9]}\nTyp: {row[8]}")
    #event.add('description', f"Ausbilder: {row[4]}")
    cal.add_component(event)

f = open('out.ics', 'wb')
f.write(cal.to_ical())
f.close()