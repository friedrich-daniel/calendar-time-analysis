#!/usr/bin/env python3
import os
import re
import argparse
import logging
import operator
from datetime import datetime, date, timedelta
import pytz
from icalendar import Calendar
from dateutil.rrule import rruleset, rrulestr, rrule

DESC = 'Simple calendar time analysis by categories within a date range' + os.linesep
DESC += '(e.g.for recording working hours)'

def check_args(args):
    if (args.week is None and args.dstart is None and args.dend is None) or \
       (args.week is not None and args.dstart is None and args.dend is None) or \
       (args.week is None and args.dstart is not None and args.dend is not None):
        return True
    else:
        return False

parser = argparse.ArgumentParser(description=DESC)
parser.add_argument('--file', default=None, help="iCalendar file e.g. *.ics")
parser.add_argument('--week', type=date.fromisoformat, default=None, help="ISO 8601 calendar week e.g. 2024-W01")
parser.add_argument('--dstart', type=date.fromisoformat, default=None, help="Start date e.g. 2024-01-31")
parser.add_argument('--dend'  , type=date.fromisoformat, default=None, help="End date e.g. 2024-01-31")
parser.add_argument('--cregex', default=r"^(\[[A-Z-a-z-0-9-+_\\\/\|]*\])|([A-Z-a-z-0-9-+_\\\/\|]{1,7} - )", help="Rgular expression to extract categories")
#parser.add_argument('--categoryacceptlist', default=[])
args = parser.parse_args() # example for test ["--dstart", "2023-05-11", "--dend", "2023-05-11"]
if not check_args(args):
    parser.error("Either both --dstart and --dend or only --week must be provided.")

if args.file is None:
    # Search for local ics file without glob
    for root, dir, files in os.walk('.'):
      for f in files:
        if f.endswith('.ics'):
          args.file = f
          break

if args.dstart is None or args.dend is None:
    args.dstart = date.fromisoformat( date.today().strftime('%G-W%V') + '-1')
    args.dend   = date.fromisoformat( date.today().strftime('%G-W%V') + '-7')
    cw = date.today().strftime('in %G-W%V')

if args.week is not None:
    args.dstart = date.fromisoformat( args.week.strftime('%G-W%V') + '-1')
    args.dend   = date.fromisoformat( args.week.strftime('%G-W%V') + '-7')
    cw = args.dstart.strftime('in %G-W%V')

CATEGORY_IGNORED = '_Uncategorized_'
category_pattern = re.compile(args.cregex)
out_tzinfo = pytz.timezone('Europe/Berlin') # TODO: ZoneInfo('...') with python 3.9
logging.basicConfig(filename=(args.file + '.log'), filemode='w', level=logging.INFO)

event_dict = {}
total_duration = timedelta()

class Event:
    """Data collecting class for calendar events"""
    def __init__(self, ev_dtstart, ev_duration, ev_summary, ev_year, ev_week, ev_category):
        self.dtstart = ev_dtstart
        self.duration = ev_duration
        self.summary = ev_summary
        self.year = ev_year
        self.week = ev_week
        self.category = ev_category

def add_event(dtstart, duration, summary):
    global event_dict
    global total_duration

    summary = summary[:200]
    if args.dstart <= dtstart.date() and args.dend >= dtstart.date():
        category = CATEGORY_IGNORED
        m = category_pattern.match(summary)
        if m is not None:
            total_duration += duration
            category = ''.join(l for l in m[0] if l.isalnum())
            #category = m[0]
            #if len(args.categoryacceptlist) != 0:
            #    for c in args.categoryacceptlist:
            #        if c.lower() in m[0].lower():
            #            category = c
        # case insensitive category
        for e in event_dict:
            if e.lower() == category.lower():
                category = e
                break
        if category in event_dict:
            event_dict[category][0] = event_dict[category][0] + duration
            event_dict[category][1].append(Event(dtstart, duration, summary, dtstart.isocalendar()[0], dtstart.isocalendar()[1], category))
        else:
            event_dict[category] = [duration, [Event(dtstart, duration, summary, dtstart.isocalendar()[0], dtstart.isocalendar()[1], category)]]


f = open(args.file, 'r', encoding="utf-8")
data = f.read()
# FIXME: icalendar does not support unfold: e.g. multiline exdate
data = data.replace(os.linesep + '\t', "")
calendar = Calendar.from_ical(data)
recurrence_list = []

for component in calendar.walk("VEVENT"):
    if component.get('RECURRENCE-ID'):
        recurrence_list.append(component)

for component in calendar.walk("VEVENT"):
    if not component.get('RECURRENCE-ID'):
        UID = component.get('UID')
        SUMMARY = str(component.get('SUMMARY'))
        dtstart = component.get('DTSTART').dt
        dtend = component.get('DTEND').dt
        DEBUG_INFO = "UID: " + UID + " Summary: " + SUMMARY + " " + str(dtstart) + " " + str(dtend)

        if not isinstance(dtstart, datetime):
            logging.debug('only date without time: %s', DEBUG_INFO)
            continue

        if not component.get('RRULE'):
            add_event(dtstart, dtend - dtstart, SUMMARY)
            continue

        rrule = component.get('RRULE')
        exdates = component.get('EXDATE', [])
        rrset = rruleset()
        rrset.rrule( rrulestr( rrule.to_ical().decode('utf-8'), dtstart = dtstart ) )
        if exdates:
            for exdate in exdates.dts:
                # FIXME: tzinfo replace bugfix since different tzinfo within exdate list
                rrset.exdate( exdate.dt.replace(tzinfo = dtstart.tzinfo))
        for i in list(rrset):
            FOUND = False
            for e in recurrence_list.copy():
                if e.get('UID') == component.get('UID'):
                    dte = e.get('RECURRENCE-ID').dt.astimezone(pytz.timezone('UTC'))
                    recur = i.astimezone(pytz.timezone('UTC'))
                    if recur == dte:
                        FOUND = True
                    if recur.date() == dte.date():
                        # FIXME: relaxed matching due to buggy Microsoft Outlook Export
                        logging.info('relaxed matching: %s', DEBUG_INFO)
                        FOUND = True
                    if FOUND is True:
                        duration = e.get('DTEND').dt - e.get('DTSTART').dt
                        if e.get('SUMMARY'):
                            add_event(e.get('DTSTART').dt, duration, e.get('SUMMARY'))
                        else:
                            add_event(e.get('DTSTART').dt, duration, SUMMARY)
                        recurrence_list.remove(e)
                        break
            if not FOUND:
                add_event(i, dtend - dtstart, SUMMARY)
f.close()

TOTAL_DURATION = '(hours:' + '{:.2f}'.format((total_duration / timedelta(hours=1))) + ')'
if cw is None:
  print('# Calendar times from', args.dstart,'until', args.dend, TOTAL_DURATION )
else:
  print('# Calendar times from', args.dstart,'until', args.dend, cw, TOTAL_DURATION )
for key in sorted(event_dict):
    print('##', key, '- hours:', '{:.2f}'.format((event_dict[key][0] / timedelta(hours=1))))
    sorted_events = sorted(event_dict[key][1], key=operator.attrgetter('dtstart'))
    for e in sorted_events:
        dtstart_converted = e.dtstart.astimezone(out_tzinfo).strftime("%Y-%m-%d %H:%M %Z")
        print(' *', dtstart_converted, e.duration, e.summary)
