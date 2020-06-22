# TODO move to utils update doc
from datetime import *
import datetime
from dateutil import tz

# TODO move TIME_ZONE to main user attribute? to ClockifyClient
# TODO get the TimeZone from Google Spreadshhet settings?
# Use https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getspreadsheettimezone
# Set special file to get this. Use credentials
TIME_ZONE = tz.tzoffset('MSK', 10800)  # ('Europe/Moscow') https://www.epochconverter.com/timezones


class DateOutOfRangeException(Exception):
    pass


# TODO update doc, change to ValueException
def week_start_end_datetime(week_number, year=date.today().year):
    '''???'''
    if week_number > 54 or week_number < 1:
        raise DateOutOfRangeException('Week must from 1 to 54, but', week_number)
    if year > 9999 or year < 1:
        raise DateOutOfRangeException('Year must from 1 to 9999, but', year)
    d = "%04d" % (year,) + '-W' + str(week_number)
    tdelta = datetime.timedelta(days=7, microseconds=-1)
    start_datetime = datetime.datetime.strptime(d + '-1', '%G-W%V-%u')
    end_datetime = start_datetime + tdelta
    week_number = datetime.date.isocalendar(end_datetime)
    return (start_datetime, end_datetime)


# TODO update doc how to make docs in line
def format_timedelta_hhmm(td: timedelta) -> str:
    '''shows timedelta in hours and minutes only
    e.g.: 192:13 - 192 hours and 13 minutes
    ===========
    Parameters:
        td = period of time
    Return:
        str in hh:mm format where hh might be as big as needed e.g. 6934:55'''
    minutes, seconds = divmod(td.seconds + td.days * 86400, 60)
    hours, minutes = divmod(minutes, 60)
    return '{:d}:{:02d}'.format(hours, minutes)
